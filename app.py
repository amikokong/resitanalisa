"""
app.py  — Resit Analisa (Streamlit + Gemini + Google Drive OAuth)
-----------------------------------------------------------------
Aliran kerja automatik:
  - Log masuk Google sekali (OAuth)
  - App BACA fail induk terus dari Drive anda
  - Upload resit baru -> AI ekstrak -> anti-pendua
  - App TULIS fail induk terus ke Drive (tiada download/upload manual)

Skop Drive: drive.file -> app HANYA nampak fail yang ia sendiri cipta.
            (Tidak boleh intai fail lain dalam Drive anda.)

Jalankan:
    pip install -r requirements.txt
    streamlit run app.py

Secrets (Streamlit Cloud: Settings > Secrets):
    GEMINI_API_KEY      = "AIza..."
    GOOGLE_CLIENT_ID    = "xxxx.apps.googleusercontent.com"
    GOOGLE_CLIENT_SECRET= "GOCSPX-xxxx"
    REDIRECT_URI        = "https://resitanalisa.streamlit.app"   # URL app anda

Dibangunkan oleh: Sulaiman Osman  (sulaimanosman03@gmail.com)
"""

import json
import time
from io import BytesIO

import streamlit as st
import pandas as pd
from PIL import Image
from google import genai
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

MODEL_NAME = "gemini-2.5-flash"
MONEY_FORMAT = "#,##0.00"
BOLD = Font(bold=True)

AUTHOR_NAME = "Sulaiman Osman"
AUTHOR_EMAIL = "sulaimanosman03@gmail.com"

MASTER_FILENAME = "fail_induk_perakaunan.xlsx"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SCOPES = ["https://www.googleapis.com/auth/drive.file"]

FIELDS = [
    "transaction_date", "vendor_name", "receipt_no", "description",
    "amount", "debit_account", "credit_account", "currency",
]


# ==================================================
# HELPER EXCEL
# ==================================================
def write_headers(ws, headers):
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.font = BOLD
        cell.alignment = Alignment(horizontal="center")


def apply_money_format(ws, columns, start_row=2):
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            if cell.column_letter in columns and isinstance(cell.value, (int, float)):
                cell.number_format = MONEY_FORMAT


def autofit(wb, width=22):
    for ws in wb.worksheets:
        for column in ws.columns:
            ws.column_dimensions[column[0].column_letter].width = width


def add_footer(ws):
    last = ws.max_row + 2
    ws.cell(row=last, column=1, value=f"Dibangunkan oleh: {AUTHOR_NAME}")
    link = ws.cell(row=last + 1, column=1, value=AUTHOR_EMAIL)
    link.hyperlink = f"mailto:{AUTHOR_EMAIL}"
    link.font = Font(color="0563C1", underline="single")


# ==================================================
# ANTI-PENDUA
# ==================================================
def _vendor(t):
    return str(t.get("vendor_name", "")).strip().lower()


def classify(txn, existing):
    """new = tambah · exact = pendua tepat (langkau) · maybe = sahkan manusia."""
    receipt_no = str(txn.get("receipt_no", "")).strip()
    if receipt_no:  # Lapisan 1: nombor resit
        for e in existing:
            if _vendor(e) == _vendor(txn) and str(e.get("receipt_no", "")).strip() == receipt_no:
                return "exact"
        return "new"
    # Lapisan 2: tiada nombor resit
    amt = round(float(txn.get("amount", 0) or 0), 2)
    date = str(txn.get("transaction_date", "")).strip()
    for e in existing:
        if (str(e.get("transaction_date", "")).strip() == date
                and _vendor(e) == _vendor(txn)
                and round(float(e.get("amount", 0) or 0), 2) == amt):
            return "maybe"
    return "new"


# ==================================================
# AI
# ==================================================
def extract_receipt_data(image, max_retries=4):
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    prompt = """
    Baca resit ini dan pulangkan HANYA objek JSON (tiada teks lain, tiada markdown).
    Cari nombor resit / invois / bil untuk 'receipt_no'. Jika tiada, letak "".
    Format:
    {
      "transaction_date": "YYYY-MM-DD",
      "vendor_name": "...",
      "receipt_no": "...",
      "description": "...",
      "amount": 0.00,
      "debit_account": "...",
      "credit_account": "Cash",
      "currency": "MYR"
    }
    """
    last_err = None
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=[prompt, image],
                config={"response_mime_type": "application/json"},
            )
            return json.loads(response.text)
        except Exception as e:
            last_err = e
            msg = str(e).lower()
            # Cuba semula HANYA untuk ralat sementara (server sibuk)
            if any(k in msg for k in ["503", "unavailable", "overload", "429", "rate"]):
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)  # tunggu 1, 2, 4 saat (backoff)
                    continue
            raise  # ralat lain (cth JSON rosak) -> terus naikkan
    raise last_err


# ==================================================
# EXCEL: baca & bina
# ==================================================
def read_master_bytes(buf):
    wb = load_workbook(buf)
    if "Transactions" not in wb.sheetnames:
        return []
    ws = wb["Transactions"]
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        out.append({FIELDS[i]: (row[i] if i < len(row) else "") for i in range(len(FIELDS))})
    return out


def build_master_excel(transactions):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Transactions"
    write_headers(ws1, ["Date", "Vendor", "Receipt No", "Description", "Amount",
                        "Debit Account", "Credit Account", "Currency"])
    for t in transactions:
        ws1.append([
            t.get("transaction_date", ""), t.get("vendor_name", ""),
            t.get("receipt_no", ""), t.get("description", ""),
            float(t.get("amount", 0) or 0), t.get("debit_account", ""),
            t.get("credit_account", ""), t.get("currency", "MYR"),
        ])
    apply_money_format(ws1, ["E"])

    ws2 = wb.create_sheet("Journal Entry")
    write_headers(ws2, ["Date", "Description", "Account", "Debit", "Credit"])
    for t in transactions:
        amt = float(t.get("amount", 0) or 0)
        ws2.append([t.get("transaction_date", ""), t.get("description", ""),
                    t.get("debit_account", ""), amt, None])
        ws2.append([t.get("transaction_date", ""), t.get("description", ""),
                    t.get("credit_account", ""), None, amt])
    apply_money_format(ws2, ["D", "E"])

    df = pd.DataFrame(transactions)
    if not df.empty:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
        df["transaction_date"] = pd.to_datetime(df["transaction_date"], errors="coerce")
        df["month"] = df["transaction_date"].dt.strftime("%Y-%m")

    ws3 = wb.create_sheet("Monthly Summary")
    write_headers(ws3, ["Month", "Total Expense (RM)", "Bil. Transaksi"])
    if not df.empty:
        monthly = df.groupby("month").agg(total=("amount", "sum"),
                                          count=("amount", "count")).reset_index()
        for _, r in monthly.iterrows():
            ws3.append([r["month"], float(r["total"]), int(r["count"])])
    apply_money_format(ws3, ["B"])
    add_footer(ws3)

    ws4 = wb.create_sheet("By Account")
    write_headers(ws4, ["Debit Account", "Total (RM)"])
    if not df.empty:
        by_acc = df.groupby("debit_account")["amount"].sum().reset_index()
        for _, r in by_acc.iterrows():
            ws4.append([r["debit_account"], float(r["amount"])])
    apply_money_format(ws4, ["B"])

    autofit(wb)
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ==================================================
# GOOGLE DRIVE (OAuth)
# ==================================================
def get_flow():
    return Flow.from_client_config(
        {
            "web": {
                "client_id": st.secrets["GOOGLE_CLIENT_ID"],
                "client_secret": st.secrets["GOOGLE_CLIENT_SECRET"],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": [st.secrets["REDIRECT_URI"]],
            }
        },
        scopes=SCOPES,
        redirect_uri=st.secrets["REDIRECT_URI"],
        autogenerate_code_verifier=False,  # matikan PKCE (elak ralat "Missing code verifier")
    )


def drive_service(creds):
    return build("drive", "v3", credentials=creds)


def find_master(service):
    """Cari fail induk yang app pernah cipta (skop drive.file)."""
    q = f"name='{MASTER_FILENAME}' and trashed=false"
    res = service.files().list(q=q, spaces="drive", fields="files(id,name)").execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


def download_master(service, file_id):
    req = service.files().get_media(fileId=file_id)
    buf = BytesIO()
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


def save_master(service, excel_bytes, file_id=None):
    media = MediaIoBaseUpload(excel_bytes, mimetype=XLSX_MIME, resumable=True)
    if file_id:
        service.files().update(fileId=file_id, media_body=media).execute()
        return file_id
    meta = {"name": MASTER_FILENAME, "mimeType": XLSX_MIME}
    created = service.files().create(body=meta, media_body=media, fields="id").execute()
    return created["id"]


# ==================================================
# ANTARA MUKA STREAMLIT
# ==================================================
st.set_page_config(page_title="Resit Analisa", page_icon="🧾", layout="centered")

# ==================================================
# GAYA / TEMA (CSS tersuai — rupa fintech profesional)
# ==================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&display=swap');

:root{
  --navy:#14305C; --navy2:#1E4A86; --accent:#2F8FBF;
  --ink:#1B2638; --muted:#5B6B85; --line:#E4E9F2;
  --bg:#EEF2F9; --card:#FFFFFF;
}
.stApp{ background:
  radial-gradient(900px 500px at 100% -10%, #E3ECFa 0%, transparent 55%),
  linear-gradient(180deg,#EEF2F9 0%,#F6F8FC 100%); }
html, body, [class*="css"]{ font-family:'IBM Plex Sans',sans-serif; color:var(--ink); }
h1,h2,h3,h4{ font-family:'Sora',sans-serif !important; color:var(--navy) !important; letter-spacing:-.01em; }

#MainMenu, footer, [data-testid="stHeader"]{ visibility:hidden; }
.block-container{ padding-top:1.4rem; max-width:780px; }

/* Hero */
.hero{ background:linear-gradient(120deg,var(--navy) 0%,var(--navy2) 70%,var(--accent) 130%);
  border-radius:20px; padding:26px 30px; color:#fff; display:flex; align-items:center; gap:18px;
  box-shadow:0 14px 34px rgba(20,48,92,.28); margin-bottom:22px; }
.hero .mark{ font-size:34px; background:rgba(255,255,255,.14); width:62px; height:62px;
  display:flex; align-items:center; justify-content:center; border-radius:16px; }
.hero h1{ color:#fff !important; margin:0; font-size:27px; }
.hero p{ margin:3px 0 0; color:#D7E3F7; font-size:13.5px; }

/* Section header */
.sect{ display:flex; align-items:center; gap:11px; margin:6px 0 12px; }
.sect .badge{ background:var(--navy); color:#fff; width:28px; height:28px; border-radius:9px;
  display:flex; align-items:center; justify-content:center; font-family:'Sora'; font-weight:700; font-size:14px; }
.sect .t{ font-family:'Sora'; font-weight:600; font-size:18px; color:var(--navy); }

/* Kad seksyen */
[data-testid="stVerticalBlockBorderWrapper"]{ background:var(--card); border:1px solid var(--line) !important;
  border-radius:16px; box-shadow:0 4px 18px rgba(20,48,92,.06); }

/* Butang */
.stButton>button, .stDownloadButton>button{
  background:var(--navy); color:#fff; border:0; border-radius:11px; padding:.55rem 1.1rem;
  font-family:'Sora'; font-weight:600; font-size:14px; transition:.18s; box-shadow:0 4px 12px rgba(20,48,92,.18); }
.stButton>button:hover, .stDownloadButton>button:hover{ background:var(--navy2); transform:translateY(-1px);
  box-shadow:0 7px 18px rgba(20,48,92,.26); color:#fff; }
.stLinkButton>a{ background:linear-gradient(120deg,var(--navy),var(--navy2)); color:#fff !important;
  border-radius:11px; padding:.7rem 1.4rem; font-family:'Sora'; font-weight:600; text-decoration:none;
  box-shadow:0 6px 16px rgba(20,48,92,.22); transition:.18s; }
.stLinkButton>a:hover{ transform:translateY(-1px); box-shadow:0 9px 22px rgba(20,48,92,.3); }

/* Metrik */
[data-testid="stMetric"]{ background:var(--card); border:1px solid var(--line); border-radius:14px;
  padding:14px 16px; box-shadow:0 3px 12px rgba(20,48,92,.05); }
[data-testid="stMetricLabel"]{ color:var(--muted); font-weight:600; }
[data-testid="stMetricValue"]{ color:var(--navy); font-family:'Sora'; }

/* Status pill */
.pill{ display:inline-flex; align-items:center; gap:8px; background:#E8F5EE; color:#1B7A46;
  border:1px solid #BCE3CC; border-radius:999px; padding:6px 14px; font-size:13px; font-weight:600; }
.pill .dot{ width:9px; height:9px; border-radius:50%; background:#22A35A; box-shadow:0 0 0 3px rgba(34,163,90,.18); }

/* Uploader & dataframe */
[data-testid="stFileUploader"]{ background:#F7F9FD; border:1.5px dashed #C5D2E8; border-radius:14px; padding:8px; }
[data-testid="stDataFrame"]{ border-radius:12px; overflow:hidden; border:1px solid var(--line); }

/* Footer tersuai */
.foot{ text-align:center; color:var(--muted); font-size:12.5px; margin-top:26px;
  padding-top:16px; border-top:1px solid var(--line); }
.foot a{ color:var(--accent); text-decoration:none; font-weight:600; }
</style>
""", unsafe_allow_html=True)


def section(num, title):
    st.markdown(
        f'<div class="sect"><div class="badge">{num}</div><div class="t">{title}</div></div>',
        unsafe_allow_html=True,
    )


def hero():
    st.markdown(
        '<div class="hero"><div class="mark">🧾</div>'
        '<div><h1>Resit Analisa</h1>'
        '<p>Automasi Perakaunan Berkuasa AI · Resit → Excel → Google Drive</p></div></div>',
        unsafe_allow_html=True,
    )


def footer():
    st.markdown(
        f'<div class="foot">Dibangunkan oleh <b>{AUTHOR_NAME}</b> · '
        f'<a href="mailto:{AUTHOR_EMAIL}">{AUTHOR_EMAIL}</a></div>',
        unsafe_allow_html=True,
    )


# ==================================================
# STATE + OAUTH REDIRECT
# ==================================================
for key, default in [("transactions", []), ("pending", []), ("master_file_id", None)]:
    if key not in st.session_state:
        st.session_state[key] = default

params = st.query_params
if "code" in params and "credentials" not in st.session_state:
    try:
        flow = get_flow()
        flow.fetch_token(code=params["code"])
        st.session_state.credentials = flow.credentials
        st.query_params.clear()
        st.rerun()
    except Exception as e:
        st.error(f"Gagal log masuk: {e}")

creds = st.session_state.get("credentials")
hero()

# ==================================================
# PINTU MASUK (log masuk wajib)
# ==================================================
if not creds:
    with st.container(border=True):
        st.markdown("#### 🔐 Log masuk diperlukan")
        st.write("Sila log masuk dengan akaun Google anda untuk mula menggunakan aplikasi. "
                 "Hanya pengguna yang dibenarkan boleh mengakses.")
        flow = get_flow()
        auth_url, _ = flow.authorization_url(
            prompt="consent", access_type="offline", include_granted_scopes="true"
        )
        st.link_button("Log masuk dengan Google", auth_url)
    footer()
    st.stop()

# ==================================================
# BAR STATUS + PAPAN PEMUKA METRIK
# ==================================================
service = drive_service(creds)
txns = st.session_state.transactions
total_amount = sum(float(t.get("amount", 0) or 0) for t in txns)
this_month = pd.Timestamp.now().strftime("%Y-%m")
month_amount = sum(
    float(t.get("amount", 0) or 0) for t in txns
    if str(t.get("transaction_date", "")).startswith(this_month)
)

c1, c2 = st.columns([3, 1])
with c1:
    st.markdown('<span class="pill"><span class="dot"></span>Tersambung ke Google Drive</span>',
                unsafe_allow_html=True)
with c2:
    if st.button("Log keluar"):
        st.session_state.clear()
        st.rerun()

m1, m2, m3 = st.columns(3)
m1.metric("Jumlah Transaksi", f"{len(txns)}")
m2.metric("Jumlah Perbelanjaan", f"RM {total_amount:,.2f}")
m3.metric(f"Bulan {this_month}", f"RM {month_amount:,.2f}")

st.write("")

# ==================================================
# SEKSYEN 1: MUAT FAIL INDUK
# ==================================================
with st.container(border=True):
    section("1", "Sambung Fail Induk")
    st.caption("Muat rekod sedia ada dari Google Drive (pilihan — langkau jika kali pertama).")
    if st.button("📂 Muat fail induk dari Drive"):
        fid = find_master(service)
        if fid:
            st.session_state.master_file_id = fid
            st.session_state.transactions = read_master_bytes(download_master(service, fid))
            st.success(f"{len(st.session_state.transactions)} transaksi dimuatkan.")
            st.rerun()
        else:
            st.session_state.master_file_id = None
            st.info("Tiada fail induk lagi — akan dicipta automatik bila anda simpan.")

# ==================================================
# SEKSYEN 2: UPLOAD RESIT
# ==================================================
with st.container(border=True):
    section("2", "Upload Resit Baru")
    st.caption("Boleh upload banyak gambar sekali gus. AI akan baca setiap satu.")
    receipts = st.file_uploader("Pilih gambar resit", type=["jpg", "jpeg", "png"],
                                accept_multiple_files=True, label_visibility="collapsed")
    if receipts and st.button("🔍 Analisa Semua Resit"):
        progress = st.progress(0)
        added, exact, pending = 0, 0, []
        for i, file in enumerate(receipts):
            try:
                data = extract_receipt_data(Image.open(file))
                status = classify(data, st.session_state.transactions)
                if status == "new":
                    st.session_state.transactions.append(data)
                    added += 1
                    st.write(f"✅ {file.name} — {data.get('vendor_name','?')} (RM{data.get('amount','?')})")
                elif status == "exact":
                    exact += 1
                    st.warning(f"⚠️ {file.name} — pendua tepat (no resit sama). Dilangkau.")
                else:
                    pending.append({"data": data, "name": file.name})
            except Exception as e:
                st.error(f"❌ {file.name}: {e}")
            progress.progress((i + 1) / len(receipts))
        st.session_state.pending = pending
        st.success(f"{added} ditambah · {exact} pendua dilangkau · {len(pending)} perlu disemak.")

# ==================================================
# SEKSYEN 3: TAMBAH TRANSAKSI MANUAL (TANPA RESIT)
# ==================================================
with st.container(border=True):
    section("3", "Tambah Transaksi Manual (Tanpa Resit)")
    st.caption("Untuk transaksi tanpa dokumen — contoh tol, parking, tip. Isi maklumat dan tambah.")

    # Kategori akaun debit yang seragam (chart of accounts ringkas)
    KATEGORI = [
        "Toll & Parking", "Travel Expense", "Fuel", "Office Supplies",
        "Utilities", "Meals & Entertainment", "Miscellaneous Expense", "Lain-lain",
    ]

    with st.form("manual_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            m_date = st.date_input("Tarikh")
            m_vendor = st.text_input("Vendor / Tempat", placeholder="cth: Tol PLUS, Parking DBKL")
            m_amount = st.number_input("Jumlah (RM)", min_value=0.0, step=0.50, format="%.2f")
        with c2:
            m_debit = st.selectbox("Akaun Debit (kategori)", KATEGORI)
            m_debit_lain = st.text_input("Jika 'Lain-lain', nyatakan kategori", placeholder="cth: Donation Expense")
            m_credit = st.text_input("Akaun Kredit", value="Cash")
        m_desc = st.text_input("Keterangan", placeholder="cth: Tol perjalanan ke pejabat klien")
        submitted = st.form_submit_button("➕ Tambah Transaksi")

    if submitted:
        if m_amount <= 0:
            st.error("Sila masukkan jumlah lebih daripada 0.")
        elif not m_vendor.strip():
            st.error("Sila isi Vendor / Tempat.")
        else:
            debit = m_debit_lain.strip() if (m_debit == "Lain-lain" and m_debit_lain.strip()) else m_debit
            txn = {
                "transaction_date": str(m_date),
                "vendor_name": m_vendor.strip(),
                "receipt_no": "",  # tiada resit (transaksi manual)
                "description": m_desc.strip() or m_vendor.strip(),
                "amount": float(m_amount),
                "debit_account": debit,
                "credit_account": m_credit.strip() or "Cash",
                "currency": "MYR",
            }
            st.session_state.transactions.append(txn)
            st.success(f"Ditambah: {txn['vendor_name']} — RM{txn['amount']:.2f} ({debit})")
            st.rerun()

# ==================================================
# SEMAKAN MANUSIA
# ==================================================
if st.session_state.pending:
    with st.container(border=True):
        section("!", "Perlu Disahkan")
        st.caption("Transaksi ini serupa & tiada nombor resit. Tanda yang BUKAN pendua.")
        for idx, item in enumerate(st.session_state.pending):
            d = item["data"]
            st.checkbox(
                f"{item['name']} — {d.get('vendor_name','?')} · RM{d.get('amount','?')} · {d.get('transaction_date','?')}",
                key=f"keep_{idx}",
            )
        if st.button("✅ Sahkan pilihan"):
            kept = 0
            for idx, item in enumerate(st.session_state.pending):
                if st.session_state.get(f"keep_{idx}"):
                    st.session_state.transactions.append(item["data"])
                    kept += 1
            st.session_state.pending = []
            st.success(f"{kept} transaksi ditambah sebagai pembelian berasingan.")
            st.rerun()

# ==================================================
# SEKSYEN 4: FAIL INDUK & LAPORAN
# ==================================================
with st.container(border=True):
    section("4", "Fail Induk & Laporan")
    if st.session_state.transactions:
        st.dataframe(pd.DataFrame(st.session_state.transactions), use_container_width=True)
        excel_file = build_master_excel(st.session_state.transactions)
        col1, col2 = st.columns(2)
        with col1:
            if st.button("☁️ Simpan ke Google Drive"):
                fid = save_master(drive_service(creds), excel_file,
                                  st.session_state.get("master_file_id"))
                st.session_state.master_file_id = fid
                st.success("Disimpan ke Google Drive ✅")
        with col2:
            st.download_button("📥 Muat turun (sandaran)", data=excel_file,
                               file_name=MASTER_FILENAME, mime=XLSX_MIME,
                               use_container_width=True)
    else:
        st.info("Belum ada transaksi. Muat fail induk atau upload resit di atas.")

footer()
