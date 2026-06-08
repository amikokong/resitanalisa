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
def extract_receipt_data(image):
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
    response = client.models.generate_content(
        model=MODEL_NAME,
        contents=[prompt, image],
        config={"response_mime_type": "application/json"},
    )
    return json.loads(response.text)


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
st.set_page_config(page_title="Resit Analisa", page_icon="🧾")
st.title("🧾 Resit Analisa → Fail Induk (Google Drive)")

for key, default in [("transactions", []), ("pending", []), ("master_file_id", None)]:
    if key not in st.session_state:
        st.session_state[key] = default

# ---- Tangani redirect balik dari Google ----
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

# ---- Seksyen 1: sambung Google Drive ----
st.subheader("1️⃣ Google Drive")
if not creds:
    flow = get_flow()
    auth_url, _ = flow.authorization_url(
        prompt="consent", access_type="offline", include_granted_scopes="true"
    )
    st.link_button("🔐 Log masuk Google Drive", auth_url)
    st.info("Log masuk sekali untuk baca/simpan fail induk terus dari Drive anda.")
else:
    st.success("✅ Tersambung ke Google Drive")
    service = drive_service(creds)
    if st.button("📂 Muat fail induk dari Drive"):
        fid = find_master(service)
        if fid:
            st.session_state.master_file_id = fid
            st.session_state.transactions = read_master_bytes(download_master(service, fid))
            st.success(f"{len(st.session_state.transactions)} transaksi dimuatkan dari Drive.")
        else:
            st.session_state.master_file_id = None
            st.info("Tiada fail induk lagi — akan dicipta automatik bila anda simpan.")

# ---- Seksyen 2: upload resit ----
st.subheader("2️⃣ Upload resit baru (boleh banyak sekali gus)")
receipts = st.file_uploader("Pilih gambar resit", type=["jpg", "jpeg", "png"],
                            accept_multiple_files=True)

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

# ---- Semakan manusia ----
if st.session_state.pending:
    st.subheader("🔎 Perlu disahkan — serupa, tiada nombor resit")
    st.write("Tanda yang **BUKAN pendua**:")
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

# ---- Seksyen 3: papar + simpan ----
st.subheader("3️⃣ Fail induk semasa")
if st.session_state.transactions:
    st.dataframe(pd.DataFrame(st.session_state.transactions))
    excel_file = build_master_excel(st.session_state.transactions)

    col1, col2 = st.columns(2)
    with col1:
        if creds and st.button("☁️ Simpan ke Google Drive"):
            fid = save_master(drive_service(creds), excel_file,
                              st.session_state.get("master_file_id"))
            st.session_state.master_file_id = fid
            st.success("Disimpan ke Google Drive ✅")
    with col2:
        st.download_button("📥 Muat turun (sandaran)", data=excel_file,
                           file_name=MASTER_FILENAME, mime=XLSX_MIME)
else:
    st.info("Belum ada transaksi. Muat fail induk dari Drive atau upload resit.")

# FOOTNOTE
st.markdown("---")
st.caption(f"Dibangunkan oleh **{AUTHOR_NAME}** · [{AUTHOR_EMAIL}](mailto:{AUTHOR_EMAIL})")
