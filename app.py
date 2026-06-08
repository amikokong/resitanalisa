"""
app.py  — Resit Analisa (Streamlit + Gemini)  [Fail Induk + Anti-Pendua Pintar]
-------------------------------------------------------------------------------
Keupayaan:
  1. Upload BANYAK resit serentak
  2. AI (Gemini) ekstrak data + NOMBOR RESIT
  3. Anti-pendua 2 lapisan:
       - Lapisan 1: nombor resit unik -> pembeza utama
       - Lapisan 2: jika tiada nombor resit & nampak serupa -> minta
                    pengesahan manusia (checkbox), bukan buang terus
  4. Fail induk + ringkasan bulanan

Jalankan:
    pip install -r requirements.txt
    streamlit run app.py

Secrets:
    GEMINI_API_KEY = "AIza..."

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

MODEL_NAME = "gemini-2.5-flash"
MONEY_FORMAT = "#,##0.00"
BOLD = Font(bold=True)

AUTHOR_NAME = "Sulaiman Osman"
AUTHOR_EMAIL = "sulaimanosman03@gmail.com"

# receipt_no ditambah sebagai pembeza utama
FIELDS = [
    "transaction_date",
    "vendor_name",
    "receipt_no",
    "description",
    "amount",
    "debit_account",
    "credit_account",
    "currency",
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
# LOGIK ANTI-PENDUA
# ==================================================
def _vendor(t):
    return str(t.get("vendor_name", "")).strip().lower()


def classify(txn, existing):
    """
    Pulang salah satu:
      'new'   -> transaksi baru, boleh terus tambah
      'exact' -> pendua tepat (nombor resit sama) -> langkau
      'maybe' -> serupa tapi tiada nombor resit -> perlu disahkan manusia
    """
    receipt_no = str(txn.get("receipt_no", "")).strip()

    # Lapisan 1: ada nombor resit -> pembeza muktamad
    if receipt_no:
        for e in existing:
            if _vendor(e) == _vendor(txn) and str(e.get("receipt_no", "")).strip() == receipt_no:
                return "exact"
        return "new"

    # Lapisan 2: tiada nombor resit -> bandingkan tarikh+vendor+jumlah
    amt = round(float(txn.get("amount", 0) or 0), 2)
    date = str(txn.get("transaction_date", "")).strip()
    for e in existing:
        same = (
            str(e.get("transaction_date", "")).strip() == date
            and _vendor(e) == _vendor(txn)
            and round(float(e.get("amount", 0) or 0), 2) == amt
        )
        if same:
            return "maybe"
    return "new"


# ==================================================
# AI: ekstrak data satu resit (termasuk nombor resit)
# ==================================================
def extract_receipt_data(image):
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
    prompt = """
    Baca resit ini dan pulangkan HANYA objek JSON (tiada teks lain, tiada markdown).
    Cari nombor resit / invois / bil (jika ada) untuk 'receipt_no'.
    Jika tiada nombor resit dijumpai, letak "" (kosong).
    Guna format ini tepat-tepat:
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
# BACA fail induk lama
# ==================================================
def read_master(uploaded_xlsx):
    wb = load_workbook(uploaded_xlsx)
    if "Transactions" not in wb.sheetnames:
        return []
    ws = wb["Transactions"]
    transactions = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        txn = {FIELDS[i]: (row[i] if i < len(row) else "") for i in range(len(FIELDS))}
        transactions.append(txn)
    return transactions


# ==================================================
# BINA fail induk
# ==================================================
def build_master_excel(transactions):
    wb = Workbook()

    # SHEET 1 : TRANSACTIONS
    ws1 = wb.active
    ws1.title = "Transactions"
    write_headers(ws1, ["Date", "Vendor", "Receipt No", "Description", "Amount",
                        "Debit Account", "Credit Account", "Currency"])
    for t in transactions:
        ws1.append([
            t.get("transaction_date", ""),
            t.get("vendor_name", ""),
            t.get("receipt_no", ""),
            t.get("description", ""),
            float(t.get("amount", 0) or 0),
            t.get("debit_account", ""),
            t.get("credit_account", ""),
            t.get("currency", "MYR"),
        ])
    apply_money_format(ws1, ["E"])  # Amount = kolum E

    # SHEET 2 : JOURNAL ENTRY
    ws2 = wb.create_sheet("Journal Entry")
    write_headers(ws2, ["Date", "Description", "Account", "Debit", "Credit"])
    for t in transactions:
        amt = float(t.get("amount", 0) or 0)
        ws2.append([t.get("transaction_date", ""), t.get("description", ""),
                    t.get("debit_account", ""), amt, None])
        ws2.append([t.get("transaction_date", ""), t.get("description", ""),
                    t.get("credit_account", ""), None, amt])
    apply_money_format(ws2, ["D", "E"])

    # Ringkasan pandas
    df = pd.DataFrame(transactions)
    if not df.empty:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
        df["transaction_date"] = pd.to_datetime(df["transaction_date"], errors="coerce")
        df["month"] = df["transaction_date"].dt.strftime("%Y-%m")

    # SHEET 3 : MONTHLY SUMMARY
    ws3 = wb.create_sheet("Monthly Summary")
    write_headers(ws3, ["Month", "Total Expense (RM)", "Bil. Transaksi"])
    if not df.empty:
        monthly = df.groupby("month").agg(
            total=("amount", "sum"), count=("amount", "count")
        ).reset_index()
        for _, r in monthly.iterrows():
            ws3.append([r["month"], float(r["total"]), int(r["count"])])
    apply_money_format(ws3, ["B"])
    add_footer(ws3)

    # SHEET 4 : BY ACCOUNT
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
# ANTARA MUKA STREAMLIT
# ==================================================
st.set_page_config(page_title="Resit Analisa", page_icon="🧾")
st.title("🧾 Resit Analisa → Fail Induk Perakaunan")

if "transactions" not in st.session_state:
    st.session_state.transactions = []
if "pending" not in st.session_state:
    st.session_state.pending = []  # senarai {"data":..., "name":...} untuk disahkan

# Langkah 1
st.subheader("1️⃣ Sambung dari fail induk lama (pilihan)")
master_up = st.file_uploader("Upload fail induk Excel lama", type=["xlsx"], key="master")
if master_up and st.button("📂 Muatkan fail induk"):
    st.session_state.transactions = read_master(master_up)
    st.success(f"{len(st.session_state.transactions)} transaksi dimuatkan.")

# Langkah 2
st.subheader("2️⃣ Upload resit baru (boleh banyak sekali gus)")
receipts = st.file_uploader(
    "Pilih gambar resit", type=["jpg", "jpeg", "png"], accept_multiple_files=True
)

if receipts and st.button("🔍 Analisa Semua Resit"):
    progress = st.progress(0)
    added, exact = 0, 0
    pending = []
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
            else:  # maybe
                pending.append({"data": data, "name": file.name})
        except Exception as e:
            st.error(f"❌ {file.name}: {e}")
        progress.progress((i + 1) / len(receipts))
    st.session_state.pending = pending
    st.success(f"{added} ditambah · {exact} pendua dilangkau · {len(pending)} perlu disemak.")

# Bahagian semakan manusia (Lapisan 2)
if st.session_state.pending:
    st.subheader("🔎 Perlu disahkan — serupa, tiada nombor resit")
    st.write("Tanda yang **BUKAN pendua** (memang pembelian berasingan):")
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

# Langkah 3
st.subheader("3️⃣ Fail induk semasa")
if st.session_state.transactions:
    st.dataframe(pd.DataFrame(st.session_state.transactions))

    excel_file = build_master_excel(st.session_state.transactions)
    st.download_button(
        label="📥 Muat Turun Fail Induk (Excel)",
        data=excel_file,
        file_name="fail_induk_perakaunan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if st.button("🗑️ Kosongkan senarai"):
        st.session_state.transactions = []
        st.rerun()
else:
    st.info("Belum ada transaksi. Muatkan fail induk lama atau upload resit.")

# FOOTNOTE
st.markdown("---")
st.caption(f"Dibangunkan oleh **{AUTHOR_NAME}** · [{AUTHOR_EMAIL}](mailto:{AUTHOR_EMAIL})")
