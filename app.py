"""
app.py  — Resit Analisa (Streamlit + Gemini)
--------------------------------------------
Aliran kerja:
  1. Pengguna upload gambar resit/invois
  2. Gemini (google-genai) baca & ekstrak data jadi JSON
  3. Jana fail Excel 4 sheet (Journal, Ledger, Trial Balance, P&L)
  4. Pengguna muat turun Excel

Jalankan setempat:
    pip install -r requirements.txt
    streamlit run app.py

Set API key (Streamlit Cloud: Settings > Secrets):
    GEMINI_API_KEY = "xxxxxxxx"
"""

import json
from io import BytesIO

import streamlit as st
from PIL import Image
from google import genai
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# Tukar nama model di sini jika perlu
MODEL_NAME = "gemini-2.5-flash"
MONEY_FORMAT = "#,##0.00"
BOLD = Font(bold=True)


# ==================================================
# HELPER EXCEL
# ==================================================
def write_headers(ws, headers):
    """Tulis header tebal pada baris 1."""
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.font = BOLD
        cell.alignment = Alignment(horizontal="center")


def apply_money_format(ws, columns, start_row=2):
    """Pakai format matawang pada kolum tertentu (cth: ['D', 'E'])."""
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            if cell.column_letter in columns and isinstance(cell.value, (int, float)):
                cell.number_format = MONEY_FORMAT


# ==================================================
# FUNGSI UTAMA: jana workbook Excel
# ==================================================
def create_accounting_excel(data):
    """Bina workbook Excel 4 sheet daripada satu transaksi."""
    wb = Workbook()

    date = data.get("transaction_date", "")
    desc = data.get("description", "")
    amount = float(data.get("amount", 0))
    debit_acc = data.get("debit_account", "")
    credit_acc = data.get("credit_account", "")

    # SHEET 1 : JOURNAL ENTRY
    ws1 = wb.active
    ws1.title = "Journal Entry"
    write_headers(ws1, ["Date", "Description", "Account", "Debit", "Credit"])
    ws1.append([date, desc, debit_acc, amount, None])
    ws1.append([date, desc, credit_acc, None, amount])
    apply_money_format(ws1, ["D", "E"])

    # SHEET 2 : GENERAL LEDGER
    ws2 = wb.create_sheet("General Ledger")
    write_headers(ws2, ["Account", "Date", "Description", "Debit", "Credit", "Balance"])
    ws2.append([debit_acc, date, desc, amount, None, amount])
    ws2.append([credit_acc, date, desc, None, amount, -amount])
    apply_money_format(ws2, ["D", "E", "F"])

    # SHEET 3 : TRIAL BALANCE
    ws3 = wb.create_sheet("Trial Balance")
    write_headers(ws3, ["Account", "Debit", "Credit"])
    ws3.append([debit_acc, amount, None])
    ws3.append([credit_acc, None, amount])
    ws3.append(["TOTAL", "=SUM(B2:B3)", "=SUM(C2:C3)"])
    apply_money_format(ws3, ["B", "C"])

    # SHEET 4 : PROFIT & LOSS
    ws4 = wb.create_sheet("Profit & Loss")
    write_headers(ws4, ["Category", "Amount (RM)"])
    ws4.append(["Expense", amount])
    ws4.append(["Revenue", 0])
    ws4.append(["Net Profit", "=B3-B2"])
    apply_money_format(ws4, ["B"])

    # Lebar kolum semua sheet
    for ws in wb.worksheets:
        for column in ws.columns:
            ws.column_dimensions[column[0].column_letter].width = 25

    excel_file = BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)
    return excel_file


# ==================================================
# FUNGSI AI: ekstrak data resit guna Gemini
# ==================================================
def extract_receipt_data(image):
    """Hantar gambar resit ke Gemini, pulang dict transaksi."""
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

    prompt = """
    Baca resit ini dan pulangkan HANYA objek JSON (tiada teks lain, tiada markdown).
    Guna format ini tepat-tepat:
    {
      "transaction_date": "YYYY-MM-DD",
      "vendor_name": "...",
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
# ANTARA MUKA STREAMLIT
# ==================================================
st.set_page_config(page_title="Resit Analisa", page_icon="🧾")
st.title("🧾 Resit Analisa → Excel Perakaunan")
st.write("Upload gambar resit. AI akan ekstrak data dan jana fail Excel automatik.")

uploaded = st.file_uploader("Pilih gambar resit", type=["jpg", "jpeg", "png"])

if uploaded:
    image = Image.open(uploaded)
    st.image(image, caption="Resit dimuat naik", width=300)

    if st.button("🔍 Analisa Resit"):
        with st.spinner("AI sedang membaca resit..."):
            try:
                data = extract_receipt_data(image)
                st.success("Berjaya diekstrak!")
                st.json(data)

                excel_file = create_accounting_excel(data)
                st.download_button(
                    label="📥 Muat Turun Excel",
                    data=excel_file,
                    file_name=f"accounting_{data.get('transaction_date', 'output')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Ralat: {e}")
