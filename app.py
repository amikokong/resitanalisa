"""
app.py
Accounting Excel Generator (Flask)
-----------------------------------
Terima satu transaksi (JSON) -> jana fail Excel 4 sheet:
  1. Journal Entry
  2. General Ledger
  3. Trial Balance
  4. Profit & Loss

Jalankan:
    pip install flask openpyxl
    python app.py

Endpoint:
    POST /generate    -> hantar JSON, dapat balik fail .xlsx
    GET  /            -> mesej status ringkas
"""

from io import BytesIO

from flask import Flask, request, send_file, jsonify
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

app = Flask(__name__)

# Format nombor matawang (boleh ubah ikut keperluan)
MONEY_FORMAT = "#,##0.00"
BOLD = Font(bold=True)


# ==================================================
# HELPER: tulis baris header tebal (elak ulang kod)
# ==================================================
def write_headers(ws, headers):
    """Tulis senarai header pada baris 1 dengan font tebal."""
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

    # ---- Ambil data input (dengan nilai default selamat) ----
    date = data.get("transaction_date", "")
    desc = data.get("description", "")
    amount = float(data.get("amount", 0))
    debit_acc = data.get("debit_account", "")
    credit_acc = data.get("credit_account", "")

    # ==================================================
    # SHEET 1 : JOURNAL ENTRY
    # ==================================================
    ws1 = wb.active
    ws1.title = "Journal Entry"
    write_headers(ws1, ["Date", "Description", "Account", "Debit", "Credit"])
    ws1.append([date, desc, debit_acc, amount, None])
    ws1.append([date, desc, credit_acc, None, amount])
    apply_money_format(ws1, ["D", "E"])

    # ==================================================
    # SHEET 2 : GENERAL LEDGER
    # ==================================================
    ws2 = wb.create_sheet("General Ledger")
    write_headers(ws2, ["Account", "Date", "Description", "Debit", "Credit", "Balance"])
    ws2.append([debit_acc, date, desc, amount, None, amount])
    ws2.append([credit_acc, date, desc, None, amount, -amount])
    apply_money_format(ws2, ["D", "E", "F"])

    # ==================================================
    # SHEET 3 : TRIAL BALANCE
    # ==================================================
    ws3 = wb.create_sheet("Trial Balance")
    write_headers(ws3, ["Account", "Debit", "Credit"])
    ws3.append([debit_acc, amount, None])
    ws3.append([credit_acc, None, amount])
    ws3.append(["TOTAL", "=SUM(B2:B3)", "=SUM(C2:C3)"])
    apply_money_format(ws3, ["B", "C"])

    # ==================================================
    # SHEET 4 : PROFIT & LOSS
    # ==================================================
    ws4 = wb.create_sheet("Profit & Loss")
    write_headers(ws4, ["Category", "Amount (RM)"])
    ws4.append(["Expense", amount])
    ws4.append(["Revenue", 0])
    ws4.append(["Net Profit", "=B3-B2"])
    apply_money_format(ws4, ["B"])

    # ==================================================
    # FORMAT LEBAR KOLUM (semua sheet)
    # ==================================================
    for ws in wb.worksheets:
        for column in ws.columns:
            ws.column_dimensions[column[0].column_letter].width = 25

    # ---- Simpan ke memory, bukan ke disk ----
    excel_file = BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)
    return excel_file


# ==================================================
# ROUTES
# ==================================================
@app.route("/")
def home():
    return jsonify({"status": "ok", "message": "Accounting Excel Generator is running."})


@app.route("/generate", methods=["POST"])
def generate():
    """Terima JSON transaksi, pulangkan fail Excel."""
    data = request.get_json(silent=True)
    if not data:
        return jsonify({"error": "JSON body diperlukan."}), 400

    excel_file = create_accounting_excel(data)
    filename = f"accounting_{data.get('transaction_date', 'output')}.xlsx"

    return send_file(
        excel_file,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
