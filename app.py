import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

default_prompt = """
You are a professional accounting assistant.

Analyze the uploaded receipt, invoice or financial document.

Extract and classify accounting entries.

Return ONLY valid JSON.

Format:

{
  "transaction_date": "YYYY-MM-DD",
  "vendor_name": "",
  "description": "",
  "amount": 0,
  "debit_account": "",
  "credit_account": "",
  "currency": "MYR"
}

Accounting Rules:

Office supplies → Debit Office Supplies Expense
Equipment → Debit Equipment Asset
Fuel → Debit Vehicle Expense
Internet → Debit Internet Expense
Utility → Debit Utility Expense
Payment by cash → Credit Cash
Payment by bank → Credit Bank Account

Never guess amount.
def create_accounting_excel(data):

    wb = Workbook()
    ws = wb.active

    ws.title = "General Ledger"

    headers = [
        "Date",
        "Description",
        "Account",
        "Debit (RM)",
        "Credit (RM)"
    ]

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.font = Font(bold=True)

    ws.append([
        data.get("transaction_date"),
        data.get("description"),
        data.get("debit_account"),
        data.get("amount"),
        ""
    ])

    ws.append([
        data.get("transaction_date"),
        data.get("description"),
        data.get("credit_account"),
        "",
        data.get("amount")
    ])

    ws.append(["","","","", ""])

    ws.append([
        "",
        "",
        "TOTAL",
        f"=SUM(D2:D3)",
        f"=SUM(E2:E3)"
    ])

    excel_file = BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)

    return excel_file
Return null if unknown.
"""
