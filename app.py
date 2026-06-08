def create_accounting_excel(data):

```
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

wb = Workbook()

date = data.get("transaction_date")
desc = data.get("description")
amount = float(data.get("amount", 0))

debit_acc = data.get("debit_account")
credit_acc = data.get("credit_account")

# ==================================================
# SHEET 1 : JOURNAL ENTRY
# ==================================================

ws1 = wb.active
ws1.title = "Journal Entry"

headers = [
    "Date",
    "Description",
    "Account",
    "Debit",
    "Credit"
]

for col, header in enumerate(headers, 1):
    cell = ws1.cell(row=1, column=col)
    cell.value = header
    cell.font = Font(bold=True)

ws1.append([
    date,
    desc,
    debit_acc,
    amount,
    ""
])

ws1.append([
    date,
    desc,
    credit_acc,
    "",
    amount
])

# ==================================================
# SHEET 2 : GENERAL LEDGER
# ==================================================

ws2 = wb.create_sheet("General Ledger")

ledger_headers = [
    "Account",
    "Date",
    "Description",
    "Debit",
    "Credit",
    "Balance"
]

for col, header in enumerate(ledger_headers, 1):
    cell = ws2.cell(row=1, column=col)
    cell.value = header
    cell.font = Font(bold=True)

ws2.append([
    debit_acc,
    date,
    desc,
    amount,
    "",
    amount
])

ws2.append([
    credit_acc,
    date,
    desc,
    "",
    amount,
    -amount
])

# ==================================================
# SHEET 3 : TRIAL BALANCE
# ==================================================

ws3 = wb.create_sheet("Trial Balance")

tb_headers = [
    "Account",
    "Debit",
    "Credit"
]

for col, header in enumerate(tb_headers, 1):
    cell = ws3.cell(row=1, column=col)
    cell.value = header
    cell.font = Font(bold=True)

ws3.append([
    debit_acc,
    amount,
    ""
])

ws3.append([
    credit_acc,
    "",
    amount
])

ws3.append([
    "TOTAL",
    "=SUM(B2:B3)",
    "=SUM(C2:C3)"
])

# ==================================================
# SHEET 4 : PROFIT & LOSS
# ==================================================

ws4 = wb.create_sheet("Profit & Loss")

pnl_headers = [
    "Category",
    "Amount (RM)"
]

for col, header in enumerate(pnl_headers, 1):
    cell = ws4.cell(row=1, column=col)
    cell.value = header
    cell.font = Font(bold=True)

ws4.append([
    "Expense",
    amount
])

ws4.append([
    "Revenue",
    0
])

ws4.append([
    "Net Profit",
    "=B3-B2"
])

# ==================================================
# FORMAT COLUMN WIDTH
# ==================================================

for ws in wb.worksheets:
    for column in ws.columns:
        ws.column_dimensions[column[0].column_letter].width = 25

excel_file = BytesIO()
wb.save(excel_file)
excel_file.seek(0)

return excel_file
```
