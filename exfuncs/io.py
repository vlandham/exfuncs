import openpyxl
import csv


def load_workbook(filename):
    wb = openpyxl.load_workbook(filename)
    return wb


def save_workbook(wb, filename):
    wb.save(filename)


def save_workbook_to_csv(wb, sheet_name, filename):
    ws = wb[sheet_name]
    with open(filename, "w") as f:
        c = csv.writer(f)
        for r in ws.rows:
            c.writerow([cell.value for cell in r])
