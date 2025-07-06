from openpyxl import Workbook

def create_excel_sheet():
    wb = Workbook()
    ws = wb.active
    ws.title = "Discrepant POs"
    ws.append(["S.No.", "PO Number", "Total Amount", "Billed Amount", "Deviation", "Site ID", "Site Name", "PO Description"])
    return wb, ws