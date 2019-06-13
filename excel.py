import win32com.client

if __name__ == "__main__":
    excel_app = win32com.client.Dispatch("Excel.Application")
    workbook = excel_app.Workbooks.Open(
        r'C:\Users\sich\Downloads\ALM login since 2019.xlsx')
    print(workbook.Worksheets('Mapping').Cells(1, 1).Value)
    excel_app.Application.Quit()
