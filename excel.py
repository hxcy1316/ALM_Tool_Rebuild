import xlwings as xw
import os


if __name__ == "__main__":
    try:
        excel_app = xw.App(visible=True, add_book=False)
        excel_app.display_alerts = False
        excel_app.screen_updating = False
        file_path = r"C:\My Doc\My Github\ALM_Tool_REBUILD\test.xlsx"
        if os.path.exists(file_path):
            wb = excel_app.books.open(file_path)        
        else:
            wb = excel_app.books.add()
            wb.save(file_path)
    except Exception as e:
        print(e)
    finally:
        wb.close()
        excel_app.quit()
