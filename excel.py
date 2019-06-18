import openpyxl
import os


def get_total_execution_detail(result_file_path, test_instance_list):
    try:
        if not os.path.exists(result_file_path):
            print("File Doesn't Exist")
            wb = openpyxl.Workbook()
            wb.save(result_file_path)
            wb.close()
        wb = openpyxl.load_workbook(result_file_path)
        sheet_name = "Total Execution Details"
        if sheet_name in wb.sheetnames:
            print("Sheet already exist, remove and recreate the sheet {}".format(sheet_name))
            wb.remove(wb[sheet_name])
        wb.create_sheet(sheet_name)
        ws_total_execution_detail = wb[sheet_name]
        # Initial Table
        ws_total_execution_detail.cell(1, 1).value = "Test ID"
        ws_total_execution_detail.cell(1, 2).value = "L1 Feature"
        ws_total_execution_detail.cell(1, 3).value = "L2 Feature"
        ws_total_execution_detail.cell(1, 4).value = "L3 Feature"
        ws_total_execution_detail.cell(1, 5).value = "L4 Feature"
        ws_total_execution_detail.cell(1, 6).value = "test_instance_path"
        ws_total_execution_detail.cell(1, 7).value = "test_set_name"
        ws_total_execution_detail.cell(1, 8).value = "Test Instance ID"
        ws_total_execution_detail.cell(1, 9).value = "test_instance_status"
        row = 2
        for test_instance_property in test_instance_list:
            ws_total_execution_detail.cell(row, 1).value = test_instance_property["test_instance_test_id"]
            ws_total_execution_detail.cell(row, 2).value = test_instance_property["test_instance_L1"]
            ws_total_execution_detail.cell(row, 3).value = test_instance_property["test_instance_L2"]
            ws_total_execution_detail.cell(row, 4).value = test_instance_property["test_instance_L3"]
            ws_total_execution_detail.cell(row, 5).value = test_instance_property["test_instance_L4"]
            ws_total_execution_detail.cell(row, 6).value = test_instance_property["test_instance_path"]
            ws_total_execution_detail.cell(row, 7).value = test_instance_property["test_set_name"]
            ws_total_execution_detail.cell(row, 8).value = test_instance_property["test_instance_id"]
            ws_total_execution_detail.cell(row, 9).value = test_instance_property["test_instance_status"]
            row = row + 1
        wb.save(result_file_path)
    except Exception as e:
        print(e)
    finally:
        wb.close()


if __name__ == "__main__":
    full_instance_list = ['1', '2']
    file_path = r"C:\My Doc\My Github\ALM_Tool_REBUILD\test.xlsx"
    get_total_execution_detail(file_path, full_instance_list)
