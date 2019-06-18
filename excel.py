import openpyxl
import os


class EXCEL():
    def __init__(self, result_file_path, test_instance_list):
        self.result_file_path = result_file_path
        self.test_instance_list = test_instance_list
        self.wb = openpyxl.Workbook()
        self.sheet_execution_detail = "Execution Details"
        self.sheet_unique_execution_detail = "Unique Execution Details"

    def __open_sheet(self, sheet_name):
        if not os.path.exists(self.result_file_path):
            print("File Doesn't Exist")
            self.wb.save(self.result_file_path)
            self.wb.close()
        self.wb = openpyxl.load_workbook(self.result_file_path)
        print("Open workbook {} successfully".format(self.result_file_path))
        if sheet_name in self.wb.sheetnames:
            print("Sheet already exist, remove and recreate the sheet {}".format(sheet_name))
            self.wb.remove(self.wb[sheet_name])
        self.wb.create_sheet(sheet_name)
        return self.wb[sheet_name]

    def get_execution_detail(self):
        work_sheet = self.__open_sheet(self.sheet_execution_detail)
        # Initial Table
        work_sheet.cell(1, 1).value = "Test ID"
        work_sheet.cell(1, 2).value = "L1 Feature"
        work_sheet.cell(1, 3).value = "L2 Feature"
        work_sheet.cell(1, 4).value = "L3 Feature"
        work_sheet.cell(1, 5).value = "L4 Feature"
        work_sheet.cell(1, 6).value = "Test_instance_path"
        work_sheet.cell(1, 7).value = "Test_set_name"
        work_sheet.cell(1, 8).value = "Test Instance ID"
        work_sheet.cell(1, 9).value = "Test_instance_status"
        row = 2
        # Fill Data
        for test_instance_property in self.test_instance_list:
            work_sheet.cell(row, 1).value = test_instance_property["test_instance_test_id"]
            work_sheet.cell(row, 2).value = test_instance_property["test_instance_L1"]
            work_sheet.cell(row, 3).value = test_instance_property["test_instance_L2"]
            work_sheet.cell(row, 4).value = test_instance_property["test_instance_L3"]
            work_sheet.cell(row, 5).value = test_instance_property["test_instance_L4"]
            work_sheet.cell(row, 6).value = test_instance_property["test_instance_path"]
            work_sheet.cell(row, 7).value = test_instance_property["test_set_name"]
            work_sheet.cell(row, 8).value = test_instance_property["test_instance_id"]
            work_sheet.cell(row, 9).value = test_instance_property["test_instance_status"]
            row = row + 1
        self.wb.save(self.result_file_path)

    def get_unique_execution_detail(self):
        work_sheet = self.__open_sheet(self.sheet_unique_execution_detail)
        # Initial Table
        work_sheet.cell(1, 1).value = "Test ID"
        work_sheet.cell(1, 2).value = "L1 Feature"
        work_sheet.cell(1, 3).value = "L2 Feature"
        work_sheet.cell(1, 4).value = "L3 Feature"
        work_sheet.cell(1, 5).value = "L4 Feature"
        work_sheet.cell(1, 6).value = "Unique Status"
        # row = 2
        # unique_case_list = []
        # Fill Data
        # for test_instance_property in self.test_instance_list:
        #     if test_instance_property["test_instance_test_id"] in unique_case_list:

    # def __get_unique_status(self, )
    #     @Todo

    def close(self):
        self.wb.close()


if __name__ == "__main__":
    full_instance_list = ['1', '2']
    file_path = r"C:\My Doc\My Github\ALM_Tool_REBUILD\test.xlsx"
    # get_execution_detail(file_path, full_instance_list)
