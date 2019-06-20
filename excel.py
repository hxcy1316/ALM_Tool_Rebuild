import openpyxl
import os
from functools import reduce


class EXCEL():
    def __init__(self, result_file_path, test_instance_list):
        self.result_file_path = result_file_path
        self.test_instance_list = test_instance_list
        self.unique_instance_dict = {}
        self.wb = openpyxl.Workbook()
        self.sheet_execution_detail = "Execution Details"
        self.sheet_unique_execution_detail = "Unique Execution Details"

    def open_sheet(self, sheet_name):
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
        work_sheet = self.open_sheet(self.sheet_execution_detail)
        print("Start initial table in {}".format(work_sheet))
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
        print("Start filling the data ...")
        for test_instance in self.test_instance_list:
            work_sheet.cell(row, 1).value = test_instance["test_instance_test_id"]
            work_sheet.cell(row, 2).value = test_instance["test_instance_L1"]
            work_sheet.cell(row, 3).value = test_instance["test_instance_L2"]
            work_sheet.cell(row, 4).value = test_instance["test_instance_L3"]
            work_sheet.cell(row, 5).value = test_instance["test_instance_L4"]
            work_sheet.cell(row, 6).value = test_instance["test_instance_path"]
            work_sheet.cell(row, 7).value = test_instance["test_set_name"]
            work_sheet.cell(row, 8).value = test_instance["test_instance_id"]
            work_sheet.cell(row, 9).value = test_instance["test_instance_status"]
            row = row + 1
        print("End filling data in sheet {}".format(self.sheet_execution_detail))
        self.wb.save(self.result_file_path)

    def get_unique_execution_detail(self):
        work_sheet = self.open_sheet(self.sheet_unique_execution_detail)
        print("Start initial table in {}".format(work_sheet))
        work_sheet.cell(1, 1).value = "Test ID"
        work_sheet.cell(1, 2).value = "L1 Feature"
        work_sheet.cell(1, 3).value = "L2 Feature"
        work_sheet.cell(1, 4).value = "L3 Feature"
        work_sheet.cell(1, 5).value = "L4 Feature"
        work_sheet.cell(1, 6).value = "Unique Status"
        row = 2
        # unique_instance_dict = {<id>: <status>, <id>: <status>, ...}
        self.unique_instance_dict = self.__format_to_unique_id_status_list(self.__format_to_id_status_list(self.test_instance_list))
        # print(self.unique_instance_dict)
        print("Start filling the data ...")
        for instance_test_id, instance_test_status in self.unique_instance_dict.items():
            test_id = instance_test_id
            test_status = instance_test_status
            test_property_dict = self.__get_properties_by_id(test_id, self.test_instance_list)
            work_sheet.cell(row, 1).value = test_id
            work_sheet.cell(row, 2).value = test_property_dict["test_instance_L1"]
            work_sheet.cell(row, 3).value = test_property_dict["test_instance_L2"]
            work_sheet.cell(row, 4).value = test_property_dict["test_instance_L3"]
            work_sheet.cell(row, 5).value = test_property_dict["test_instance_L4"]
            work_sheet.cell(row, 6).value = test_status
            row = row + 1
        print("End filling data in sheet {}".format(self.sheet_unique_execution_detail))
        self.wb.save(self.result_file_path)

    def __get_properties_by_id(self, test_id, instance_list):
        for instance in instance_list:
            if test_id in instance.values():
                return instance

    def __format_to_id_status_list(self, full_instance_list):
        # id_status_list = []
        # for instance_dict in full_instance_list:
        #     id_status_list.append(self.__format_dict(instance_dict))
        id_status_list = map(self.__format_dict, full_instance_list)
        return id_status_list

    def __format_dict(self, temp_dict):
        formated_dict = {temp_dict["test_instance_test_id"]: temp_dict["test_instance_status"]}
        # print("formated dict is {}".format(formated_dict))
        return formated_dict

    def __format_to_unique_id_status_list(self, id_status_list):
        unique_id_status_list = reduce(self.__update_dict, id_status_list)
        return unique_id_status_list

    def __update_dict(self, dict1, dict2):
        for key, value in dict2.items():
            if key in dict1.keys():
                unique_status = self.__get_unique_status([value, dict1[key]])
                dict1.update({f'{key}': unique_status})
            else:
                dict1.update({f'{key}': dict2[key]})
        # print(dict1)
        return dict1

    def __get_unique_status(self, status_list):
        if "Failed" in status_list:
            return "Failed"
        elif "Passed" in status_list:
            return "Passed"
        else:
            return "No Run"

    def close(self):
        self.wb.close()


if __name__ == "__main__":
    full_instance_list = ['1', '2']
    file_path = r"C:\My Doc\My Github\ALM_Tool_REBUILD\test.xlsx"
    # get_execution_detail(file_path, full_instance_list)
