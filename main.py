import alm
import excel
import pandas_handler
import os
from datetime import datetime


def Get_Execution_Summary(url, user_name, password, domain, project, path):
    # url = "http://15.83.240.100/qcbin"
    # user_name = "chen.si_hp.com"
    # password = "P@ssw0rd"
    # domain = "DEFAULT"
    # project = "WES_2016"
    # path = r"Root\E625_WES7_2019_OOC"
    user_label_list = ['L1 Feature', 'L2 Feature', 'L3 Feature', 'L4 Feature']
    alm_instance = alm.ALM(url, user_name, password, domain, project)
    #
    # Get Instance List from ALM
    try:
        alm_instance.login()
        alm_instance.get_test_lab_sub_folder_recursively(path)
        print("Interate test folder successfully")
        alm_instance.map_dict = alm_instance.map_column_label(user_label_list)
        # Get test set from test_set_root_path and load instance to list
        print("Start pulling test instance ...")
        if alm_instance.get_test_set_list(path).count > 0:
            for test_set in alm_instance.get_test_set_list(path):
                for instance in alm_instance.get_test_instance_list(test_set):
                    alm_instance.full_instance_list.append(
                        alm_instance.get_test_instance_property(
                            path, test_set.Name, instance))
        # Get test set from sub folder of test_set_root_path and loal instance to list
        for sub_folder in alm_instance.full_lab_sub_folder_list:
            for test_set in alm_instance.get_test_set_list(sub_folder.Path):
                for instance in alm_instance.get_test_instance_list(test_set):
                    # print(instance.id)
                    alm_instance.full_instance_list.append(
                        alm_instance.get_test_instance_property(
                            sub_folder.path, test_set.Name, instance))
        print("Add {} instances into list.".format(
            len(alm_instance.full_instance_list)))
    except Exception as e:
        print(e)
    finally:
        alm_instance.disconnect()
    #
    # Write to Excel
    try:
        now = datetime.now()
        dt_string = now.strftime("%Y_%m_%d_%H_%M_%S")
        file_path = os.path.join(os.getcwd(), "Project_Execution_Status_" + dt_string + "." + "xlsx")
        excel_app = excel.EXCEL(file_path, alm_instance.full_instance_list)
        excel_app.get_execution_detail()
        excel_app.get_unique_execution_detail()
        sheet_unique_execution_detail = excel_app.sheet_unique_execution_detail
        sheet_execution_detail = excel_app.sheet_execution_detail
        #
        # Manipulate Data with pandas
        pdh = pandas_handler.PD()
        # Prepare unique data frame
        print(
            "Preparing data frame of {}".format(sheet_unique_execution_detail))
        raw_data_frame_unique = pdh.get_data_frame(
            file_path, sheet_unique_execution_detail)
        pivoted_data_frame_unique = pdh.pivot_data_frame_unique(
            raw_data_frame_unique)
        full_list_unique = pdh.unique_status_list
        data_frame_unique = pdh.format_pivot_data_frame(
            pivoted_data_frame_unique, full_list_unique)
        data_frame_unique = data_frame_unique.rename(
            columns={
                'Passed': 'Unique_Passed_Test',
                'Executed_Test': 'Unique_Executed_Test',
                'All': 'Unique_Planned_Test',
                'Pass_Rate': 'Unique_Pass_Rate',
                'Execution_Rate': 'Unique_Execution_Rate'
            })
        data_frame_unique = data_frame_unique.drop(['Failed', 'No Run'],
                                                   axis=1)
        # Prepare total data frame
        print("Preparing data frame of {}".format(sheet_execution_detail))
        raw_data_frame_total = pdh.get_data_frame(file_path,
                                                  sheet_execution_detail)
        pivoted_data_frame_total = pdh.pivot_data_frame_total(
            raw_data_frame_total)
        full_list_total = pdh.full_status_list
        data_frame_total = pdh.format_pivot_data_frame(
            pivoted_data_frame_total, full_list_total)
        data_frame_total = data_frame_total.rename(
            columns={
                'Passed': 'Total_Passed_Test',
                'Executed_Test': 'Total_Executed_Test',
                'All': 'Total_Planned_Test',
                'Pass_Rate': 'Total_Pass_Rate',
                'Execution_Rate': 'Total_Execution_Rate'
            })
        data_frame_total = data_frame_total.drop(
            ['Failed', 'No Run', 'NA', 'Block', 'Not Completed'], axis=1)
        columns_order = [
            'Unique_Planned_Test', 'Unique_Executed_Test',
            'Unique_Passed_Test', 'Unique_Execution_Rate', 'Unique_Pass_Rate',
            'Total_Planned_Test', 'Total_Executed_Test', 'Total_Passed_Test',
            'Total_Execution_Rate', 'Total_Pass_Rate'
        ]
        # Join two data frame together and reset the index
        merged_data_frame = data_frame_total.join(data_frame_unique)
        merged_data_frame = merged_data_frame[columns_order]
        merged_data_frame = merged_data_frame.reset_index(
            level=['L1 Feature', 'L2 Feature', 'L3 Feature', 'L4 Feature'])
        # Remove the last 'All' row
        merged_data_frame.drop(merged_data_frame.tail(1).index, inplace=True)
        # Save merged_data_frame back to array
        data_array = merged_data_frame.get_values().tolist()
        excel_app.load_array_to_sheet(
            data_array, excel_app.sheet_execution_summary)
        excel_app.wb.save(file_path)
        print("Load data successfully in sheet {}. Start formating ...".format(
            excel_app.sheet_execution_summary))
        # Format with Win32com
        excel.format_execution_summary(file_path, excel_app.sheet_execution_summary)
    except Exception as e:
        print(e)
    finally:
        excel_app.close()
        pdh = None


if __name__ == "__main__":
    Get_Execution_Summary()
