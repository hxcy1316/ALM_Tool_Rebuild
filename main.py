import alm
import excel


def main():
    url = "http://15.83.240.100/qcbin"
    user_name = "chen.si_hp.com"
    password = "P@ssw0rd"
    domain = "DEFAULT"
    project = "WES_2016"
    path = r"Root\E625_WES7_2019_OOC"
    user_label_list = [
        'L1 Feature',
        'L2 Feature',
        'L3 Feature',
        'L4 Feature'
    ]
    alm_instance = alm.ALM(url, user_name, password, domain, project)
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
                    alm_instance.full_instance_list.append(alm_instance.get_test_instance_property(path, test_set.Name, instance))
        # Get test set from sub folder of test_set_root_path and loal instance to list
        for sub_folder in alm_instance.full_lab_sub_folder_list:
            for test_set in alm_instance.get_test_set_list(sub_folder.Path):
                for instance in alm_instance.get_test_instance_list(test_set):
                    # print(instance.id)
                    alm_instance.full_instance_list.append(alm_instance.get_test_instance_property(sub_folder.path, test_set.Name, instance))
        print("Add {} instances into list.".format(len(alm_instance.full_instance_list)))
        # Write to Excel
        file_path = r"C:\My Doc\My Github\ALM_Tool_REBUILD\test.xlsx"
        excel_app = excel.EXCEL(file_path, alm_instance.full_instance_list)
        try:
            excel_app.get_execution_detail()
            excel_app.get_unique_execution_detail()
        except Exception as e:
            print(e)
        finally:
            excel_app.close()
    # Continue with alm
    except Exception as e:
        print(e)
    finally:
        alm_instance.disconnect()


if __name__ == "__main__":
    main()
