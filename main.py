import alm
import excel


def main():
    url = "http://15.83.240.100/qcbin"
    user_name = "chen.si_hp.com"
    password = "P@ssw0rd"
    domain = "TEST"
    project = "Test_WES"
    path = r"Root\Test_Chen"
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
        alm_instance.map_dict = alm_instance.map_column_label(user_label_list)
        # Get test set from test_set_root_path and load instance to list
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
        print(len(alm_instance.full_instance_list))
        # Write to Excel
        file_path = r"C:\My Doc\My Github\ALM_Tool_REBUILD\test.xlsx"
        excel.get_execution_detail(file_path, alm_instance.full_instance_list)
    # Continue with alm
    except Exception as e:
        print(e)
    finally:
        alm_instance.disconnect()


if __name__ == "__main__":
    main()
