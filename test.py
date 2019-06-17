import win32com.client


def alm_login(username, password, domain, alm_project):
    td.InitConnection(url)
    td.Login(username, password)
    td.Connect(domain, alm_project)
    if td.Connected:
        print("connect successfully")


def get_sub_test_set_folder(parent_path):
    test_set_folder_factory = td.TestSetTreeManager
    test_set_folder = test_set_folder_factory.NodeByPath(parent_path)
    if test_set_folder.count > 0:
        return test_set_folder.SubNodes
    else:
        return False


def get_sub_test_set_folder_recursively(folder_path):
    sub_folders_list = get_sub_test_set_folder(folder_path)
    if sub_folders_list is not False:
        full_sub_folders_list.extend(sub_folders_list)
        for node in sub_folders_list:
            get_sub_test_set_folder_recursively(node.path)
    return full_sub_folders_list


def get_test_set_list(project_path):
    test_set_folder_factory = td.TestSetTreeManager
    test_set_folder = test_set_folder_factory.NodeByPath(project_path)
    test_set_factory = test_set_folder.TestSetFactory
    test_set_list = test_set_factory.NewList("")
    return test_set_list


def get_test_instance_list(test_set):
    test_instance_factory = test_set.TSTestFactory
    test_instance_list = test_instance_factory.NewList("")
    return test_instance_list


def get_column_by_label(table, label):
    field_list = td.Fields(table)
    find_label = False
    for field in field_list:
        field_property = field.Property
        if field_property.UserLabel == label:
            find_label = True
            return field_property.DBColumnName
    if find_label is False:
        print("Can't find the property " + label)
        return False


def map_column_label(label_list):
    map_column_label_dict = {}
    for label in label_list:
        map_column_label_dict.update(
            {label: get_column_by_label("Test", label)})
    return map_column_label_dict


def get_test_instance_property(test_instance_path, test_instance_set_name, test_instance):
    current_case_property_dict = get_test_case_property(test_instance.testid)
    instance_property_dict = {
        "test_instance_path": test_instance_path,
        "test_set_name": test_instance_set_name,
        "test_instance_id": test_instance.id,
        "test_instance_status": test_instance.status,
        "test_instance_test_id": test_instance.testid,
        # "test_instance_test_name": test_instance.testname,
        "test_instance_L1": current_case_property_dict["L1"],
        "test_instance_L2": current_case_property_dict["L2"],
        "test_instance_L3": current_case_property_dict["L3"],
        "test_instance_L4": current_case_property_dict["L4"]
    }
    return instance_property_dict


def get_test_case_property(test_case_id):
    test_factory = td.TestFactory
    test_filter = test_factory.Filter
    test_filter["TS_TEST_ID"] = test_case_id
    test_list = test_filter.NewList()
    test_case = test_list[0]
    case_property_dict = {
        "test_id": test_case.ID,
        "test_name": test_case.Name,
        "L1": test_case.Field(map_dict.get("L1 Feature")),
        "L2": test_case.Field(map_dict.get("L2 Feature")),
        "L3": test_case.Field(map_dict.get("L3 Feature")),
        "L4": test_case.Field(map_dict.get("L4 Feature"))
    }
    return case_property_dict


if __name__ == "__main__":
    url = "http://15.83.240.100/qcbin"

    try:
        td = win32com.client.Dispatch("TDApiOle80.TDConnection")
    except Exception as e:
        print(e)
        print(
            "ALM OTA Library not found, make sure you have registered OTAClient.dll successfully"
        )
        exit()
    try:
        alm_login("chen.si_hp.com", "P@ssw0rd", "TEST", "Test_WES")
        test_set_root_path = r'Root\Test_Chen'
        full_sub_folders_list = []
        full_instance_list = []
        get_sub_test_set_folder_recursively(test_set_root_path)
        # print(len(full_sub_folders_list))
        user_label_list = [
            'L1 Feature',
            'L2 Feature',
            'L3 Feature',
            'L4 Feature'
        ]
        map_dict = map_column_label(user_label_list)
        # Get test set from test_set_root_path and load instance to list
        if get_test_set_list(test_set_root_path).count > 0:
            for test_set in get_test_set_list(test_set_root_path):
                for instance in get_test_instance_list(test_set):
                    full_instance_list.append(get_test_instance_property(test_set_root_path, test_set.Name, instance))
        # Get test set from sub folder of test_set_root_path and loal instance to list
        for sub_folder in full_sub_folders_list:
            for test_set in get_test_set_list(sub_folder.Path):
                for instance in get_test_instance_list(test_set):
                    # print(instance.id)
                    full_instance_list.append(get_test_instance_property(sub_folder.path, test_set.Name, instance))
        print(len(full_instance_list))
    except Exception as e:
        print(e)
    finally:
        td.Disconnect()
