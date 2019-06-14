import win32com.client


def alm_login(username, password, domain, alm_project):
    td.InitConnection(url)
    td.Login(username, password)
    td.Connect(domain, alm_project)
    if td.Connected:
        print("connect successfully")


def get_build_list(project_path):
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
        map_column_label_dict.update({label: get_column_by_label("Test", label)})
    return map_column_label_dict


def get_test_instance_property(test_instance):
    instance_property_dict = {
        "test_instance_id": test_instance.id,
        "test_instance_status": test_instance.status,
        "test_instance_L1": test_instance.Field(map_dict.get("L1 Feature")),
        "test_instance_L2": test_instance.Field(map_dict.get("L2 Feature")),
        "test_instance_L3": test_instance.Field(map_dict.get("L3 Feature")),
        "test_instance_L4": test_instance.Field(map_dict.get("L4 Feature")),
    }
    return instance_property_dict


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
        alm_login("chen.si", "P@ssw0rd", "TEST", "Test_WES")
        print(get_build_list(r"Root\Test_Chen").count)
        user_label_list = [
            'L1 Feature',
            'L2 Feature',
            'L3 Feature',
            'L4 Feature'
        ]
        map_dict = map_column_label(user_label_list)
        for ts in (get_build_list(r"Root\Test_Chen")):
            # print(get_test_instance_list(ts).count)
            print(get_test_instance_property(ts))

    except Exception as e:
        print(e)
    finally:
        td.Disconnect()
