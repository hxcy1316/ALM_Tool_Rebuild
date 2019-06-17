import win32com.client

td = win32com.client.Dispatch("TDApiOle80.TDConnection")


class ALM:
    def __init__(self, url, user_name, password, domain, project):
        self.url = url
        self.user_name = user_name
        self.password = password
        self.domain = domain
        self.project = project
        self.map_dict = {}
        self.full_lab_sub_folder_list = []
        self.full_instance_list = []

    def login(self):
        td.InitConnection(self.url)
        td.Login(self.user_name, self.password)
        td.Connect(self.domain, self.project)
        if td.Connected:
            print("connect successfully")

    def get_test_lab_sub_folder(self, parent_path):
        test_set_folder_factory = td.TestSetTreeManager
        test_set_folder = test_set_folder_factory.NodeByPath(parent_path)
        if test_set_folder.count > 0:
            return test_set_folder.SubNodes
        else:
            return False

    def get_test_lab_sub_folder_recursively(self, folder_path):
        sub_folders_list = self.get_test_lab_sub_folder(folder_path)
        if sub_folders_list is not False:
            self.full_lab_sub_folder_list.extend(sub_folders_list)
            for node in sub_folders_list:
                self.get_test_lab_sub_folder_recursively(node.path)
        return self.full_lab_sub_folder_list

    def get_test_set_list(self, test_lab_folder_path):
        test_set_folder_factory = td.TestSetTreeManager
        test_set_folder = test_set_folder_factory.NodeByPath(test_lab_folder_path)
        test_set_factory = test_set_folder.TestSetFactory
        test_set_list = test_set_factory.NewList("")
        return test_set_list

    def get_test_instance_list(self, test_set):
        test_instance_factory = test_set.TSTestFactory
        test_instance_list = test_instance_factory.NewList("")
        return test_instance_list

    def get_table_column_by_label(self, table, label):
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

    def map_column_label(self, label_list):
        map_column_label_dict = {}
        for label in label_list:
            map_column_label_dict.update(
                {label: self.get_table_column_by_label("Test", label)})
        return map_column_label_dict

    def get_test_instance_property(self, test_instance_path, test_instance_set_name, test_instance):
        current_case_property_dict = self.get_test_case_property(test_instance.testid)
        instance_property_dict = {
            "test_instance_path": test_instance_path,
            "test_set_name": test_instance_set_name,
            "test_instance_id": test_instance.id,
            "test_instance_status": test_instance.status,
            "test_instance_test_id": test_instance.testid,
            "test_instance_test_name": test_instance.testname,
            "test_instance_L1": current_case_property_dict["L1"],
            "test_instance_L2": current_case_property_dict["L2"],
            "test_instance_L3": current_case_property_dict["L3"],
            "test_instance_L4": current_case_property_dict["L4"]
        }
        return instance_property_dict

    def get_test_case_property(self, test_case_id):
        test_factory = td.TestFactory
        test_filter = test_factory.Filter
        test_filter["TS_TEST_ID"] = test_case_id
        test_list = test_filter.NewList()
        test_case = test_list[0]
        case_property_dict = {
            "test_id": test_case.ID,
            "test_name": test_case.Name,
            "L1": test_case.Field(self.map_dict.get("L1 Feature")),
            "L2": test_case.Field(self.map_dict.get("L2 Feature")),
            "L3": test_case.Field(self.map_dict.get("L3 Feature")),
            "L4": test_case.Field(self.map_dict.get("L4 Feature"))
        }
        return case_property_dict

    def disconnect(self):
        td.Disconnect()


if __name__ == "__main__":
    url = "http://15.83.240.100/qcbin"
    user_name = "chen.si_hp.com"
    password = "P@ssw0rd"
    domain = "TEST"
    project = "Test_WES"
    path = r"Root\Test_Chen"
    a = ALM(url, user_name, password, domain, project)
    try:

        a.login()
        a.get_test_lab_sub_folder_recursively(path)
        user_label_list = [
            'L1 Feature',
            'L2 Feature',
            'L3 Feature',
            'L4 Feature'
        ]
        a.map_dict = a.map_column_label(user_label_list)
        # Get test set from test_set_root_path and load instance to list
        if a.get_test_set_list(path).count > 0:
            for test_set in a.get_test_set_list(path):
                for instance in a.get_test_instance_list(test_set):
                    a.full_instance_list.append(a.get_test_instance_property(path, test_set.Name, instance))
        # Get test set from sub folder of test_set_root_path and loal instance to list
        for sub_folder in a.full_lab_sub_folder_list:
            for test_set in a.get_test_set_list(sub_folder.Path):
                for instance in a.get_test_instance_list(test_set):
                    # print(instance.id)
                    a.full_instance_list.append(a.get_test_instance_property(sub_folder.path, test_set.Name, instance))
        print(len(a.full_instance_list))
    except Exception as e:
        print(e)
    finally:
        a.disconnect()
