import pandas as pd


class PD():
    def __init__(self):
        pass

    def __read_excel(self, file_path, sheet_name):
        return pd.read_excel(file_path, sheet_name)

    def get_data_frame(self, file_path, sheet_name):
        data_frame = pd.DataFrame(self.__read_excel(file_path, sheet_name))
        print("Load excel data to pandas Data Frame format successfully")
        return data_frame

    def pivot_data_frame(self, data_frame):
        # Fill in 'NA' for empty cells
        data_frame_remove_na = data_frame.fillna('NA')
        pivoted_data_frame = data_frame_remove_na.pivot_table(
            index=['L1 Feature', 'L2 Feature', 'L3 Feature', 'L4 Feature'],
            columns='Test_instance_status',
            values='Test Instance ID',
            aggfunc=len,
            fill_value=0,
            margins=True)
        return pivoted_data_frame

    def format_pivot_data_frame(self, pivot_data_frame, custom_list):
        status_list = [
            'Passed', 'Failed', 'No Run', 'Not Completed', 'Block', 'NA'
        ]
        formated_data_frame = pivot_data_frame
        for status in status_list:
            if status not in formated_data_frame.columns:
                formated_data_frame = self.__add_column(
                    formated_data_frame, status, '0')

        for column in custom_list:
            formated_data_frame = self.__add_column(formated_data_frame, column, '0')
        return formated_data_frame

    def __add_column(self, pivot_data_frame, column_name, value):
        kwargs = {}
        kwargs[column_name] = pd.Series(value, index=pivot_data_frame.index)
        data_frame = pivot_data_frame.assign(**kwargs)
        return data_frame
