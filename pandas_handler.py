import pandas as pd


class PD():
    def __init__(self):
        self.full_status_list = ['NA', 'Block', 'No Run', 'Not Completed', 'Failed', 'Passed']
        self.unique_status_list = ['No Run', 'Failed', 'Passed']

    def __read_excel(self, file_path, sheet_name):
        return pd.read_excel(file_path, sheet_name)

    def get_data_frame(self, file_path, sheet_name):
        data_frame = pd.DataFrame(self.__read_excel(file_path, sheet_name))
        print("Load excel data to pandas Data Frame format successfully")
        return data_frame

    def pivot_data_frame_total(self, data_frame):
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

    def pivot_data_frame_unique(self, data_frame):
        # Fill in 'NA' for empty cells
        data_frame_remove_na = data_frame.fillna('NA')
        pivoted_data_frame = data_frame_remove_na.pivot_table(
            index=['L1 Feature', 'L2 Feature', 'L3 Feature', 'L4 Feature'],
            columns='Unique Status',
            values='Test ID',
            aggfunc=len,
            fill_value=0,
            margins=True)
        return pivoted_data_frame   

    def format_pivot_data_frame(self, pivot_data_frame, column_list):
        full_status_data_frame = self.__add_status_columns(pivot_data_frame, column_list)
        formatted_data_frame = self.__add_summary_columns(full_status_data_frame)
        return formatted_data_frame
    
    def __add_summary_columns(self, pivot_data_frame):
        # "{0:.0f}%".format(0.33 * 100)
        pivot_data_frame = pivot_data_frame.assign(Total_Execution_Rate=lambda x: ((x['Failed'] + x['Passed']) / x['All']) * 100)
        pivot_data_frame_final = pivot_data_frame.assign(Total_Pass_Details=lambda x: (x['Passed'] / x['All']) * 100)
        return pivot_data_frame_final

    def __add_status_columns(self, pivot_data_frame, column_list):
        kwargs = self.__init_status_columns(pivot_data_frame, column_list)
        full_status_data_frame = pivot_data_frame.assign(**kwargs)
        return full_status_data_frame

    def __init_status_columns(self, pivot_data_frame, column_list):
        columns_in_dict = {}
        for status in column_list:
            if status not in pivot_data_frame.columns:
                columns_in_dict[status] = pd.Series('0', index=pivot_data_frame.index)
        return columns_in_dict
