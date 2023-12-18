import pandas as pd
from .utils import BaseDataFrameConstructor

class DataFrameConstructor(BaseDataFrameConstructor):
    def construct_dataframe(self, openpyxl_ws, table_range, header_rows_cnt, **kwargs):
        data = []
        for row in openpyxl_ws[table_range]:
            data.append([cell.value for cell in row])

        # Construct the dataframe
        df = pd.DataFrame(data[header_rows_cnt:], columns=data[0:header_rows_cnt][0])
        return df
