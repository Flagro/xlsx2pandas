import pandas as pd
from .utils import BaseDataFrameConstructor


def join_header_rows(header_rows):
    """
    Joins a list of header rows into a single list of strings by underscore.
    """
    return [str("_".join([str(cell) for cell in row])) for row in header_rows]

class DataFrameConstructor(BaseDataFrameConstructor):
    def construct_dataframe(self, openpyxl_ws, table_range, header_rows_cnt, **kwargs):
        data = []
        for row in openpyxl_ws[table_range]:
            data.append([cell.value for cell in row])

        # Construct the dataframe
        df = pd.DataFrame(data[header_rows_cnt:], columns=join_header_rows(data[0:header_rows_cnt]))
        return df
