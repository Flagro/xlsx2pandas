import pandas as pd
from .utils import BaseDataFrameConstructor


def join_header_rows(header_rows):
    """
    Joins a list of header rows into a single list of strings by underscore.
    """
    n = len(header_rows[0])
    headers = [""] * n
    for row in header_rows:
        for i, cell in enumerate(row):
            if cell is not None:
                headers[i] += "_" + str(cell) if headers[i] != "" else str(cell)
    return headers


class DataFrameConstructor(BaseDataFrameConstructor):
    def construct_dataframe(self, openpyxl_ws, table_range, header_rows_cnt, **kwargs):
        data = []
        for row in openpyxl_ws[table_range]:
            data.append([cell.value for cell in row])

        # Construct the dataframe
        df = pd.DataFrame(data[header_rows_cnt:], columns=join_header_rows(data[0:header_rows_cnt]))
        return df
