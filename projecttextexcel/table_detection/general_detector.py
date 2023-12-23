import openpyxl
from openpyxl.utils import get_column_letter
from .utils import BaseTableDetector


def is_empty(cell):
    return cell is None or cell.value is None or cell.value == ""


class TableDetector(BaseTableDetector):
    def find_table_end(openpyxl_ws, start_row, start_col):
        max_row = start_row
        max_col = start_col
        empty_row_count = 0
        empty_col_count = 0

        for row in openpyxl_ws.iter_rows(min_row=start_row, max_col=openpyxl_ws.max_column, max_row=openpyxl_ws.max_row):
            if all(is_empty(cell) for cell in row[start_col - 1:]):
                empty_row_count += 1
            else:
                empty_row_count = 0
                max_row = row[0].row

            if empty_row_count > 1:
                break

        for col in openpyxl_ws.iter_cols(min_col=start_col, max_row=openpyxl_ws.max_row, max_col=openpyxl_ws.max_column):
            if all(is_empty(cell) for cell in col[start_row - 1:max_row]):
                empty_col_count += 1
            else:
                empty_col_count = 0
                max_col = col[0].column

            if empty_col_count > 1:
                break

        return max_row, max_col

    def get_table_ranges(self, openpyxl_ws, **kwargs):
        tables = []
        visited = set()

        for row in openpyxl_ws.iter_rows():
            for cell in row:
                if cell.coordinate in visited or is_empty_cell(cell):
                    continue

                end_row, end_col = self.find_table_end(openpyxl_ws, cell.row, cell.column)
                table_range = f"{openpyxl.utils.get_column_letter(cell.column)}{cell.row}:" \
                              f"{openpyxl.utils.get_column_letter(end_col)}{end_row}"
                tables.append(table_range)

                # Mark cells as visited
                for r in range(cell.row, end_row + 1):
                    for c in range(cell.column, end_col + 1):
                        visited.add((r, c))

        return tables
