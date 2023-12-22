from openpyxl.utils import get_column_letter
from .utils import BaseTableDetector


def is_empty(cell):
    return cell is None or cell.value is None


class TableDetector(BaseTableDetector):
    def get_table_ranges(self, openpyxl_ws, **kwargs):

        tables = []
        visited = set()

        for row in openpyxl_ws.iter_rows():
            for cell in row:
                if is_empty(cell) or cell.coordinate in visited:
                    continue

                # Initialize boundaries of the table
                top, left = cell.row, cell.column
                bottom, right = top, left

                # Expand to the right
                for j in range(left + 1, openpyxl_ws.max_column + 1):
                    if is_empty(openpyxl_ws.cell(row=top, column=j)):
                        break
                    right = j

                # Expand downwards
                for i in range(top + 1, openpyxl_ws.max_row + 1):
                    if all(is_empty(openpyxl_ws.cell(row=i, column=j)) for j in range(left, right + 1)):
                        break
                    bottom = i

                # Add found table range
                start = f"{get_column_letter(left)}{top}"
                end = f"{get_column_letter(right)}{bottom}"
                tables.append(f"{start}:{end}")

                # Mark cells as visited
                for i in range(top, bottom + 1):
                    for j in range(left, right + 1):
                        visited.add(openpyxl_ws.cell(row=i, column=j).coordinate)

        return tables
