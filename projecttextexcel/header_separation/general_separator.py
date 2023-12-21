from openpyxl.utils import range_boundaries
from .utils import BaseHeaderSeparator
from ..openpyxl_utils import get_merged_cell


def get_cell_type(self, cell):
        """ Determine the type of a cell's value. """
        if cell.data_type == 'n':
            return "float" if '.' in str(cell.value) else "int"
        elif cell.data_type == 'd':
            return "datetime"
        elif cell.data_type == 's':
            return "str"
        elif cell.data_type == 'f':
            return 'formula'
        else:
            return 'unknown'


class HeaderSeparator(BaseHeaderSeparator):
    def get_header_rows_cnt(self, openpyxl_ws, table_range, **kwargs):
        # Parse the table range
        min_col, min_row, max_col, max_row = range_boundaries(table_range)

        # Initialize variables
        header_scores = [0] * (max_row - min_row + 1)

        # Iterate through columns and rows
        for row_idx, row in enumerate(openpyxl_ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col)):
            last_types = [None] * (max_col - min_col + 1)
            for col_idx, cell in enumerate(row):
                cell = get_merged_cell(openpyxl_ws, cell)
                cell_type = self.get_cell_type(cell)

                # Update score based on type change
                if cell_type != last_types[col_idx] and last_types[col_idx] is not None:
                    header_scores[row_idx] += self.get_type_change_score(last_types[col_idx], cell_type, **kwargs)
                
                last_types[col_idx] = cell_type

        # Predict number of header rows
        return header_scores.index(max(header_scores)) + 1
    
    def get_type_change_score(self, last_type, current_type, **kwargs):
        """ Weighted scoring for type change. """
        weights = kwargs.get('weights', {
            ("int", "str"): 2,
            ("float", "str"): 2,
            ("datetime", "str"): 3,
            ("str", "int"): 1,
            ("str", "float"): 1,
            ("str", "datetime"): 2,
            ('formula', "str"): 1
        })
        return weights.get((last_type, current_type), 0)
