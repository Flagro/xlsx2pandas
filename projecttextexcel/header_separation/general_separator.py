from .utils import BaseHeaderSeparator
from openpyxl.utils import range_boundaries
from ..openpyxl_utils import get_merged_cell, get_cell_type, range_generator


class HeaderSeparator(BaseHeaderSeparator):
    def get_header_rows_cnt(self, openpyxl_ws, table_range, **kwargs):
        # Parse the table range
        min_col, min_row, max_col, max_row = range_boundaries(table_range)
        # Initialize variables
        header_scores = [0] * (max_row - min_row + 1)
        last_types = [None] * (max_col - min_col + 1)
        # Iterate through columns and rows
        for row_idx, col_idx, cell in range_generator(openpyxl_ws, table_range):
            cell = get_merged_cell(openpyxl_ws, cell)
            cell_type = get_cell_type(cell)

            # Update score based on type change
            if cell_type != last_types[col_idx - min_col] and last_types[col_idx - min_col] is not None:
                header_scores[row_idx - min_row] += self.get_type_change_score(last_types[col_idx - min_col], cell_type, **kwargs)
            
            last_types[col_idx - min_col] = cell_type

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
