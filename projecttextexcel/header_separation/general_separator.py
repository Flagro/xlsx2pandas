from openpyxl.utils import range_boundaries
from .utils import BaseHeaderSeparator


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
        for col in range(min_col, max_col + 1):
            last_type = None
            for row in range(min_row, max_row + 1):
                cell = openpyxl_ws.cell(row, col)
                cell_type = get_cell_type(cell)

                # Update score based on type change
                if cell_type != last_type and last_type is not None:
                    header_scores[row - min_row] += self.get_type_change_score(last_type, cell_type, **kwargs)
                
                last_type = cell_type

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
