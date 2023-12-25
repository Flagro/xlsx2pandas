from ..openpyxl_utils import get_merged_openpyxl_cell, range_generator, is_cell_empty, get_excel_coordinate
from .utils import BaseTableDetector


def dfs_cluster_search(cur_point, cur_table, table_dict, allowed_horizontal_gap=1, allowed_vertical_gap=1):
    table_dict[cur_point] = cur_table
    for dx in range(-allowed_vertical_gap, allowed_vertical_gap + 1):
        for dy in range(-allowed_horizontal_gap, allowed_horizontal_gap + 1):
            next_point = (cur_point[0] + dx, cur_point[1] + dy)
            if next_point in table_dict and table_dict[next_point] is None:
                dfs_cluster_search(next_point, cur_table, table_dict)


class TableDetector(BaseTableDetector):
    def get_table_ranges(self, openpyxl_ws, **kwargs):
        non_empty_cells_tables = dict()
        # Get all cells based on a certain requirement (emptiness in this case)
        for row_idx, col_idx, cell in range_generator(openpyxl_ws):
            cell = get_merged_openpyxl_cell(openpyxl_ws, cell)
            if not is_cell_empty(cell):
                non_empty_cells_tables[(row_idx, col_idx)] = None

        print(f"Non-empty cells: {non_empty_cells_tables}")
        # Join cells in continuous ranges
        for i, cell_coordinate in enumerate(non_empty_cells_tables.keys(), start=1):
            if non_empty_cells_tables[cell] is None:
                dfs_cluster_search(cell_coordinate, i, non_empty_cells_tables, 2, 2)
        
        # For each range get tuple (min_col, min_row, max_col, max_row)
        continuous_ranges = dict()
        for cell_coordinate, table_idx in non_empty_cells_tables.items():
            if table_idx not in continuous_ranges:
                continuous_ranges[table_idx] = (cell_coordinate[1], cell_coordinate[0], cell_coordinate[1], cell_coordinate[0])
            else:
                continuous_ranges[table_idx] = (
                    min(continuous_ranges[table_idx][0], cell_coordinate[1]),
                    min(continuous_ranges[table_idx][1], cell_coordinate[0]),
                    max(continuous_ranges[table_idx][2], cell_coordinate[1]),
                    max(continuous_ranges[table_idx][3], cell_coordinate[0])
                )
        ranges = continuous_ranges.values()
        print(f"Continuous ranges: {ranges}")

        # Apply waterfall algorithm to find the lower bound of each table
        merged_ranges = ranges
        # TODO: implement ranges merge operation based on datatype changes and range horizontal lengths

        # Convert to A1 notation
        tables = []
        for min_col, min_row, max_col, max_row in merged_ranges:
            tables.append(get_excel_coordinate(min_col, min_row, max_col, max_row))
        print(f"Table ranges: {tables}")

        return tables
