def get_merged_cell(openpyxl_ws, openpyxl_cell):
    merged_range = [s for s in openpyxl_ws.merged_cells.ranges if openpyxl_cell.coordinate in s]
    if merged_range:
        return openpyxl_ws.cell(merged_range[0].min_row, merged_range[0].min_col)
    else:
        openpyxl_cell
