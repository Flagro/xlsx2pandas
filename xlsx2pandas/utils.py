import io


def get_in_memory_file(file_path):
    """
    Returns an in-memory file of the xlsx/xlsm file.
    """
    with open(file_path, "rb") as f:
        in_memory_file = io.BytesIO(f.read())
    return in_memory_file


def prettify_workbook_dataframes_output(dataframes_dict):
    output = dict()
    for sheet, dataframes in dataframes_dict.items():
        if len(dataframes) == 1:
            output[sheet] = dataframes[0]
        else:
            output[sheet] = dataframes
    if len(output) == 1:
        return list(output.values())[0]
    else:
        return output
