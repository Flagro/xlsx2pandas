import io


def get_in_memory_file(file_path):
    """
    Returns an in-memory file of the xlsx/xlsm file.
    """
    with open(file_path, "rb") as f:
        in_memory_file = io.BytesIO(f.read())
    return in_memory_file
