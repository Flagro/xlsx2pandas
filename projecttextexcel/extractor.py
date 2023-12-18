import pandas as pd
from typing import Optional, List, Dict, Union
from openpyxl import load_workbook
from .strategies.utils import get_strategy_parser
from .utils import get_in_memory_file


def get_df(file_path, 
           sheet_names: Optional[Union[str, List[str]]], 
           strategy='general', 
           **kwargs) -> Union[List[pd.DataFrame], Dict[str, List[pd.DataFrame]], Dict[str, pd.DataFrame], pd.DataFrame]:
    """
    Returns a list or a single pandas dataframe of the extracted data from the xlsx/xlsm file.
    """ 
    
    strategy_parser = get_strategy_parser(strategy)

    in_memory_file = get_in_memory_file(file_path)
    wb = load_workbook(in_memory_file, read_only=True)

    if isinstance(sheet_names, str):
        sheets = [sheet_names]
    elif isinstance(sheet_names, list):
        sheets = sheet_names
    else:
        sheets = wb.sheetnames

    dataframes_dict = dict()

    for sheet in sheets:
        if sheet not in wb.sheetnames:
            raise ValueError(f"Sheet {sheet} not found.")
        else:
            ws = wb[sheet]
            sheet_dataframes = strategy_parser.parse(ws, **kwargs)
            if len(sheet_dataframes) == 1:
                dataframes_dict[sheet] = sheet_dataframes[0]
            else:
                dataframes_dict[sheet] = sheet_dataframes
    wb.close()
    if len(dataframes_dict) == 1:
        return dataframes_dict[dataframes_dict.keys()[0]]
    else:
        return dataframes_dict
