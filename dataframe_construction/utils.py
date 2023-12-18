import importlib
from abc import ABC, abstractmethod


class BaseDataFrameConstructor(ABC):
    def __init__(self):
        pass

    @abstractmethod
    def construct_dataframe(self, openpyxl_ws, table_range, header_rows_cnt, **kwargs):
        """
        Constructs a pandas dataframe from the given table range and header rows count.
        """
        pass


def get_header_separator(strategy):
    try:
        # Dynamically import the strategy class
        strategy_module = importlib.import_module(f".{strategy}_constructor", __package__)
    except ImportError:
        raise ValueError(f"Strategy {strategy} not found.")

    # Get the strategy class
    strategy_class = getattr(strategy_module, "DataFrameConstructor")
    if not issubclass(strategy_class, BaseDataFrameConstructor):
        raise TypeError("Invalid strategy type.")

    # Create an instance of the strategy and parse the file
    strategy = strategy_class()
    return strategy
