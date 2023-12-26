import importlib
from abc import ABC, abstractmethod


class BaseHeaderSeparator(ABC):
    def __init__(self):
        pass

    @abstractmethod
    def get_header_rows_cnt(self, openpyxl_ws, table_range, **kwargs):
        pass


def get_header_separator(strategy):
    try:
        # Dynamically import the strategy class
        strategy_module = importlib.import_module(f".{strategy}_separator", __package__)
    except ImportError:
        raise ValueError(f"Strategy {strategy} not found.")

    # Get the strategy class
    strategy_class = getattr(strategy_module, "HeaderSeparator")
    if not issubclass(strategy_class, BaseHeaderSeparator):
        raise TypeError("Invalid strategy type.")

    # Create an instance of the strategy and parse the file
    strategy = strategy_class()
    return strategy
