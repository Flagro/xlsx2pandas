import importlib
from abc import ABC, abstractmethod


class BaseTableDetector(ABC):
    def __init__(self):
        pass

    @abstractmethod
    def get_table_ranges(self, openpyxl_ws, **kwargs):
        pass


def get_table_detector(strategy):
    try:
        # Dynamically import the strategy class
        strategy_module = importlib.import_module(f".{strategy}_detector", __package__)
    except ImportError:
        raise ValueError(f"Strategy {strategy} not found.")

    # Get the strategy class
    strategy_class = getattr(strategy_module, "TableDetector")
    if not issubclass(strategy_class, BaseTableDetector):
        raise TypeError("Invalid strategy type.")

    # Create an instance of the strategy and parse the file
    strategy = strategy_class()
    return strategy
