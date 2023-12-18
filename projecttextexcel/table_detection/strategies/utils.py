import importlib
from abc import ABC, abstractmethod


class BaseExcelParser(ABC):
    def __init__(self):
        pass

    @abstractmethod
    def get_dataframes(self, **kwargs):
        pass


def get_strategy_parser(strategy):
    try:
        # Dynamically import the strategy class
        strategy_module = importlib.import_module(f".{strategy}_strategy", __package__)
    except ImportError:
        raise ValueError(f"Strategy {strategy} not found.")

    # Get the strategy class
    strategy_class = getattr(strategy_module, "ExcelParser")
    if not issubclass(strategy_class, BaseExcelParser):
        raise TypeError("Invalid strategy type.")

    # Create an instance of the strategy and parse the file
    strategy = strategy_class()
    return strategy
