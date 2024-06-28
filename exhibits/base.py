from abc import ABC, abstractmethod
from typing import Any


class BaseExhibit(ABC):
    def __init__(self, data: Any) -> None:
        self.data = data

    @property
    @abstractmethod
    def controls(self) -> list:
        pass
