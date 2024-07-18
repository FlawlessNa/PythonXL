from abc import ABC, abstractmethod
from typing import Any


class BaseExhibit(ABC):
    """
    Base class to represent a generic exhibit.
    """
    def __init__(self, data: Any) -> None:
        self.data = data

    @property
    @abstractmethod
    def components(self) -> list["BaseComponent"]:
        pass


class BaseComponent(ABC):
    """
    Base class to represent a generic component of an exhibit.
    """


class BaseDependency(ABC):
    """
    Base class to represent a generic dependency between multiple components.
    """