from abc import ABC, abstractmethod
from typing import Any

import win32com.client.gencache


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
    def __init__(self, exhibit: BaseExhibit) -> None:
        self.exhibit = exhibit
        self._com_obj = None

    def __getattr__(self, name: str) -> Any:
        """
        Converts pythonic snake_case to PascalCase for Excel object properties.
        :param name: Property name
        :return: The Excel COM object property, if it exists.
        """
        return getattr(self, "".join(name_part.title() for name_part in name.split('_')))

    def _create_com_obj(self, clsid, event_handler=None):
        if event_handler is not None:
            self._com_obj = win32com.client.DispatchWithEvents(clsid, event_handler)
        else:
            self._com_obj = win32com.client.gencache.EnsureDispatch(clsid)
        self.setup_com_object()

    def setup_com_object(self) -> None:
        pass

    @staticmethod
    def _snake_to_pascal(name: str) -> str:
        return "".join(name_part.title() for name_part in name.split('_'))

    @staticmethod
    def _pascal_to_snake(name: str) -> str:
        return "".join(f"_{char.lower()}" if char.isupper() else char for char in name)

    @abstractmethod
    def _add_to_worksheet(self, *args, **kwargs) -> None:
        pass

    def _refresh(self, *args, **kwargs) -> None:
        pass


class BaseDependency(ABC):
    """
    Base class to represent a generic dependency between multiple components.
    """