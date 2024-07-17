from abc import ABC


# noinspection PyPep8Naming
class OLEObjectEvents(ABC):
    """
    Events that can be triggered through interactions with an OLEObject object in Excel.
    """
    def OnGotFocus(self) -> None:
        """
        Occurs when an ActiveX control gets input focus.
        """
        print(f'Got Focus: {self}')

    def OnLostFocus(self) -> None:
        """
        Occurs when an ActiveX control loses input focus.
        """
        print(f'Lost Focus: {self}')

    def OnClick(self) -> None:
        """
        Occurs when an ActiveX control is clicked.
        """
        print(f'Clicked: {self}')

    def OnChange(self) -> None:
        """
        Occurs when the content of the ActiveX control changes.
        """
        print(f'Changed: {self}')
