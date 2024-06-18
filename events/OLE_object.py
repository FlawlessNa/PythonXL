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
        raise NotImplementedError

    def OnLostFocus(self) -> None:
        """
        Occurs when an ActiveX control loses input focus.
        """
        raise NotImplementedError
