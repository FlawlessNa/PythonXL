from events import OLEObjectEvents


class ListBoxEventHandler(OLEObjectEvents):
    def __init__(self):
        self.exhibit = None

    def OnChange(self, *args, **kwargs) -> None:
        # breakpoint()
        if self.exhibit._ready:
            self.exhibit._refresh_data()
