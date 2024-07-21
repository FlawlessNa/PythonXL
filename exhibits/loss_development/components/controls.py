from exhibits.base import BaseComponent, BaseDependency


class ControlListBox(BaseComponent):
    PROG_ID = "Forms.ListBox.1"

    def __init__(self, exhibit: BaseComponent, segment: str, values: list) -> None:
        super().__init__(exhibit)
        self.segment = segment
        self.values = values
        self._ole = None

    def _add_to_worksheet(self, offset: int = 0) -> None:
        self._ole = self._ws.OLEObjects().Add(self.PROG_ID)
        cell = self._ws.Cells(offset + 1, 1)
        self._ole.Left, self._ole.Top, self._ole.Width = cell.Left, cell.Top, cell.Width
        self._ole.Height = cell.Height * len(self.values)
        self._ole.Placement = c.xlMoveAndSize
        self._ole.FontBold = True
        self._ole.List = self.values
        self._ole.Value = self.values[0]
        self._ole.Name = f'{self.segment}_dropdown'
