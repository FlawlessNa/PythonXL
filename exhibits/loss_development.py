import chainladder as cl
import pywintypes
from win32com.client import DispatchBaseClass, constants as c
from excel.utilities import get_range_ref_for_shape
from .base import BaseExhibit


class LossDevelopmentExhibit:

    _ws_name = 'LossDevelopment'

    def __init__(self, data: cl.Triangle) -> None:
        self.data = data
        if not self.data.is_cumulative:
            self.data = self.data.incr_to_cum()
        self._validations = self.data.index.to_dict(orient='list')
        self._amounts = self.data.columns.to_list()
        self._ws = None

    def load_into(self, wb: DispatchBaseClass):
        if self._ws_name in (ws.Name for ws in wb.Worksheets):
            self._ws = wb.Worksheets(self._ws_name)
        else:
            self._ws = wb.Worksheets.Add()
            self._ws.Name = self._ws_name
        self._add_dropdowns()
        # self.add_segment_dropdowns(ws)
        breakpoint()
        # ws.Range(get_range_ref_for_shape(*self.data.shape[-2:], first_row=5)).Value = (
        #     self.data.iloc[0, 0].to_frame().fillna('=NA()').values
        # )
        return wb

    def _add_dropdowns(self) -> None:
        longest_val = max(
            [item for sublist in self._validations.values() for item in sublist],
            key=len
        )
        self._ws.Cells(1, 1).Value = f"{'x' * max(30, len(longest_val))}"
        self._ws.Columns(1).AutoFit()
        self._ws.Cells(1, 1).ClearContents()

        num_values = 0
        for segment, values in self._validations.items():
            ole = self._ws.OLEObjects().Add("Forms.ListBox.1")  # TODO - Dispatch with Events
            cell = self._ws.Cells(num_values + 1, 1)
            ole.Left, ole.Top, ole.Width = cell.Left, cell.Top, cell.Width
            ole.Height = cell.Height * len(values)
            ole.Placement = c.xlMoveAndSize
            active_x = ole.Object
            breakpoint()
            active_x.List = values
            active_x.Object.Value = values[0]
            active_x.Object.FontBold = True
            active_x.BackColor = ...
            active_x.ForeColor = ...

            num_values += len(values)
        breakpoint()

    def add_segment_dropdowns(self, ws) -> None:
        for idx, col in enumerate(self.data.index.columns):
            cell = ws.Cells(idx + 1, 1)
            validator = cell.Validation
            validator.Add(
                Type=c.xlValidateList,
                Formula1=', '.join(self.data.index[col].values),
            )
            validator.InputTitle = col
            validator.InputMessage = f'Used for data segmentation'
            validator.ErrorTitle = col
            validator.ErrorMessage = (
                f'Invalid value for {col}. '
                f'Options are [{", ".join(self.data.index[col].values)}]'
            )
            validator.InCellDropdown = validator.ShowInput = validator.ShowError = True
            validator.IgnoreBlank = False
            cell.Value = self.data.index[col].values[0]
