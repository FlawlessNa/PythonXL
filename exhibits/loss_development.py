import chainladder as cl
import pywintypes
from win32com.client import DispatchBaseClass, constants as c
from excel.utilities import get_range_ref_for_shape
from .base import BaseExhibit


class LossDevelopmentExhibit(BaseExhibit):
    @property
    def controls(self) -> list:
        pass

    _ws_name = 'LossDevelopment'

    def __init__(self, data: cl.Triangle) -> None:
        super().__init__(data)
        if not self.data.is_cumulative:
            self.data = self.data.incr_to_cum()
        self._validations = self.data.index.to_dict(orient='list')

    def load_into(self, wb: DispatchBaseClass):
        try:
            ws = wb.Worksheets(self._ws_name)
        except pywintypes.com_error:
            ws = wb.Worksheets.Add()
            ws.Name = self._ws_name
        self.add_segment_dropdowns(ws)
        breakpoint()
        ws.Range(get_range_ref_for_shape(*self.data.shape[-2:], first_row=5)).Value = (
            self.data.iloc[0, 0].to_frame().fillna('=NA()').values
        )
        return wb

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
