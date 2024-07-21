import chainladder as cl
import math
import pywintypes
import win32com.client
from win32com.client import DispatchBaseClass, constants as c, gencache
from excel.utilities import get_range_ref_for_shape
from exhibits.base import BaseExhibit, BaseComponent
from ._events import ListBoxEventHandler


class LossDevelopmentExhibit:

    _ws_name = 'LossDevelopment'

    @property
    def components(self) -> list[BaseComponent]:
        return []

    def __init__(self, data: cl.Triangle) -> None:
        self.data = data
        if not self.data.is_cumulative:
            self.data = self.data.incr_to_cum()
        self._amounts_filter = self._index_filters = self._ws = None
        self._ready = False

    def load_into(self, wb: DispatchBaseClass):
        if self._ws_name in (ws.Name for ws in wb.Worksheets):
            self._ws = gencache.EnsureDispatch(wb.Worksheets(self._ws_name))
        else:
            self._ws = gencache.EnsureDispatch(wb.Worksheets.Add())
            self._ws.Name = self._ws_name
        self._index_filters = self._add_dropdowns(self.data.index.to_dict(orient='list'))
        self._amounts_filter = self._add_dropdowns(
            {'Amount': self.data.columns.to_list()},
            offset=math.prod(self.data.index.shape)
        ).pop()
        self._refresh_data()
        self._ready = True
        return wb

    def _add_dropdowns(self, filters: dict, offset: int = 0) -> list:
        dropdowns = []
        longest_val = max(
            [item for sublist in filters for item in sublist],
            key=len
        )
        self._ws.Cells(1, 1).Value = f"{'x' * max(30, len(longest_val))}"
        self._ws.Columns(1).AutoFit()
        self._ws.Cells(1, 1).ClearContents()

        num_values = offset
        for segment, values in filters.items():
            ole = gencache.EnsureDispatch(
                self._ws.OLEObjects().Add("Forms.ListBox.1")
            )
            cell = self._ws.Cells(num_values + 1, 1)
            ole.Left, ole.Top, ole.Width = cell.Left, cell.Top, cell.Width
            ole.Height = cell.Height * len(values)
            ole.Placement = c.xlMoveAndSize
            # event_handler = ListBoxEventHandler(self)
            # No Need to EnsureDispatch when using DispatchWithEvents, already done behind scene
            active_x = win32com.client.DispatchWithEvents(ole.Object, ListBoxEventHandler)
            print(hasattr(active_x, 'exhibit'))
            active_x.exhibit = self
            active_x.FontBold = True
            active_x.List = values
            active_x.Value = values[0]
            active_x.Name = f'{segment}_dropdown'
            dropdowns.append(active_x)
            # active_x.BackColor = ...
            # active_x.ForeColor = ...

            num_values += len(values)
        return dropdowns

    def _refresh_data(self):
        tri = self.data.loc[
            tuple(dropdown.Value for dropdown in self._index_filters),
            self._amounts_filter.Value
        ]
        self._ws.Range(
            get_range_ref_for_shape(len(tri.origin), 1, first_row=3, first_col='C')
        ).Value = tri.origin.astype(str).values.reshape((len(tri.origin), 1))
        self._ws.Range(
            get_range_ref_for_shape(1, len(tri.origin), first_row=2, first_col='D')
        ).Value = tri.development.values
        self._ws.Range(
            get_range_ref_for_shape(*self.data.shape[-2:], first_row=3, first_col='D')
        ).Value = tri.to_frame().fillna('=NA()').values
