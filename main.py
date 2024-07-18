import atexit
import chainladder as cl
import os
import win32com.client
from win32com.client import constants as c

from events import ExcelApplicationEvents
from excel.utilities import get_range_ref_for_shape
from exhibits import LossDevelopmentExhibit


def cleanup(app):
    if app is not None:
        app.Quit()
        app.DisplayAlerts = True


if __name__ == '__main__':
    win32com.client.gencache.EnsureDispatch("Excel.Application")
    # try:
    import time
    start = time.perf_counter()
    data = cl.load_sample('prism').groupby(['Line', 'Type']).sum()
    print(f"Data loaded in {time.perf_counter() - start:.2f}s")
    # app = win32com.client.DispatchWithEvents("Excel.Application", ExcelApplicationEvents)
    ldf_exhibit = LossDevelopmentExhibit(data)

    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = True
    app.DisplayAlerts = False
    atexit.register(cleanup, app)
    wb = win32com.client.Dispatch(app.Workbooks.Add())
    ldf_exhibit.load_into(wb)

