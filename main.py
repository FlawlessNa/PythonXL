import atexit
import chainladder as cl
import os
import pythoncom
from win32com.client import constants as c, gencache

from excel.utilities import get_range_ref_for_shape
from exhibits import LossDevelopmentExhibit


def cleanup(wb, app):
    if wb is not None:
        wb.Close()
        # app.Quit()
    app.DisplayAlerts = True


if __name__ == '__main__':
    data = cl.load_sample('prism').groupby(['Line', 'Type']).sum()
    # app = win32com.client.DispatchWithEvents("Excel.Application", ExcelApplicationEvents)
    ldf_exhibit = LossDevelopmentExhibit(data)

    app = gencache.EnsureDispatch("Excel.Application")
    app.Visible = True
    app.DisplayAlerts = False
    wb = gencache.EnsureDispatch(app.Workbooks.Add())
    atexit.register(cleanup, wb, app)
    ldf_exhibit.load_into(wb)
    # breakpoint()
    app.CommandBars.ExecuteMso('DesignMode')
    app.CommandBars.ExecuteMso('DesignMode')
    while True:
        pythoncom.PumpWaitingMessages()

