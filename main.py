import chainladder as cl
import os
import win32com.client
from win32com.client import constants as c

from events import ExcelApplicationEvents
from excel.utilities import get_range_ref_for_shape
from exhibits import LossDevelopmentExhibit

RUN_ID = 'dev'
WB_ID = 'test.xlsx'

if __name__ == '__main__':
    # try:
    data = cl.load_sample('prism').groupby(['Line', 'Type']).sum()
    # app = win32com.client.DispatchWithEvents("Excel.Application", ExcelApplicationEvents)
    app = win32com.client.Dispatch("Excel.Application")
    if os.path.exists(f'output/{RUN_ID}'):
        wb = win32com.client.Dispatch(app.Workbooks.Open(f'output/{RUN_ID}/{WB_ID}'))
    else:
        wb = win32com.client.Dispatch(app.Workbooks.Add())
    ldf_exhibit = LossDevelopmentExhibit(data)
    wb = ldf_exhibit.load_into(wb)
    breakpoint()
    # finally:
    #     os.remove(f'output/{RUN_ID}')
