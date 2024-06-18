import win32com.client
from win32com.client import constants as c
from events import ExcelApplicationEvents

if __name__ == '__main__':
    app = win32com.client.DispatchWithEvents("Excel.Application", ExcelApplicationEvents)
    wb = app.Workbooks.Add()
    ws = wb.Worksheets[1]
    breakpoint()
    breakpoint()
