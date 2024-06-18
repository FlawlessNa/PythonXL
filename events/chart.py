from abc import ABC


# noinspection PyPep8Naming
class ExcelChartEvents(ABC):
    """
    Events that can be triggered through interactions with a Chart object in Excel.
    """

    def OnActivate(self) -> None:
        """
        Occurs when a workbook, worksheet, chart sheet, or embedded chart is activated.
        Note: This event doesn't occur when you create a new window.
        When you switch between two windows showing the same workbook, the
        WindowActivate event occurs, but the Activate event for the workbook doesn't
        occur.
        """
        raise NotImplementedError

    def OnBeforeDoubleClick(
        self, cancel: bool, arg1: int, arg2: int, element_id: int
    ) -> None:
        """
        Occurs when a chart element is double-clicked, before the default double-click
        action.
        Note: The DoubleClick method doesn't cause this event to occur.
        This event doesn't occur when the user double-clicks the border of a cell.
        The meaning of Arg1 and Arg2 depends on the ElementID value.
        https://learn.microsoft.com/en-us/office/vba/api/excel.chart.beforedoubleclick
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, the default double-click action isn't performed when the
            procedure is finished.
        :param arg1: Additional event information, depending on the value of ElementID.
            For more information about this parameter, see the Remarks section.
        :param arg2: Additional event information, depending on the value of ElementID.
            For more information about this parameter, see the Remarks section.
        :param element_id: The double-clicked object. The value of this parameter
        determines the expected values of Arg1 and Arg2. For more information about this
        parameter, see the Remarks section.
        """
        raise NotImplementedError

    def OnBeforeRightClick(self, cancel: bool) -> None:
        """
        Occurs when a chart element is right-clicked, before the default right-click
        action.
        Note: Like other worksheet events, this event doesn't occur if you right-click
        while the pointer is on a shape or a command bar (a toolbar or menu bar).
        Note: The RightClick method doesn't cause this event to occur.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, the default right-click action isn't performed when the
            procedure is finished.
        """
        raise NotImplementedError

    def OnCalculate(self) -> None:
        """
        Occurs after the chart plots new or changed data for the Chart object.
        """
        raise NotImplementedError

    def OnDeactivate(self) -> None:
        """
        Occurs when the chart, worksheet, or workbook is deactivated.
        """
        raise NotImplementedError

    def OnMouseDown(self, button: int, shift: int, x: int, y: int) -> None:
        """
        Occurs when a mouse button is pressed while the pointer is over a chart.
        Note: The following table specifies the values for the Shift parameter.
            0 - No keys
            1 - Shift key
            2 - Ctrl key
            4 - Alt key
        :param button: The mouse button that was released. Can be one of the following
            XlMouseButton constants: xlNoButton, xlPrimaryButton, or xlSecondaryButton.
        :param shift: The state of the Shift, Ctrl, and AlShift, Ctrl, and AlttShift,
            Ctrl, and Alt keys when the event occurred. Can be one of or a sum of
            values.
        :param x: The x coordinate of the mouse pointer in chart object client
            coordinates.
        :param y: The y coordinate of the mouse pointer in chart object client
            coordinates.
        """
        raise NotImplementedError

    def OnMouseMove(self, button: int, shift: int, x: int, y: int) -> None:
        """
        Occurs when the position of the mouse pointer changes over a chart.
        Note: The following table specifies the values for the Shift parameter.
            0 - No keys
            1 - Shift key
            2 - Ctrl key
            4 - Alt key
        :param button: The mouse button that was released. Can be one of the following
            XlMouseButton constants: xlNoButton, xlPrimaryButton, or xlSecondaryButton.
        :param shift: The state of the Shift, Ctrl, and Alt keys when the event
            occurred. Can be one of or a sum of values.
        :param x: The x coordinate of the mouse pointer in chart object client
            coordinates.
        :param y: The y coordinate of the mouse pointer in chart object client
            coordinates.
        """
        raise NotImplementedError

    def OnMouseUp(self, button: int, shift: int, x: int, y: int) -> None:
        """
        Occurs when a mouse button is released while the pointer is over a chart.
        Note: The following table specifies the values for the Shift parameter.
            0 - No keys
            1 - Shift key
            2 - Ctrl key
            4 - Alt key
        :param button: The mouse button that was released. Can be one of the following
            XlMouseButton constants: xlNoButton, xlPrimaryButton, or xlSecondaryButton.
        :param shift: The state of the Shift, Ctrl, and Alt keys when the event
            occurred. Can be one of or a sum of values.
        :param x: The x coordinate of the mouse pointer in chart object client
            coordinates.
        :param y: The y coordinate of the mouse pointer in chart object client
            coordinates.
        """
        raise NotImplementedError

    def OnResize(self) -> None:
        """
        Occurs when the chart is resized.
        """
        raise NotImplementedError

    def OnSelect(self, element_id: int, arg1: int, arg2: int) -> None:
        """
        Occurs when a chart element is selected.
        :param element_id: The selected chart element. For more information about this
            argument, see the OnBeforeDoubleClick event.
        :param arg1: The selected chart element. For more information about this
            argument, see the OnBeforeDoubleClick event.
        :param arg2: The selected chart element. For more information about this
            argument, see the OnBeforeDoubleClick event.
        """
        raise NotImplementedError

    def OnSeriesChange(self, series_index: int, point_index: int) -> None:
        """
        Occurs when the user changes the value of a chart data point by choosing a bar
        in the chart and dragging the top edge up or down thus changing the value of the
        data point.
        Note: This event is NOT functional in Excel 2007 and later versions.
        :param series_index: The offset within the Series collection for the changed
            series.
        :param point_index: The offset within the Points collection for the changed
            point.
        """
        raise NotImplementedError
