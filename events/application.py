from abc import ABC
from win32com.client import Dispatch


class ExcelApplicationEvents(ABC):
    def OnAfterCalculate(self) -> None:
        """
        Occurs when all pending refresh activity (both synchronous and asynchronous) and
        all the resultant calculation activities have been completed.
        """
        raise NotImplementedError

    def OnNewWorkbook(self, wb: Dispatch) -> None:
        """
        Occurs when a new workbook is created.
        :param wb: The new workbook.
        """
        raise NotImplementedError

    def OnProtectedViewWindowActivate(self, pvw: Dispatch) -> None:
        """
        Occurs when a protected view window is activated.
        :param pvw: The protected view window that is being activated.
        """
        raise NotImplementedError

    def OnProtectedViewWindowBeforeClose(
        self,
        pvw: Dispatch,
        reason: int,
        cancel: bool
    ) -> None:
        """
        Occurs before a protected view window is closed.
        :param pvw: The protected view window that is being closed.
        :param reason: The reason the window is being closed.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, the window does not close when the procedure is finished.
        """
        raise NotImplementedError

    def OnProtectedViewWindowBeforeEdit(self, pvw: Dispatch, cancel: bool) -> None:
        """
        Occurs immediately before editing is enabled on the workbook in the specified
        Protected View window.
        :param pvw: The protected view window that contains the workbook that is enabled
            for editing.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, editing is not enabled on the workbook.
        """
        raise NotImplementedError

    def OnProtectedViewWindowDeactivate(self, pvw: Dispatch) -> None:
        """
        Occurs when a protected view window is deactivated.
        :param pvw: The protected view window that is being deactivated.
        """
        raise NotImplementedError

    def OnProtectedViewWindowOpen(self, pvw: Dispatch) -> None:
        """
        Occurs when a workbook is opened in a Protected View window.
        :param pvw: The protected view window that is opened.
        """
        raise NotImplementedError

    def OnProtectedViewWindowResize(self, pvw: Dispatch) -> None:
        """
        Occurs when a protected view window is resized.
        :param pvw: The protected view window that is being resized.
        """
        raise NotImplementedError

    def OnSheetActivate(self, sh: Dispatch) -> None:
        """
        Occurs when any sheet is activated.
        :param sh: The activated sheet Can be a Chart or Worksheet object.
        """
        raise NotImplementedError

    def OnSheetBeforeDelete(self, sh: Dispatch) -> None:
        """
        Occurs before any worksheet is deleted.
        :param sh: The worksheet that is deleted. Can be a Chart or Worksheet object.
        """
        raise NotImplementedError

    def OnSheetBeforeDoubleClick(
        self, sh: Dispatch, target: Dispatch, cancel: bool
    ) -> None:
        """
        Occurs when any worksheet is double-clicked, before the default double-click
        action. This event does NOT occur on chart sheets.
        :param sh: The worksheet that is double-clicked. Worksheet object.
        :param target: The cell nearest to the mouse pointer when the double-click
            occurred.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, the default double-click action is not performed when the
            procedure is finished.
        """
        raise NotImplementedError

    def OnSheetBeforeRightClick(
        self, sh: Dispatch, target: Dispatch, cancel: bool
    ) -> None:
        """
        Occurs when any worksheet is right-clicked, before the default right-click
        action. This event does NOT occur on chart sheets.
        :param sh: The worksheet that is right-clicked. Worksheet object.
        :param target: The cell nearest to the mouse pointer when the right-click
            occurred.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, the default right-click action is not performed when the
            procedure is finished.
        """
        raise NotImplementedError

    def OnSheetCalculate(self, sh: Dispatch) -> None:
        """
        Occurs after any worksheet is recalculated or after any changed data is plotted
        on a chart.
        :param sh: Chart or Worksheet object.
        """
        raise NotImplementedError

    def OnSheetChange(self, sh: Dispatch, target: Dispatch) -> None:
        """
        Occurs when cells in any worksheet are changed by the user or by an external
        link. This event does NOT occur on chart sheets.
        :param sh: The worksheet that is changed. Worksheet object.
        :param target: Range object that represents the cell or cells that changed.
        """
        raise NotImplementedError

    def OnSheetDeactivate(self, sh: Dispatch) -> None:
        """
        Occurs when any sheet is deactivated.
        :param sh: The deactivated sheet. Can be a Chart or Worksheet object.
        """
        raise NotImplementedError

    def OnSheetFollowHyperlink(
        self, sh: Dispatch, target: Dispatch
    ) -> None:
        """
        Occurs when a hyperlink is clicked in any worksheet.
        :param sh: The worksheet that contains the hyperlink.
        :param target: The hyperlink object.
        """
        raise NotImplementedError

    def OnSheetLensGalleryRenderComplete(
        self, sh: str
    ) -> None:
        """
        Occurs after a callout gallery's icon (dynamic and static) have finished
        rendering.
        :param sh: Name of a worksheet.
        """
        raise NotImplementedError

    def OnSheetPivotTableAfterValueChange(
        self, sh: Dispatch, target_pivot_table: Dispatch, target_range: Dispatch
    ) -> None:
        """
        Occurs after a cell or range of cells inside a PivotTable are edited or
        recalculated (for cells that contain formulas).
        Note: The PivotTableAfterValueChange event does not occur under any conditions
        other than editing or recalculating cells. For example, it will not occur when
        the PivotTable is refreshed, sorted, filtered, or drilled down on, even though
        those operations move cells and potentially retrieve new values from the OLAP
        data source.
        :param sh: The worksheet that contains the PivotTable report.
        :param target_pivot_table: The PivotTable that contains the edited or
            recalculated cells.
        :param target_range: The Range of cells that were edited or recalculated.
        """
        raise NotImplementedError

    def OnSheetPivotTableBeforeAllocateChanges(
        self,
        sh: Dispatch,
        target_pivot_table: Dispatch,
        value_change_start: int,
        value_change_end: int,
        cancel: bool
    ) -> None:
        """
        Occurs before changes are applied to PivotTable.
        Note: The SheetPivotTableBeforeAllocateChanges event occurs immediately before
        Excel executes an UPDATE CUBE statement to apply all changes to the PivotTable's
        OLAP data source, and immediately after the user has chosen to apply changes in
        the user interface.
        :param sh: The worksheet that contains the PivotTable report.
        :param target_pivot_table: The PivotTable containing changes to apply.
        :param value_change_start: The index to the first change in the associated
            PivotTableChangeList collection. The index is specified by the Order property
            of the ValueChange object in the PivotTableChangeList collection.
        :param value_change_end: The index to the last change in the associated
            PivotTableChangeList collection. The index is specified by the Order property
            of the ValueChange object in the PivotTableChangeList collection.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, the changes are not applied to the PivotTable,
            and all edits are lost.
        """
        raise NotImplementedError

    def OnSheetPivotTableBeforeCommitChanges(
        self,
        sh: Dispatch,
        target_pivot_table: Dispatch,
        value_change_start: int,
        value_change_end: int,
        cancel: bool
    ) -> None:
        """
        Occurs before changes are committed against the OLAP data source for a PivotTable
        The SheetPivotTableBeforeCommitChanges event occurs immediately before Excel
        executes a COMMIT TRANSACTION against the PivotTable's OLAP data source, and
        immediately after the user has chosen to save changes for the entire PivotTable.
        :param sh: The worksheet that contains the PivotTable report.
        :param target_pivot_table: The PivotTable containing changes to commit.
        :param value_change_start: The index to the first change in the associated
            PivotTableChangeList collection. The index is specified by the Order property
            of the ValueChange object in the PivotTableChangeList collection.
        :param value_change_end: The index to the last change in the associated
            PivotTableChangeList collection. The index is specified by the Order property
            of the ValueChange object in the PivotTableChangeList collection.
        :param cancel: False when the event occurs. If the event procedure sets this
        argument to True, the changes are not committed against the OLAP data source of
        the PivotTable.
        """
        raise NotImplementedError

    def OnSheetPivotTableBeforeDiscardChanges(
        self,
        sh: Dispatch,
        target_pivot_table: Dispatch,
        value_change_start: int,
        value_change_end: int
    ) -> None:
        """
        Occurs before changes are discarded from the PivotTable.
        Occurs immediately before Excel executes a ROLLBACK TRANSACTION statement against
        the OLAP data source, if a transaction is still active, and then discards all
        edited values in the PivotTable after the user has chosen to discard changes.
        :param sh: The worksheet that contains the PivotTable report.
        :param target_pivot_table: The PivotTable containing changes to discard.
        :param value_change_start: The index to the first change in the associated
            PivotTableChangeList collection. The index is specified by the Order property
            of the ValueChange object in the PivotTableChangeList collection.
        :param value_change_end: The index to the last change in the associated
            PivotTableChangeList collection. The index is specified by the Order property
            of the ValueChange object in the PivotTableChangeList collection.
        """
        raise NotImplementedError

    def OnSheetPivotTableUpdate(
        self, sh: Dispatch, target_pivot_table: Dispatch
    ) -> None:
        """
        Occurs after the sheet of the PivotTable report has been updated.
        :param sh: The selected worksheet.
        :param target_pivot_table: The selected PivotTable report.
        """
        raise NotImplementedError

    def OnSheetSelectionChange(
        self, sh: Dispatch, target: Dispatch
    ) -> None:
        """
        Occurs when the selection changes on any worksheet (doesn't occur if the
        selection is on a chart sheet).
        :param sh: The worksheet that contains the new selection.
        :param target: The new selected Range object.
        """
        raise NotImplementedError

    def OnSheetTableUpdate(
        self, sh: Dispatch, target: Dispatch
    ) -> None:
        """
        Occurs when a table on a worksheet is updated.
        :param sh: The selected worksheet.
        :param target: The TableObject.
        """
        raise NotImplementedError