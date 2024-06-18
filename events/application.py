from abc import ABC
from win32com.client import Dispatch


# noinspection PyPep8Naming
class ExcelApplicationEvents(ABC):
    """
    Events for the Excel Application object.
    """

    def OnAfterCalculate(self) -> None:
        """
        Occurs when all pending refresh activity (both synchronous and asynchronous) and
        all the resultant calculation activities have been completed.
        """
        pass

    def OnNewWorkbook(self, wb: Dispatch) -> None:
        """
        Occurs when a new workbook is created.
        :param wb: The new workbook.
        """
        pass

    def OnProtectedViewWindowActivate(self, pvw: Dispatch) -> None:
        """
        Occurs when a protected view window is activated.
        :param pvw: The protected view window that is being activated.
        """
        pass

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
        pass

    def OnProtectedViewWindowBeforeEdit(self, pvw: Dispatch, cancel: bool) -> None:
        """
        Occurs immediately before editing is enabled on the workbook in the specified
        Protected View window.
        :param pvw: The protected view window that contains the workbook that is enabled
            for editing.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, editing is not enabled on the workbook.
        """
        pass

    def OnProtectedViewWindowDeactivate(self, pvw: Dispatch) -> None:
        """
        Occurs when a protected view window is deactivated.
        :param pvw: The protected view window that is being deactivated.
        """
        pass

    def OnProtectedViewWindowOpen(self, pvw: Dispatch) -> None:
        """
        Occurs when a workbook is opened in a Protected View window.
        :param pvw: The protected view window that is opened.
        """
        pass

    def OnProtectedViewWindowResize(self, pvw: Dispatch) -> None:
        """
        Occurs when a protected view window is resized.
        :param pvw: The protected view window that is being resized.
        """
        pass

    def OnSheetActivate(self, sh: Dispatch) -> None:
        """
        Occurs when any sheet is activated.
        :param sh: The activated sheet Can be a Chart or Worksheet object.
        """
        pass

    def OnSheetBeforeDelete(self, sh: Dispatch) -> None:
        """
        Occurs before any worksheet is deleted.
        :param sh: The worksheet that is deleted. Can be a Chart or Worksheet object.
        """
        pass

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
        pass

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
        pass

    def OnSheetCalculate(self, sh: Dispatch) -> None:
        """
        Occurs after any worksheet is recalculated or after any changed data is plotted
        on a chart.
        :param sh: Chart or Worksheet object.
        """
        pass

    def OnSheetChange(self, sh: Dispatch, target: Dispatch) -> None:
        """
        Occurs when cells in any worksheet are changed by the user or by an external
        link. This event does NOT occur on chart sheets.
        :param sh: The worksheet that is changed. Worksheet object.
        :param target: Range object that represents the cell or cells that changed.
        """
        pass

    def OnSheetDeactivate(self, sh: Dispatch) -> None:
        """
        Occurs when any sheet is deactivated.
        :param sh: The deactivated sheet. Can be a Chart or Worksheet object.
        """
        pass

    def OnSheetFollowHyperlink(
        self, sh: Dispatch, target: Dispatch
    ) -> None:
        """
        Occurs when a hyperlink is clicked in any worksheet.
        :param sh: The worksheet that contains the hyperlink.
        :param target: The hyperlink object.
        """
        pass

    def OnSheetLensGalleryRenderComplete(
        self, sh: str
    ) -> None:
        """
        Occurs after a callout gallery's icon (dynamic and static) have finished
        rendering.
        :param sh: Name of a worksheet.
        """
        pass

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
        pass

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
        pass

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
        pass

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
        pass

    def OnSheetPivotTableUpdate(
        self, sh: Dispatch, target_pivot_table: Dispatch
    ) -> None:
        """
        Occurs after the sheet of the PivotTable report has been updated.
        :param sh: The selected worksheet.
        :param target_pivot_table: The selected PivotTable report.
        """
        pass

    def OnSheetSelectionChange(
        self, sh: Dispatch, target: Dispatch
    ) -> None:
        """
        Occurs when the selection changes on any worksheet (doesn't occur if the
        selection is on a chart sheet).
        :param sh: The worksheet that contains the new selection.
        :param target: The new selected Range object.
        """
        pass

    def OnSheetTableUpdate(
        self, sh: Dispatch, target: Dispatch
    ) -> None:
        """
        Occurs when a table on a worksheet is updated.
        :param sh: The selected worksheet.
        :param target: The TableObject.
        """
        pass

    def OnWindowActivate(self, wb: Dispatch, wn: Dispatch) -> None:
        """
        Occurs when any workbook window is activated.
        :param wb: The workbook displayed in the activated window.
        :param wn: The window that is activated.
        """
        pass

    def OnWindowDeactivate(self, wb: Dispatch, wn: Dispatch) -> None:
        """
        Occurs when any workbook window is deactivated.
        :param wb: The workbook displayed in the deactivated window.
        :param wn: The window that is deactivated.
        """
        pass

    def OnWindowResize(self, wb: Dispatch, wn: Dispatch) -> None:
        """
        Occurs when any workbook window is resized.
        :param wb: The workbook displayed in the resized window.
        :param wn: The window that is resized.
        """
        pass

    def OnWorkbookActivate(self, wb: Dispatch) -> None:
        """
        Occurs when any workbook is activated.
        :param wb: The workbook that is activated.
        """
        pass

    def OnWorkbookAddinInstall(self, wb: Dispatch) -> None:
        """
        Occurs when a workbook is installed as an add-in.
        :param wb: The installed Workbook.
        """
        pass

    def OnWorkbookAddinUninstall(self, wb: Dispatch) -> None:
        """
        Occurs when a workbook is uninstalled as an add-in.
        :param wb: The uninstalled Workbook.
        """
        pass

    def OnWorkbookAfterRemoteChange(self, wb: Dispatch) -> None:
        """
        Occurs after a remote user's edits to the workbook are merged.
        :param wb: The workbook which has been changed by a remote user.
        """
        pass

    def OnWorkbookAfterSave(self, wb: Dispatch, success: bool) -> None:
        """
        Occurs after the workbook is saved.
        :param wb: The workbook that was saved.
        :param success: True if the workbook was saved successfully.
        """
        pass

    def OnWorkbookAfterXmlExport(
        self, wb: Dispatch, map_: Dispatch, url: str, result: bool
    ) -> None:
        """
        Occurs after Microsoft Excel saves or exports XML data from the specified
        workbook.
        Note: XlXmlExportResult can be one of the following constants:
        - xlXmlExportSuccess. The XML data file was successfully exported.
        - xlXmlExportValidationFailed. The contents of the XML data file don't match the
            specified schema map.
        Use the AfterXmlExport event of the Workbook object if you want to perform an
        operation after XML data has been exported from a particular workbook.
        :param wb: The target workbook.
        :param map_: (XmlMap) The XML map that was used to save or export data.
        :param url: The location of the XML file that was exported.
        :param result: (XlXmlExportResult) Indicates the results of the save or export
            operation.
        """
        pass

    def OnWorkbookAfterXmlImport(
        self, wb: Dispatch, map_: Dispatch, is_refresh: bool, result: bool
    ) -> None:
        """
        Occurs after an existing XML data connection is refreshed, or new XML data is
        imported into any open Microsoft Excel workbook.
        Note: XlXmlImportResult can be one of the following constants:
        - xlXmlImportElementsTruncated. The contents of the specified XML data file have
            been truncated because the XML data file is too large for the worksheet.
        - xlXmlImportSuccess. The XML data file was successfully imported.
        - xlXmlImportValidationFailed. The contents of the XML data file don't match
            the specified schema map.
        Use the AfterXmlImport event of the Workbook object if you want to perform an
        operation after XML data has been imported into a particular workbook.
        :param wb: The target workbook.
        :param map_: (XmlMap) The XML map that was used to import data.
        :param is_refresh: True if the event was triggered by refreshing an existing
            connection to XML data; False if a new mapping was created.
        :param result: (XlXmlImportResult) Indicates the results of the refresh or
            import operation.
        """
        pass

    def OnWorkbookBeforeClose(self, wb: Dispatch, cancel: bool) -> None:
        """
        Occurs immediately before any open workbook closes.
        :param wb: The workbook that's being closed.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, the workbook doesn't close when the procedure is finished.
        """
        pass

    def OnWorkbookBeforePrint(self, wb: Dispatch, cancel: bool) -> None:
        """
        Occurs before any open workbook is printed.
        :param wb: The workbook that is being printed.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, the workbook doesn't print when the procedure is finished.
        """
        pass

    def OnWorkbookBeforeRemoteChange(self, wb: Dispatch) -> None:
        """
        Occurs before a remote user's edits to the workbook are merged.
        :param wb: The workbook that has been changed by a remote user.
        """
        pass

    def OnWorkbookBeforeSave(self, wb: Dispatch, save_as_ui: bool, cancel: bool) -> None:
        """
        Occurs before any open workbook is saved.
        :param wb: The workbook that is being saved.
        :param save_as_ui: 	True if the Save As dialog box will be displayed due to
            changes made that need to be saved in the workbook.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, the workbook doesn't save when the procedure is finished.
        """
        pass

    def OnWorkbookBeforeXmlExport(
        self, wb: Dispatch, map_: Dispatch, url: str, cancel: bool
    ) -> None:
        """
        Occurs before Microsoft Excel saves or exports XML data from the specified
        workbook.
        :param wb: The target workbook.
        :param map_: (XmlMap) The XML map that was used to save or export data.
        :param url: The location of the XML file that was exported.
        :param cancel: Set to True to cancel the save or export operation.
        """
        pass

    def OnWorkbookBeforeXmlImport(
        self, wb: Dispatch, map_: Dispatch, url: str, is_refresh: bool, cancel: bool
    ) -> None:
        """
        Occurs before an existing XML data connection is refreshed, or new XML data is
        imported into any open Microsoft Excel workbook.
        :param wb: The target workbook.
        :param map_: (XmlMap) The XML map that was used to import data.
        :param url: The location of the XML file that was imported.
        :param is_refresh: True if the event was triggered by refreshing an existing
            connection to XML data; False if a new mapping was created.
        :param cancel: Set to True to cancel the import or refresh operation.
        """
        pass

    def OnWorkbookDeactivate(self, wb: Dispatch) -> None:
        """
        Occurs when any workbook is deactivated.
        :param wb: The workbook that is deactivated.
        """
        pass

    def OnWorkbookModelChange(self, wb: Dispatch, changes) -> None:
        """
        Occurs when the data model is updated.
        :param wb: The workbook that contains the data model.
        :param changes: The changes to the data model.
        """
        pass

    def OnWorkbookNewChart(self, wb: Dispatch, ch: Dispatch) -> None:
        """
        Occurs when a new chart is created in any open workbook.
        Note: the WorkbookNewChart event occurs when a new chart is inserted or pasted
        on a worksheet, a chart sheet, or other sheet types. If multiple charts are
        inserted or pasted, the event will occur for each chart in the order they are
        inserted or pasted.
        If a chart object or chart sheet is moved from one location to another, the
        event will not occur. However, if the chart is moved between a chart object and
        a chart sheet, the event will occur because a new chart must be created.
        The WorkbookNewChart event will not occur in the following scenarios: copying or
        pasting a chart sheet, changing a chart type, changing a chart data source,
        undoing or redoing inserting or pasting a chart, and loading a workbook that
        contains a chart.
        :param wb: The workbook that contains the new chart.
        :param ch: The new chart.
        """
        pass

    def OnWorkbookNewSheet(self, wb: Dispatch, sh: Dispatch) -> None:
        """
        Occurs when a new sheet is created in any open workbook.
        :param wb: The workbook that contains the new sheet.
        :param sh: The new sheet.
        """
        pass

    def OnWorkbookOpen(self, wb: Dispatch) -> None:
        """
        Occurs when a workbook is opened.
        :param wb: The workbook that is opened.
        """
        pass

    def OnWorkbookPivotTableCloseConnection(
        self, wb: Dispatch, target: Dispatch
    ) -> None:
        """
        Occurs after a PivotTable report connection has been closed.
        :param wb: The workbook that contains the PivotTable report.
        :param target: The selected PivotTable report.
        """
        pass

    def OnWorkbookPivotTableOpenConnection(
        self, wb: Dispatch, target: Dispatch
    ) -> None:
        """
        Occurs after a PivotTable report connection has been opened.
        :param wb: The workbook that contains the PivotTable report.
        :param target: The selected PivotTable report.
        """
        pass

    def OnWorkbookRowsetComplete(
        self, wb: Dispatch, description: str, sheet: str, success: bool
    ) -> None:
        """
        Occurs when the user either drills through the recordset or invokes the rowset
        action on an OLAP PivotTable.
        Note: Because the recordset is created asynchronously, the event allows
        automation to determine when the action has been completed. Additionally,
        because the recordset is created on a separate sheet, the event needs to be on
        the workbook level.
        :param wb: The workbook that contains the query table.
        :param description: A brief description of the event.
        :param sheet: The name of the worksheet on which the recordset is created.
        :param success: Contains a Boolean value to indicate success or failure.
        """
        pass
