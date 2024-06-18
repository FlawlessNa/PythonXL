from abc import ABC


# noinspection PyPep8Naming
class QueryTableEvents(ABC):
    """
    Events that can be triggered through interactions with a QueryTable object in Excel.
    """
    def OnAfterRefresh(self, success: bool) -> None:
        """
        Occurs after a query is completed or canceled.
        :param success: True if the query was completed successfully.
        """
        raise NotImplementedError

    def OnBeforeRefresh(self, cancel: bool) -> None:
        """
        Occurs before any refreshes of the query table. This includes refreshes
        resulting from calling the Refresh method, from the user's actions in the
        product, and from opening the workbook containing the query table.
        :param cancel: False when the event occurs. If the event procedure sets this
            argument to True, the refresh doesn't occur when the procedure is finished.
        """
        raise NotImplementedError
