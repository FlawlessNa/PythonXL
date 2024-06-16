from enum import Enum


class XlProtectedViewCloseReason(Enum):
    """
    Specifies the reason a protected view window was closed.
    - xlProtectedViewCloseEdit: Window was closed when the user clicked the Enable
        Editing button.
    - xlProtectedViewCloseForced: Window was closed because the application shut it down
        forcefully or stopped responding.
    - xlProtectedViewCloseNormal: Window was closed normally.
    """
    xlProtectedViewCloseEdit = 1
    xlProtectedViewCloseForced = 2
    xlProtectedViewCloseNormal = 0
