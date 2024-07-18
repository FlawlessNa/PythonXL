from abc import ABC


# noinspection PyPep8Naming
class OLEObjectEvents(ABC):
    """
    Events that can be triggered through interactions with an OLEObject object in Excel.
    """
    def OnAddControl(self) -> None:
        """
        Occurs when a control is inserted onto a form, a Frame, or a Page of a
        MultiPage.
        :return:
        """
        pass

    def OnAddRef(self) -> None:
        pass

    def OnAfterUpdate(self) -> None:
        """
        Occurs after data in a control is changed through the user interface.
        :return:
        """
        pass

    def OnBeforeDragOver(self, *args, **kwargs) -> None:
        """
        Occurs when a drag-and-drop operation is in progress.
        """
        pass

    def OnBeforeDropOrPaste(self, *args, **kwargs) -> None:
        """
        Occurs when the user is about to drop or paste data onto an object.
        """
        pass

    def OnBeforeUpdate(self, cancel: bool) -> None:
        """
        Occurs before data in a control is changed.
        :param cancel: Required. Event status. False indicates that the control should
         handle the event (default). True cancels the update and indicates the
         application should handle the event.
        :return:
        """
        pass

    def OnChange(self) -> None:
        """
        Occurs when the Value property changes.
        """
        pass

    def OnClick(self, *args, **kwargs) -> None:
        """
        Occurs in one of two cases:
            The user clicks a control with the mouse.
            The user definitively selects a value for a control with more than one
            possible value.
        """
        pass

    def OnDblClick(self, *args, **kwargs) -> None:
        """
        Occurs when the user points to an object and then clicks a mouse button twice.
        """
        pass

    def OnDropButtonClick(self) -> None:
        """
        Occurs whenever the drop-down list appears or disappears.
        Note: You can initiate the DropButtonClick event through code or by taking
            certain actions in the user interface.
        In code, calling the DropDown method initiates the DropButtonClick event.
        In the user interface, any of the following actions initiates the event:
            Clicking the drop-down button on the control.
            Pressing F4.
        Any of the previous actions, in code or in the interface, cause the drop-down
        box to appear on the control. The system initiates the DropButtonClick event
        when the drop-down box goes away.
        """
        pass

    def OnError(self, *args, **kwargs) -> None:
        """
        Occurs when a control detects an error and cannot return the error information
        to a calling program.
        """
        pass

    def OnGetIDsOfNames(self, *args, **kwargs) -> None:
        pass

    def OnGetTypeInfo(self, *args, **kwargs) -> None:
        pass

    def OnGetTypeInfoCount(self, *args, **kwargs) -> None:
        pass

    def OnGotFocus(self) -> None:
        """
        Occurs when an ActiveX control gets input focus.
        """
        pass

    def OnKeyDown(self, *args, **kwargs) -> None:
        """
        Occur in sequence when a user presses and releases a key. KeyDown occurs when
        the user presses a key. KeyUp occurs when the user releases a key.
        """
        pass
    def OnKeyUp(self, *args, **kwargs) -> None:
        """
        Occur in sequence when a user presses and releases a key. KeyDown occurs when
        the user presses a key. KeyUp occurs when the user releases a key.
        """
        pass

    def OnKeyPress(self, key_ansi: int) -> None:
        """
        Occurs when the user presses an ANSI key.
        :param key_ansi: Required. An integer value that represents a standard numeric
        ANSI key code.
        :return:
        """
        pass

    def OnInvoke(self, *args, **kwargs) -> None:
        pass

    def OnLayout(self, index: int) -> None:
        """
        Occurs when a form, Frame, or MultiPage changes size.
        :param index: Required. The index of the page in a MultiPage that changed size.
        :return:
        """
        pass

    def OnMouseDown(self, *args, **kwargs) -> None:
        """
        Occur when the user clicks a mouse button. MouseDown occurs when the user
        presses the mouse button; MouseUp occurs when the user releases the mouse button.
        """
        pass

    def OnMouseUp(self, *args, **kwargs) -> None:
        """
        Occur when the user clicks a mouse button. MouseDown occurs when the user
        presses the mouse button; MouseUp occurs when the user releases the mouse button.
        """
        pass

    def OnMouseMove(self, *args, **kwargs) -> None:
        """
        Occurs when the user moves the mouse.
        """
        pass

    def OnQueryInterface(self, *args, **kwargs):
        pass

    def OnRelease(self):
        pass

    def OnRemoveControl(self, *args, **kwargs) -> None:
        """
        Occurs when a control is deleted from the container.
        """
        pass

    def OnScroll(self, *args, **kwargs) -> None:
        """
        Occurs when the scroll box is repositioned.
        """
        pass

    def OnSpinDown(self) -> None:
        """
        SpinDown occurs when the user clicks the lower or left spin-button arrow.
        SpinUp occurs when the user clicks the upper or right spin-button arrow.
        """
        pass

    def OnSpinUp(self) -> None:
        """
        SpinDown occurs when the user clicks the lower or left spin-button arrow.
        SpinUp occurs when the user clicks the upper or right spin-button arrow.
        """
        pass

    def OnZoom(self, *args, **kwargs) -> None:
        """
        Occurs when the value of the Zoom property changes.
        """
        pass

    def OnLostFocus(self) -> None:
        """
        Occurs when an ActiveX control loses input focus.
        """
        pass
