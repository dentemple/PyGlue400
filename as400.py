import win32com.client
from const import SEARCH_FORWARD, NO_TIMEOUT


class AS400:
    def __init__(self):
        self.ConnList = win32com.client.Dispatch("PCOMM.autECLConnList")
        self.Session = win32com.client.Dispatch("PCOMM.autECLSession")
        self.Metrics = win32com.client.Dispatch("PCOMM.autECLWinMetrics")
        self.Operator = win32com.client.Dispatch("PCOMM.autECLOIA")
        self.Presentation = None

    def refresh(self):
        self.ConnList.Refresh

    def set_connection(self, name):
        self.Session.SetConnectionByName(name)
        self.Presentation = self.Session.autECLPS
        self.refresh()

    # --- ConnList methods ---
    def return_connection_count(self):
        self.refresh()
        return self.ConnList.Count

    def return_connection_name(self, connection_number=1):
        return self.ConnList(connection_number).Name

    def return_connection_handle(self, connection_number=1):
        self.refresh()
        return self.ConnList(connection_number).Handle

    def return_connection_type(self, connection_number=1):
        self.refresh()
        return self.ConnList(self, connection_number).ConnType

    # --- Session methods ---
    def is_started(self):
        return self.Session.Started

    def is_ready(self):
        return self.Session.Ready

    # --- Presentation methods ---
    def wait(self, milliseconds):
        self.Presentation.Wait(milliseconds)

    def set_cursor(self, row, col):
        self.Presentation.SetCursorPos(row, col)

    def set_text(self, row, col, text):
        self.Presentation.SetText(text, row, col)

    def send_keys(self, text, row=None, col=None):
        if row and col:
            self.Presentation.SendKeys(text, row, col)
        else:
            self.Presentation.SendKeys(text)

    def search_text(self, find, row=None, col=None, direction=SEARCH_FORWARD):
        if row and col:
            return self.Presentation.SearchText(find, direction, row, col)
        else:
            return self.Presentation.SearchText(find, direction)

    # --- Operator methods ---
    def is_inhibited(self):
        return self.Operator.InputInhibited

    def wait_for_input(self, timeout=NO_TIMEOUT):
        self.Operator.WaitForInputReady(timeout)

    def wait_for_app(self, timeout=NO_TIMEOUT):
        self.Operator.WaitForAppAvailable(timeout)

    def pause(self, timeout=NO_TIMEOUT, optional_add_milliseconds=None):
        self.wait_for_app(timeout=timeout)
        self.wait_for_input(timeout=timeout)
        if optional_add_milliseconds:
            self.wait(milliseconds=optional_add_milliseconds)

    def cancel_waits(self):
        self.Operator.CancelWaits
