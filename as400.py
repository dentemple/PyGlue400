import win32com.client


class AS400:
    def __init__(self):

        # ----- Names for connecting to the AS400's Host Automation Class Library -----
        _app_name = "PCOMM"
        _connection_automation_object = "%s.autECLConnList" % _app_name
        _session_automation_object = "%s.autECLSession" % _app_name
        _metrics_automation_object = "%s.autECLWinMetrics" % _app_name
        _operator_automation_object = "%s.autECLOIA" % _app_name

        # ----- Setting the connections to the AS400's Host Automation Class Library -----
        self.ConnList = win32com.client.Dispatch(_connection_automation_object)
        self.Session = win32com.client.Dispatch(_session_automation_object)
        self.Metrics = win32com.client.Dispatch(_metrics_automation_object)
        self.Operator = win32com.client.Dispatch(_operator_automation_object)
        self.Presentation = None

        # ----- AS400 magic values -----
        self.DEFAULT_TO_FIRST_SESSION_CREATED = 1
        self.SEARCH_FORWARD = 1
        self.SEARCH_BACKWARD = 2
        self.NO_TIMEOUT = 0
        self.NOT_INHIBITED = 0
        self.INHIBITED_FROM_SYSTEM_WAIT = 1
        self.INHIBITED_FROM_COMMUNICATION_CHECK = 2
        self.INHIBITED_FROM_PROGRAM_CHECK = 3
        self.INHIBITED_FROM_MACHINE_CHECK = 4
        self.INHIBITED_FROM_OTHER = 5

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

    def return_connection_name(self, connection_number=None):
        if connection_number is None:
            connection_number = self.DEFAULT_TO_FIRST_SESSION_CREATED
        return self.ConnList(connection_number).Name

    def return_connection_handle(self, connection_number=None):
        if connection_number is None:
            connection_number = self.DEFAULT_TO_FIRST_SESSION_CREATED
        self.refresh()
        return self.ConnList(connection_number).Handle

    def return_connection_type(self, connection_number=None):
        if connection_number is None:
            connection_number = self.DEFAULT_TO_FIRST_SESSION_CREATED
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

    def wait_for_input(self, timeout=None):
        if timeout is None:
            timeout = self.NO_TIMEOUT
        self.Operator.WaitForInputReady(timeout)

    def wait_for_app(self, timeout=None):
        if timeout is None:
            timeout = self.NO_TIMEOUT
        self.Operator.WaitForAppAvailable(timeout)

    def pause(self, timeout=None, optional_add_milliseconds=None):
        if timeout is None:
            timeout = self.NO_TIMEOUT
        self.wait_for_app(timeout=timeout)
        self.wait_for_input(timeout=timeout)
        if optional_add_milliseconds:
            self.wait(milliseconds=optional_add_milliseconds)

    def cancel_waits(self):
        self.Operator.CancelWaits
