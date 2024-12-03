# MIT License
# Copyright (c) 2024 Anoop G R
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
from pywinauto import Application
import pywinauto.timings


class ApplicationKeywords:
    """Keywords for interacting with Windows applications."""

    def __init__(self):
        self.app = None

    def set_timeout(self, timeout_type, new_timeout_time):
        """Set timeout values for pywinauto."""
        timeout_types = ["window_find_timeout", "window_find_retry", "app_start_timeout", "app_start_retry",
                         "exists_timeout",
                         "exists_retry", "after_click_wait", "after_clickinput_wait", "after_menu_wait",
                         "after_sendkeys_key_wait",
                         "after_button_click_wait", "before_closeclick_wait", "closeclick_retry",
                         "closeclick_dialog_close_wait",
                         "after_closeclick_wait", "after_windowclose_timeout", "after_windowclose_retry",
                         "after_setfocus_wait",
                         "after_setcursorpos_wait", "sendmessagetimeout_timeout", "after_tabselect_wait",
                         "after_listviewselect_wait",
                         "after_listviewcheck_wait", "after_treeviewselect_wait", "after_toobarpressbutton_wait",
                         "after_updownchange_wait",
                         "after_movewindow_wait", "after_buttoncheck_wait", "after_comboselect_wait",
                         "after_listboxselect_wait",
                         "after_listboxfocuschange_wait", "after_editsetedittext_wait", "after_editselect_wait",
                         "default"]

        if timeout_type not in timeout_types:
            raise ValueError(f"{timeout_type} is not one of the valid timeout types.")

        if timeout_type == "default":
            if new_timeout_time == "default":
                pywinauto.timings.Timings.Defaults()
            elif new_timeout_time == "fast":
                pywinauto.timings.Timings.Fast()
            elif new_timeout_time == "slow":
                pywinauto.timings.Timings.Slow()
            else:
                raise ValueError(f'{new_timeout_time} must be "fast", "slow", or "default"')
        else:
            new_timeout_time = float(new_timeout_time)
            setattr(pywinauto.timings.Timings, timeout_type, new_timeout_time)

    def launch_application(self, app_path, backend="win32"):
        """Launch a Windows application."""
        self.app = Application(backend).start(app_path)
        return self.app

    def connect_to_application(self, title_regex, backend="win32"):
        """Connect to a running application using a window title regex."""
        self.app = Application(backend).connect(title_re=title_regex)
        return self.app

    def close_application(self):
        """Close the current application."""
        if self.app:
            self.app.kill()
        self.app = None

    def disconnect_from_application(self):
        """Disconnect from the current application."""
        self.app = None
