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
class DialogKeywords:
    """Keywords for interacting with dialogs in Windows applications."""

    def __init__(self, app=None):
        self.app = app
        self.dlg = None

    def get_dialog_from_regex(self, title_re):
        """Get the dialog matching the regex from the application."""
        if not self.app:
            raise RuntimeError("No application is currently connected.")
        self.dlg = self.app.window(title_re=title_re)

    def get_dialog(self, title):
        """Get the dialog by its exact title."""
        if not self.app:
            raise RuntimeError("No application is currently connected.")
        self.dlg = self.app.window(title=title)

    def disconnect_from_dialog(self):
        """Disconnect from the current dialog."""
        self.dlg = None

    def get_number_of_children(self):
        """Get the number of child elements in the current dialog."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return len(self.dlg.children())

    def outline_dialog(self):
        """Outline the current dialog."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg.draw_outline()

    def maximize_window(self):
        """Maximize the current dialog window."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg.maximize()

    def minimize_window(self):
        """Minimize the current dialog window."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg.minimize()

    def close_window(self):
        """Close the current dialog window."""
        if self.dlg:
            self.dlg.close()
        self.dlg = None

    def restore_window(self):
        """Restore the current dialog window."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg.restore()

    def get_window_text(self):
        """Get the text of the current dialog window."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg.window_text()

    def set_window_focus(self):
        """Set focus to the current dialog window."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg.set_focus()

    def print_control_identifiers(self):
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg.print_control_identifiers()

    def popup_menu_select(self, menulocation):
        """Select a menu item from a popup menu by its location."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.app.PopupMenu.menu_select(menulocation)
