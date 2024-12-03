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
from robot.api.deco import keyword
from .keywords import ApplicationKeywords, DialogKeywords, ControlKeywords

__version__ = "1.0.0"


class PywinautoLibrary:
    """PywinautoLibrary is Robot Framework library for interacting with Windows GUI applications."""

    ROBOT_LIBRARY_SCOPE = "GLOBAL"
    ROBOT_LIBRARY_VERSION = __version__

    def __init__(self):
        self.app_keywords = ApplicationKeywords()
        self.dialog_keywords = None
        self.control_keywords = None

    # Application-related keywords
    @keyword
    def set_timeout(self, timeout_type, new_timeout_time):
        """
        Set custom timeout values for pywinauto timings.
        Below are the acceptable timeout types, and their associated timeout defaults.
        Also acceptable are the timeout type of "default" and the timeout time of "fast", "slow" or "default".
        These are predefined timeout templates.

        * window_find_timeout (default 3)
        * window_find_retry (default .09)
        * app_start_timeout (default 10)
        * app_start_retry (default .90)
        * exists_timeout (default .5)
        * exists_retry (default .3)
        * after_click_wait (default .09)
        * after_clickinput_wait (default .01)
        * after_menu_wait (default .05)
        * after_sendkeys_key_wait (default .01)
        * after_button_click_wait (default 0)
        * before_closeclick_wait (default .1)
        * closeclick_retry (default .05)
        * closeclick_dialog_close_wait (default .05)
        * after_closeclick_wait (default .2)
        * after_windowclose_timeout (default 2)
        * after_windowclose_retry (default .5)
        * after_setfocus_wait (default .06)
        * after_setcursorpos_wait (default .01)
        * sendmessagetimeout_timeout (default .001)
        * after_tabselect_wait (default .01)
        * after_listviewselect_wait (default .01)
        * after_listviewcheck_wait default(.001)
        * after_treeviewselect_wait default(.001)
        * after_toobarpressbutton_wait default(.01)
        * after_updownchange_wait default(.001)
        * after_movewindow_wait default(0)
        * after_buttoncheck_wait default(0)
        * after_comboselect_wait default(0)
        * after_listboxselect_wait default(0)
        * after_listboxfocuschange_wait default(0)
        * after_editsetedittext_wait default(0)
        * after_editselect_wait default(0)
        """
        self.app_keywords.set_timeout(timeout_type, new_timeout_time)

    @keyword
    def launch_application(self, app_path, backend="win32"):
        """Launch a Windows application."""
        app = self.app_keywords.launch_application(app_path, backend)
        self.dialog_keywords = DialogKeywords(app)  # Inject the application instance
        self.control_keywords = ControlKeywords(self.dialog_keywords.dlg)  # Inject the dialog instance

    @keyword
    def connect_to_application(self, title_regex, backend="win32"):
        """Connect to a running application using a window title regex."""
        app = self.app_keywords.connect_to_application(title_regex, backend)
        self.dialog_keywords = DialogKeywords(app)  # Inject the application instance
        self.control_keywords = ControlKeywords(self.dialog_keywords.dlg)  # Inject the dialog instance

    @keyword
    def close_application(self):
        """Close the current application."""
        self.dialog_keywords.close_window()
        self.app_keywords.close_application()

    @keyword
    def disconnect_from_application(self):
        """Disconnect from the current application."""
        self.dialog_keywords.disconnect_from_dialog()
        self.app_keywords.disconnect_from_application()

    # Dialog-related keywords
    @keyword
    def get_dialog_from_regex(self, title_re):
        """Get the dialog matching the regex from the application."""
        self.dialog_keywords.get_dialog_from_regex(title_re)
        self.control_keywords.set_dialog(self.dialog_keywords.dlg)

    @keyword
    def get_dialog(self, title):
        """Get the dialog by its exact title."""
        self.dialog_keywords.get_dialog(title)
        self.control_keywords.set_dialog(self.dialog_keywords.dlg)

    @keyword
    def get_number_of_children(self):
        """Get the number of child elements in the current dialog."""
        return self.dialog_keywords.get_number_of_children()

    @keyword
    def outline_dialog(self):
        """Outline the current dialog."""
        self.dialog_keywords.outline_dialog()

    @keyword
    def maximize_window(self):
        """Maximize the current dialog window."""
        self.dialog_keywords.maximize_window()

    @keyword
    def minimize_window(self):
        """Minimize the current dialog window."""
        self.dialog_keywords.minimize_window()

    @keyword
    def close_window(self):
        """Close the current dialog window."""
        self.dialog_keywords.close_window()

    @keyword
    def restore_window(self):
        """Restore the current dialog window."""
        self.dialog_keywords.restore_window()

    @keyword
    def get_window_text(self):
        """Get the text of the current dialog window."""
        return self.dialog_keywords.get_window_text()

    @keyword
    def set_window_focus(self):
        """Set focus to the current dialog window."""
        self.dialog_keywords.set_window_focus()

    @keyword
    def print_control_identifiers(self):
        """Print control identifiers of the current dialog."""
        return self.dialog_keywords.print_control_identifiers()

    # Control-related keywords
    @keyword
    def get_control_text(self, control_name):
        """Retrieve the text of a specified control."""
        return self.control_keywords.get_control_text(control_name)

    @keyword
    def menu_select(self, menulocation):
        """Select a menu item by its location (e.g., 'File -> Save')."""
        self.control_keywords.menu_select(menulocation)

    @keyword
    def type_text(self, control_name, text):
        """Type text into a specified control."""
        self.control_keywords.type_text(control_name, text)

    @keyword
    def send_keys(self, keys):
        """Send keyboard input to the current dialog."""
        self.control_keywords.send_keys(keys)

    @keyword
    def click(self, control_name):
        """Click on a specified control in the current dialog."""
        self.control_keywords.click(control_name)

    @keyword
    def real_click(self, control_name):
        """Real click (simulated as physical) on a specified control."""
        self.control_keywords.real_click(control_name)

    @keyword
    def right_click(self, control_name):
        """Right-click on a specified control."""
        self.control_keywords.right_click(control_name)

    @keyword
    def real_right_click(self, control_name):
        """Real right-click (simulated as physical) on a specified control."""
        self.control_keywords.real_right_click(control_name)

    @keyword
    def double_click(self, control_name):
        """Double-click on a specified control."""
        self.control_keywords.double_click(control_name)

    @keyword
    def real_double_click(self, control_name):
        """Real double-click (simulated as physical) on a specified control."""
        self.control_keywords.real_double_click(control_name)

    @keyword
    def drag_mouse(self, control_name, dst, src=None, button='left', pressed='', absolute=True):
        """Click on src, drag it and drop on dst.
        dst is a destination wrapper object or just coordinates.
        src is a source wrapper object or coordinates. If src is None the self is used as a source object.
        button is a mouse button to hold during the drag. It can be “left”, “right”, “middle” or “x”
        pressed is a key on the keyboard to press during the drag.
        absolute specifies whether to use absolute coordinates for the mouse pointer locations"""

        self.control_keywords.drag_mouse(control_name, dst, src, button, pressed, absolute)

    @keyword
    def control_is_active(self, control_name):
        """Verify that the element is both visible and enabled.\
        Raise either ElementNotEnabled or ElementNotVisible if not enabled or visible respectively."""
        self.control_keywords.control_is_active(control_name)

    @keyword
    def control_is_visible(self, control_name):
        """Verify that the element is visible.
        Check first if the element’s parent is visible (skip if no parent), then check if element itself is visible."""
        self.control_keywords.control_is_visible(control_name)

    @keyword
    def control_is_enabled(self, control_name):
        """Verify that the element is enabled.
        Check first if the element’s parent is enabled (skip if no parent), then check if element itself is enabled."""
        self.control_keywords.control_is_enabled(control_name)

    @keyword
    def set_control_focus(self, control_name):
        """Set focus to the specified control."""
        self.control_keywords.set_control_focus(control_name)

    @keyword
    def scroll(self, control_name, direction, amount, count=1, retry_interval=None):
        """Scroll a specified control in a given direction.
        direction can be any of “up”, “down”, “left”, “right”
        amount can be only “line” or “page”
        count (optional) the number of times to scroll
        retry_interval (optional) interval between scroll actions"""
        self.control_keywords.scroll(control_name, direction, amount, count, retry_interval)

    @keyword
    def control_has_focus(self, control_name):
        """Check if a control has focus."""
        return self.control_keywords.control_has_focus(control_name)

    @keyword
    def control_is_checkbox(self, control_name):
        """Check if a control is a checkbox."""
        return self.control_keywords.control_is_checkbox(control_name)

    @keyword
    def control_is_button(self, control_name):
        """Check if a control is a button."""
        return self.control_keywords.control_is_button(control_name)

    @keyword
    def control_is_radiobutton(self, control_name):
        """Check if a control is a radio button."""
        return self.control_keywords.control_is_radiobutton(control_name)

    @keyword
    def control_is_groupbox(self, control_name):
        """Check if a control is a group box."""
        return self.control_keywords.control_is_groupbox(control_name)

    @keyword
    def control_is_edit(self, control_name):
        """Check if a control is an edit box."""
        return self.control_keywords.control_is_edit(control_name)

    @keyword
    def control_is_checked(self, control_name):
        """Check if a control is checked."""
        return self.control_keywords.control_is_checked(control_name)

    @keyword
    def control_is_unchecked(self, control_name):
        """Check if a control is unchecked."""
        return self.control_keywords.control_is_unchecked(control_name)

    @keyword
    def control_is_indeterminate(self, control_name):
        """Check if a control is indeterminate."""
        return self.control_keywords.control_is_indeterminate(control_name)

    @keyword
    def set_checkbox_to_checked(self, control_name):
        """Set a checkbox to checked."""
        self.control_keywords.set_checkbox_to_checked(control_name)

    @keyword
    def set_checkbox_to_unchecked(self, control_name):
        """Set a checkbox to unchecked."""
        self.control_keywords.set_checkbox_to_unchecked(control_name)

    @keyword
    def set_checkbox_to_indeterminate(self, control_name):
        """Set a checkbox to indeterminate."""
        self.control_keywords.set_checkbox_to_indeterminate(control_name)

    @keyword
    def get_combobox_items(self, control_name):
        """Retrieve the items of a combobox."""
        return self.control_keywords.get_combobox_items(control_name)

    @keyword
    def get_combobox_item_count(self, control_name):
        """Retrieve the item count of a combobox."""
        return self.control_keywords.get_combobox_item_count(control_name)

    @keyword
    def get_combobox_selected_index(self, control_name):
        """Retrieve the selected index of a combobox starting from 0."""
        return self.control_keywords.get_combobox_selected_index(control_name)

    @keyword
    def get_combobox_selected_value(self, control_name):
        """Retrieve the selected value of a combobox."""
        return self.control_keywords.get_combobox_selected_value(control_name)

    @keyword
    def combobox_select_index(self, control_name, value):
        """Select a combobox item by index."""
        self.control_keywords.combobox_select_index(control_name, value)

    @keyword
    def combobox_select_value(self, control_name, value):
        """Select a combobox item by value."""
        self.control_keywords.combobox_select_value(control_name, value)

    @keyword
    def get_editbox_line_count(self, control_name):
        """Retrieve the line count of an edit box."""
        return self.control_keywords.get_editbox_line_count(control_name)

    @keyword
    def get_editbox_line_text(self, control_name, line_index):
        """Retrieve the text of a specified line in an edit box."""
        return self.control_keywords.get_editbox_line_text(control_name, line_index)

    @keyword
    def get_editbox_text(self, control_name):
        """Retrieve the text of an edit box."""
        return self.control_keywords.get_editbox_text(control_name)

    @keyword
    def set_editbox_text(self, control_name, textblock):
        """Set the text of an edit box."""
        self.control_keywords.set_editbox_text(control_name, textblock)

    @keyword
    def get_listbox_items(self, control_name):
        """Retrieve the items of a listbox."""
        return self.control_keywords.get_listbox_items(control_name)

    @keyword
    def get_listbox_item_count(self, control_name):
        """Retrieve the item count of a listbox."""
        return self.control_keywords.get_listbox_item_count(control_name)

    @keyword
    def get_listbox_selected_index(self, control_name):
        """Retrieve the selected index of a listbox starting from 0."""
        return self.control_keywords.get_listbox_selected_index(control_name)

    @keyword
    def get_listbox_selected_value(self, control_name):
        """Retrieve the selected value of a listbox."""
        return self.control_keywords.get_listbox_selected_value(control_name)

    @keyword
    def listbox_select_index(self, control_name, value):
        """Select a listbox item by index."""
        self.control_keywords.listbox_select_index(control_name, value)

    @keyword
    def listbox_select_value(self, control_name, value):
        """Select a listbox item by value."""
        self.control_keywords.listbox_select_value(control_name, value)

    @keyword
    def listbox_deselect_all(self, control_name):
        """Deselect all currently selected items in a listbox.
        Further adjustments may be needed to address this scenario effectively."""
        self.control_keywords.listbox_deselect_all(control_name)

    @keyword
    def get_listview_column_count(self, control_name):
        """Retrieve the column count of a listview."""
        return self.control_keywords.get_listview_column_count(control_name)

    @keyword
    def get_listview_item_count(self, control_name):
        """Retrieve the item count of a listview."""
        return self.control_keywords.get_listview_item_count(control_name)

    @keyword
    def listview_header_text(self, control_name):
        """Retrieve the header text of a listview."""
        return self.control_keywords.listview_header_text(control_name)

    @keyword
    def listview_get_selected_count(self, control_name):
        """Retrieve the selected item count of a listview."""
        return self.control_keywords.listview_get_selected_count(control_name)

    @keyword
    def listview_index_is_selected(self, control_name, index):
        """Check if a specified index in a list view is selected."""
        return self.control_keywords.listview_index_is_selected(control_name, index)

    @keyword
    def listview_index_is_not_selected(self, control_name, index):
        """Check if a specified index in a list view is not selected."""
        return self.control_keywords.listview_index_is_not_selected(control_name, index)

    @keyword
    def listview_select_index(self, control_name, index):
        """Select a specified index in a list view."""
        self.control_keywords.listview_select_index(control_name, index)

    @keyword
    def listview_deselect_index(self, control_name, index):
        """Deselect a specified index in a list view."""
        self.control_keywords.listview_deselect_index(control_name, index)

    @keyword
    def listview_index_is_checked(self, control_name, index):
        """Check if a specified index in a list view is checked."""
        return self.control_keywords.listview_index_is_checked(control_name, index)

    @keyword
    def listview_index_is_not_checked(self, control_name, index):
        """Check if a specified index in a list view is not checked."""
        return self.control_keywords.listview_index_is_not_checked(control_name, index)

    @keyword
    def listview_check_index(self, control_name, index):
        """Check a specified index in a list view."""
        self.control_keywords.listview_check_index(control_name, index)

    @keyword
    def listview_uncheck_index(self, control_name, index):
        """Uncheck a specified index in a list view."""
        self.control_keywords.listview_uncheck_index(control_name, index)

    @keyword
    def get_statusbar_part_count(self, control_name):
        """Retrieve the number of parts in a status bar."""
        return self.control_keywords.get_statusbar_part_count(control_name)

    @keyword
    def get_statusbar_part_text(self, control_name, index):
        """Retrieve the text of a specified part in a status bar."""
        return self.control_keywords.get_statusbar_part_text(control_name, index)

    @keyword
    def get_statusbar_text(self, control_name):
        """Retrieve the text of a status bar."""
        return self.control_keywords.get_statusbar_text(control_name)

    @keyword
    def get_tab_count(self, control_name):
        """Retrieve the number of tabs."""
        return self.control_keywords.get_tab_count(control_name)

    @keyword
    def get_selected_tab_index(self, control_name):
        """Retrieve the index of the selected tab."""
        return self.control_keywords.get_selected_tab_index(control_name)

    @keyword
    def get_tab_text(self, control_name, index):
        """Retrieve the text of a specified tab."""
        return self.control_keywords.get_tab_text(control_name, index)

    @keyword
    def get_all_tab_texts(self, control_name):
        """Retrieve the texts of all tabs."""
        return self.control_keywords.get_all_tab_texts(control_name)

    @keyword
    def select_tab_by_text(self, control_name, text):
        """Select a tab by its text."""
        self.control_keywords.select_tab_by_text(control_name, text)

    @keyword
    def select_tab_by_index(self, control_name, index):
        """Select a tab by its index."""
        self.control_keywords.select_tab_by_index(control_name, index)

    @keyword
    def get_toolbar_button_count(self, control_name):
        """Retrieve the number of buttons in a toolbar."""
        return self.control_keywords.get_toolbar_button_count(control_name)

    @keyword
    def get_toolbar_button_text(self, control_name, index):
        """Retrieve the text of a specified toolbar button."""
        return self.control_keywords.get_toolbar_button_text(control_name, index)

    @keyword
    def click_toolbar_button(self, control_name, text):
        """Click a toolbar button by its text."""
        self.control_keywords.click_toolbar_button(control_name, text)

    @keyword
    def get_tree_text(self, control_name):
        """Retrieve the text of a tree control."""
        return self.control_keywords.get_tree_text(control_name)

    @keyword
    def click_tree_element(self, control_name, path):
        """
        Click a treeview control element by the path within the tree.
        The path should take the same form as a menu, where -> delimits the nodes.
        For example "Tree->Node2->SubNode1" would specify SubNode1 on the tree below

        + Tree
        ----+ Node1
        ----+ Node2
        --------+ SubNode1
        ----+ Node3
        --------+ SubNode2
        """
        self.control_keywords.click_tree_element(control_name, path)

    @keyword
    def right_click_tree_element(self, control_name, path):
        """
        Right click a treeview control element by the path within the tree.
        The path should take the same form as a menu, where -> delimits the nodes.
        For example "Tree->Node2->Subnode1" would specify Subnode1 on the tree below

        + Tree
        ----+ Node1
        ----+ Node2
        --------+ Subnode1
        ----+ Node3
        --------+ Subnode2
        """
        self.control_keywords.right_click_tree_element(control_name, path)

    @keyword
    def double_click_tree_element(self, control_name, path):
        """
        Double Click a treeview control element by the path within the tree.
        The path should take the same form as a menu, where -> delimits the nodes.
        For example "Tree->Node2->Subnode1" would specify Subnode1 on the tree below

        + Tree
        ----+ Node1
        ----+ Node2
        --------+ Subnode1
        ----+ Node3
        --------+ Subnode2
        """
        self.control_keywords.double_click_tree_element(control_name, path)

    @keyword
    def expand_tree_element(self, control_name, path):
        """
        Expand a treeview control element by the path within the tree.
        The path should take the same form as a menu, where -> delimits the nodes.
        For example "Tree->Node2->Subnode1" would specify Subnode1 on the tree below

        + Tree
        ----+ Node1
        ----+ Node2
        --------+ Subnode1
        ----+ Node3
        --------+ Subnode2
        """
        self.control_keywords.expand_tree_element(control_name, path)

    @keyword
    def popup_menu_select(self, menulocation):
        """Select a menu item from a popup menu by its location."""
        self.dialog_keywords.popup_menu_select(menulocation)
