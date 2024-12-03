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
from pywinauto.keyboard import send_keys


class ControlKeywords:
    """Keywords for interacting with controls in dialogs."""

    def __init__(self, dlg=None):
        self.dlg = dlg

    def set_dialog(self, dlg):
        """Set the dialog instance for the control keywords."""
        self.dlg = dlg

    def get_control_text(self, control_name):
        """Retrieve the text of a specified control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control_name].window_text()

    def menu_select(self, menulocation):
        """Select a menu item by its location (e.g., 'File -> Save')."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg.menu_select(menulocation)

    def type_text(self, control_name, text):
        """Type text into a specified control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control_name].type_keys(text, with_spaces=True)

    def send_keys(self, keys):
        """Send keyboard input to the current dialog."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        send_keys(keys, with_spaces=True)

    def click(self, control_name):
        """Click on a specified control in the current dialog."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control_name].click()

    def real_click(self, control_name):
        """Real click (simulated as physical) on a specified control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control_name].click_input()

    def right_click(self, control_name):
        """Right-click on a specified control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control_name].right_click()

    def real_right_click(self, control_name):
        """Real right-click (simulated as physical) on a specified control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control_name].right_click_input()

    def double_click(self, control_name):
        """Double-click on a specified control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control_name].double_click()

    def real_double_click(self, control_name):
        """Real double-click (simulated as physical) on a specified control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control_name].double_click_input()

    def drag_mouse(self, control_name, dst, src, button, pressed, absolute):
        """Click on src, drag it and drop on dst"""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        if isinstance(dst, str) and dst.startswith("(") and dst.endswith(")"):
            dst = tuple(map(int, dst.strip("()").split(",")))
        #else:
        #    dst = self.dlg.child_window(title=dst)
        if isinstance(src, str) and src.startswith("(") and src.endswith(")"):
            src = tuple(map(int, src.strip("()").split(",")))
        #else:
        #   src = self.dlg.child_window(title=src)

        self.dlg[control_name].drag_mouse_input(dst=dst, src=src, button=button, pressed=pressed, absolute=absolute)

    def control_is_active(self, control_name):
        """Check if a control is active."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control_name].verify_actionable()

    def control_is_visible(self, control_name):
        """Check if a control is visible."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control_name].verify_visible()

    def control_is_enabled(self, control_name):
        """Check if a control is enabled."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control_name].verify_enabled()

    def set_control_focus(self, control):
        """Set focus to the specified control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].set_focus()

    def scroll(self, control, direction, amount, count, retry_interval):
        """Scroll the specified control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].scroll(direction, amount, count, retry_interval)

    def control_has_focus(self, control):
        """Check if the specified control has focus."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        exp = self.dlg[control]
        act = self.dlg.get_focus()
        assert act == exp, f"Expected {exp}. But got {act}."

    def control_is_checkbox(self, control):
        """Check if the specified control is a checkbox."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        exp = 'CheckBox'
        act = self.dlg[control].friendly_class_name()
        assert act == exp, f"Expected {exp}. But got {act}."

    def control_is_button(self, control):
        """Check if the specified control is a button."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        exp = 'Button'
        act = self.dlg[control].friendly_class_name()
        assert act == exp, f"Expected {exp}. But got {act}."

    def control_is_radiobutton(self, control):
        """Check if the specified control is a radio button."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        exp = 'RadioButton'
        act = self.dlg[control].friendly_class_name()
        assert act == exp, f"Expected {exp}. But got {act}."

    def control_is_groupbox(self, control):
        """Check if the specified control is a group box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        exp = 'GroupBox'
        act = self.dlg[control].friendly_class_name()
        assert act == exp, f"Expected {exp}. But got {act}."

    def control_is_edit(self, control):
        """Check if the specified control is an edit box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        exp = 'Edit'
        act = self.dlg[control].friendly_class_name()
        assert act == exp, f"Expected {exp}. But got {act}."

    def control_is_checked(self, control):
        """Check if the specified control is checked."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        exp = 1
        act = self.dlg[control].get_check_state()
        assert act == exp, f"Expected {exp}. But got {act}."

    def control_is_unchecked(self, control):
        """Check if the specified control is unchecked."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        exp = 0
        act = self.dlg[control].get_check_state()
        assert act == exp, f"Expected {exp}. But got {act}."

    def control_is_indeterminate(self, control):
        """Check if the specified control is indeterminate."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        exp = 2
        act = self.dlg[control].get_check_state()
        assert act == exp, f"Expected {exp}. But got {act}."

    def set_checkbox_to_checked(self, control):
        """Set the specified checkbox to checked."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].check()

    def set_checkbox_to_unchecked(self, control):
        """Set the specified checkbox to unchecked."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].uncheck()

    def set_checkbox_to_indeterminate(self, control):
        """Set the specified checkbox to indeterminate."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].set_check_indeterminate()

    def get_combobox_items(self, control):
        """Get items of the specified combobox."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].item_texts()

    def get_combobox_item_count(self, control):
        """Get item count of the specified combobox."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].item_count()

    def get_combobox_selected_index(self, control):
        """Get selected index of the specified combobox."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].selected_index()

    def get_combobox_selected_value(self, control):
        """Get selected value of the specified combobox."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].texts()[0]

    def combobox_select_index(self, control, value):
        """Select item by index in the specified combobox."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].select(int(value))

    def combobox_select_value(self, control, value):
        """Select item by value in the specified combobox."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].select(value)

    def get_editbox_line_count(self, control):
        """Get line count of the specified edit box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].line_count()

    def get_editbox_line_text(self, control, line_index):
        """Get line text of the specified edit box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].get_line(int(line_index))

    def get_editbox_text(self, control):
        """Get text of the specified edit box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].text_block()

    def set_editbox_text(self, control, textblock):
        """Set text of the specified edit box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].set_text(textblock)

    def get_listbox_items(self, control):
        """Get items of the specified list box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].item_texts()

    def get_listbox_item_count(self, control):
        """Get item count of the specified list box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].item_count()

    def get_listbox_selected_index(self, control):
        """Get selected index of the specified list box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].selected_indices()

    def get_listbox_selected_value(self, control):
        """Get selected value of the specified list box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        texts = self.dlg[control].texts()

        selected = []
        for i in self.dlg[control].selected_indices():
            # Select from the texts, the values of each selected item.
            selected.append(texts[i+1])
        return "|".join(selected)

    def listbox_select_index(self, control, value):
        """Select item by index in the specified list box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].select(int(value))

    def listbox_select_value(self, control, value):
        """Select item by value in the specified list box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].select(value)

    def listbox_deselect_all(self, control):
        """Deselect all items in the specified list box."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        for i in self.dlg[control].selected_indices():
            item_rect = self.dlg[control].item_rect(i)
            left = item_rect.left
            top = item_rect.top
            right = item_rect.right
            bottom = item_rect.bottom
            # Calculate the mid-point coordinates
            mid_x = (left + right) // 2
            mid_y = (top + bottom) // 2
            self.dlg[control].click(coords=(mid_x, mid_y))

    def get_listview_column_count(self, control):
        """Get column count of the specified list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].column_count()

    def get_listview_item_count(self, control):
        """Get item count of the specified list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].item_count()

    def listview_header_text(self, control):
        """Get header text of the specified list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        texts = []
        for i in self.dlg[control].columns():
            texts.append(i["text"])
        return texts

    def listview_get_selected_count(self, control):
        """Get selected item count of the specified list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].get_selected_count()

    def listview_index_is_selected(self, control, index):
        """Check if the specified index is selected in the list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        assert self.dlg[control].is_selected(int(index)), f"Index {index} is not selected."

    def listview_index_is_not_selected(self, control, index):
        """Check if the specified index is not selected in the list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        assert not self.dlg[control].is_selected(int(index)), f"Index {index} is selected."

    def listview_select_index(self, control, index):
        """Select the specified index in the list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].select(int(index))

    def listview_deselect_index(self, control, index):
        """Deselect the specified index in the list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].deselect(int(index))

    def listview_index_is_checked(self, control, index):
        """Check if the specified index is checked in the list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        assert self.dlg[control].is_checked(int(index)), f"Index {index} is not checked."

    def listview_index_is_not_checked(self, control, index):
        """Check if the specified index is not checked in the list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        assert not self.dlg[control].is_checked(int(index)), f"Index {index} is checked."

    def listview_check_index(self, control, index):
        """Check the specified index in the list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].check(int(index))

    def listview_uncheck_index(self, control, index):
        """Uncheck the specified index in the list view."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].uncheck(int(index))

    def get_statusbar_part_count(self, control):
        """Get part count of the specified status bar."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].part_count()

    def get_statusbar_part_text(self, control, index):
        """Get part text of the specified status bar."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].get_part_text(int(index))

    def get_statusbar_text(self, control):
        """Get text of the specified status bar."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].texts()

    def get_tab_count(self, control):
        """Get tab count of the specified tab control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].tab_count()

    def get_selected_tab_index(self, control):
        """Get selected tab index of the specified tab control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].get_selected_tab()

    def get_tab_text(self, control, index):
        """Retrieve the text of a specified tab."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].get_tab_text(int(index))

    def get_all_tab_texts(self, control):
        """Retrieve the texts of all tabs."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].texts()

    def select_tab_by_text(self, control, text):
        """Select a tab by its text."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].select(text)

    def select_tab_by_index(self, control, index):
        """Select a tab by its index."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].select(int(index))

    def get_toolbar_button_count(self, control):
        """Retrieve the number of buttons in a toolbar."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].button_count()

    def get_toolbar_button_text(self, control, index):
        """Retrieve the text of a specified toolbar button."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].get_button(int(index)).text

    def click_toolbar_button(self, control, text):
        """Click a toolbar button by its text."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].press_button(text)

    def get_tree_text(self, control):
        """Retrieve the text of a tree control."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        return self.dlg[control].texts()

    def click_tree_element(self, control, path):
        """Click a tree element by its path."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        path = '\\' + path.replace('->', '\\')
        print(path)
        self.dlg[control].get_item(path).click()

    def right_click_tree_element(self, control, path):
        """Right-click a tree element by its path."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        path = '\\' + path.replace('->', '\\')
        self.dlg[control].get_item(path).click(button='right')

    def double_click_tree_element(self, control, path):
        """Double-click a tree element by its path."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        path = '\\' + path.replace('->', '\\')

        self.dlg[control].get_item(path).click(double=True)

    def expand_tree_element(self, control, path):
        """Expand a tree element by its path."""
        if not self.dlg:
            raise RuntimeError("No dialog is currently active.")
        self.dlg[control].get_item(path).expand()
