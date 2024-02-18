import unittest

from comtypes.client import *
from comtypes.gen.UIAutomationClient import *

from package.win_ui_automation import WinUIAutomation


class TestWinUIAutomation(unittest.TestCase):
    def setUp(self):
        self.win_ui_automation = WinUIAutomation()

    def test_get_window_element_with_valid_title(self):
        window_element = self.win_ui_automation.get_window_element("電卓")
        self.assertIsNotNone(window_element)

    def test_get_window_element_with_invalid_title(self):
        with self.assertRaises(ValueError):
            self.win_ui_automation.get_window_element(123)

    def test_find_control(self):
        # Arrange
        window_element = self.win_ui_automation.get_window_element("電卓")
        control_type = UIA_ButtonControlTypeId

        # Act
        control_elements = self.win_ui_automation.find_control(
            window_element, control_type
        )

        # Assert
        self.assertEqual(len(control_elements), 36)

    def test_lookup_by_name(self):
        # Arrange
        window_element = self.win_ui_automation.get_window_element("電卓")
        control_type = UIA_ButtonControlTypeId
        button_name = "OK"

        # Act
        button_element = self.win_ui_automation.lookup_by_name(
            self.win_ui_automation.find_control(window_element, control_type),
            button_name,
        )

        # Assert
        self.assertIsNotNone(button_element)
        self.assertEqual(button_element.CurrentName, button_name)

    def test_lookup_by_automationid(self):
        # Arrange
        window_element = self.win_ui_automation.get_window_element("Notepad")
        control_type = UIA_ButtonControlTypeId

        # Act
        button_element = self.win_ui_automation.lookup_by_automationid(
            self.win_ui_automation.find_control(window_element, control_type),
            "Button",
        )

        # Assert
        self.assertIsNotNone(button_element)

    def test_click_button(self):
        window_element = self.win_ui_automation.get_window_element("電卓")
        control_type = UIA_ButtonControlTypeId

        button_element = self.win_ui_automation.find_control(
            window_element, control_type
        )
        self.win_ui_automation.click_button(button_element)


if __name__ == "__main__":
    unittest.main()
