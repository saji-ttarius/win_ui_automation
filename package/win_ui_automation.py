# -*- coding: utf-8 -*-
"""
win_ui_automation : Windows UI Automation utility
"""
import comtypes
import comtypes.client
from comtypes.client import *

comtypes.client.GetModule("UIAutomationCore.dll")

from comtypes import CoCreateInstance
from comtypes.gen.UIAutomationClient import *


class WinUIAutomation():
    _instance = None
    __uia = None
    __root_element = None

    def __new__(cls):
        """
        The class constructor.

        Args:
            cls (type): The class object.

        Returns:
            An instance of the WinUIAutomation class.
        """
        if cls._instance is None:
            global __uia, __root_element
            cls._instance = super(WinUIAutomation, cls).__new__(cls)
            __uia = CoCreateInstance(
                CUIAutomation._reg_clsid_,
                interface=IUIAutomation,
                clsctx=comtypes.CLSCTX_INPROC_SERVER,
            )
            __root_element = __uia.GetRootElement()
        return cls._instance

    def get_window_element(self, title):
        """
        Get the AutomationElement for a window with the specified title.

        Args:
            title (str): The title of the window to find.

        Returns:
            IUIAutomationElement: The AutomationElement for the window,
            or None if no window was found with the specified title.

        Raises:
            ValueError: If the window title is not a string.
        """
        if not isinstance(title, str):
            raise ValueError("Window title must be a string")

        # Find the window element using the UI Automation library
        win_element = __root_element.FindFirst(
            TreeScope_Children, __uia.CreatePropertyCondition(UIA_NamePropertyId, title)
        )
        return win_element

    def find_control(self, base_element, ctrl_type):
        """
        Find all controls of a specific type under a base element.

        Args:
            base_element (IUIAutomationElement): base element
            ctrl_type (int): control type ID

        Returns:
            List[IUIAutomationElement]: list of control elements
        """
        condition = __uia.CreatePropertyCondition(UIA_ControlTypePropertyId, ctrl_type)
        ctl_elements = base_element.FindAll(TreeScope_Subtree, condition)
        return [ctl_elements.GetElement(i) for i in range(ctl_elements.Length)]

    def lookup_by_name(self, elements, name):
        """
        Find the first element in the given iterable with a name matching
        the given name.

        Args:
            elements: An iterable of IUIAutomationElement objects.
            name: The name to match against.

        Returns:
            The first element in the iterable with a name matching
            the given name, or None if no match was found.
        """
        for element in elements:
            if element.CurrentName == name:
                return element

        return None

    def lookup_by_automationid(self, elements, id):
        """
        Find the first element in the given iterable with an AutomationId
        matching the given id.

        Args:
            elements (Iterable[IUIAutomationElement]): An iterable of
            AutomationElement objects.
            id (str): The AutomationId to match against.

        Returns:
            IUIAutomationElement: The first element in the iterable with an
            AutomationId matching the given id, or None if no match was found.
        """
        for element in elements:
            if element.CurrentAutomationId == id:
                return element
        return None

    def click_button(self, element):
        """
        Clicks a button using UI Automation.

        Args:
            element (IUIAutomationElement): The button element to click.

        Raises:
            ElementNotInteractableException: If the element is not clickable.
        """
        # Check if the element is clickable
        is_clickable = element.GetCurrentPropertyValue(
            UIA_IsInvokePatternAvailablePropertyId
        )
        if is_clickable is True:
            # Get the InvokePattern for the element
            invoke_pattern = element.GetCurrentPattern(UIA_InvokePatternId)
            # Invoke the click action
            invoke_pattern.QueryInterface(IUIAutomationInvokePattern).Invoke()
        else:
            raise ElementNotInteractableException("The element is not clickable.")
