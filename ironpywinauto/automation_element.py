
import inspect

import clr
clr.AddReference('UIAutomationClient')
clr.AddReference('UIAutomationTypes')
clr.AddReference('System.Windows.Forms')
from System.Windows.Automation import AutomationElement, PropertyCondition, TreeScope, Condition, Automation, InvokePattern, TextPattern, TreeWalker

'''
from System.Windows.Automation import AutomationPattern, BasePattern, DockPattern, ExpandCollapsePattern, GridItemPattern, GridPattern
from System.Windows.Automation import ItemContainerPattern, MultipleViewPattern, RangeValuePattern, ScrollItemPattern, ScrollPattern
from System.Windows.Automation import SelectionItemPattern, SelectionPattern, SynchronizedInputPattern, TableItemPattern, TablePattern
from System.Windows.Automation import TextPattern, TogglePattern, TransformPattern, ValuePattern, VirtualizedItemPattern, WindowPattern
'''
from System.Windows.Forms import SendKeys

class PythonicAutomationElement(object):

    def __init__(self, auto_elem):
        if not isinstance(auto_elem, AutomationElement):
            raise TypeError('PythonicAutomationElement can be initialized with AutomationElement instance only!')
        self.elem = auto_elem

    AutomationId = property(lambda self: str(self.elem.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty).strip("'")),
            doc="AutomationId property")

    Name = property(lambda self: repr(self.elem.GetCurrentPropertyValue(AutomationElement.NameProperty)).encode('utf-8').strip("'"),
            doc="Name property")

    ClassName = property(lambda self: str(self.elem.GetCurrentPropertyValue(AutomationElement.ClassNameProperty).strip("'")),
            doc="ClassName property")

    ControlType = property(lambda self: str(self.elem.GetCurrentPropertyValue(AutomationElement.ControlTypeProperty).ProgrammaticName).lstrip('ControlType.').strip("'"),
            doc="ControlType property")

    BoundingRectangle = property(lambda self: self.elem.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty),
            doc="Rectangle property")
    Rectangle = BoundingRectangle

    IsEnabled = property(lambda self: self.elem.GetCurrentPropertyValue(AutomationElement.IsEnabledProperty),
            doc="IsEnabled property")

    def __getattribute__(self, attr_name):
        default_attrs = [attr for attr in dir(PythonicAutomationElement) if attr != '__getattribute__']
        default_attrs.extend(dir(self))
        if attr_name in default_attrs:
            return object.__getattribute__(self, attr_name)
        for prop in self.elem.GetSupportedProperties():
            prop_name = str(Automation.PropertyName(prop))
            if prop_name == attr_name:
                return self.elem.GetCurrentPropertyValue(prop)
        # TODO: get child using best match algorithm
        # FindAll(TreeScope.Descendants, Condition.TrueCondition)        
        raise AttributeError()

    def FindAll(self, scope, condition):
        return [PythonicAutomationElement(elem) for elem in self.elem.FindAll(scope, condition)]

    def FindFirst(self, scope, condition):
        return PythonicAutomationElement(self.elem.FindFirst(scope, condition))

    def FindRelative(self, path):
        '''
        Find element using AutonationId or Name with relative navigation 
        path like lblNewItem.Edit - separated by point        
        '''
        #replace this with RK code
        path_parts = path.split(".")
        elements = self.FindAll(TreeScope.Descendants, Condition.TrueCondition)
        el = [l for l in elements if l.AutomationId == path_parts[0]][0]
        #endreplaceelement

        tw = TreeWalker(Condition.TrueCondition)
        siblingElement = tw.GetNextSibling(el.elem)
        currentElement = PythonicAutomationElement(el.elem)
        while siblingElement is not None:
            currentElement = PythonicAutomationElement(siblingElement)
            if currentElement.ControlType.lower() == path_parts[1].lower():
                break
            siblingElement = tw.GetNextSibling(currentElement.elem)

        return PythonicAutomationElement(siblingElement)

    def GetSupportedProperties(self):
        properties = {}
        for prop in self.elem.GetSupportedProperties():
            name = str(Automation.PropertyName(prop))
            if not (name in dir(PythonicAutomationElement)):
                properties[name] = self.elem.GetCurrentPropertyValue(prop)
        return properties

    def GetImportantProperties(self):
        properties = self.GetSupportedProperties()
        del properties['HelpText']
        del properties['IsKeyboardFocusable']
        del properties['IsPassword']
        del properties['Orientation']
        del properties['IsRequiredForForm']
        del properties['IsOffscreen']
        del properties['RuntimeId']
        del properties['LabeledBy']
        del properties['IsContentElement']
        del properties['LocalizedControlType']
        del properties['ItemStatus']
        del properties['ProcessId']
        del properties['HasKeyboardFocus']
        del properties['FrameworkId']
        del properties['IsControlElement']
        del properties['ItemType']
        del properties['AcceleratorKey']
        del properties['AccessKey']
        return properties

    def __print_immediate_controls(self, indent=0):
        children = self.FindAll(TreeScope.Children, Condition.TrueCondition)

        indent_str = ""
        for i in xrange(indent):
            indent_str += "        "

        for ctrl in children:
            print("%s%s - '%s'   %s"% (indent_str, ctrl.ControlType, ctrl.Name, str(ctrl.Rectangle))) # ctrl.WindowText()
            print(indent_str + "\tProperties: " + str(ctrl.GetImportantProperties())) #.keys()
            print(indent_str + "\tAutomationId: '" + str(ctrl.AutomationId) + "'\n")
            ctrl.__print_immediate_controls(indent + 1)

    def PrintControlIdentifiers(self):
        self.__print_immediate_controls(0)
        '''
        allSubElements = self.FindAll(TreeScope.Descendants, Condition.TrueCondition)

        for ctrl in allSubElements:
            print("%s - '%s'   %s"% (ctrl.ControlType, ctrl.Name, str(ctrl.Rectangle))) # ctrl.WindowText()

            print("\tProperties: " + str(ctrl.GetImportantProperties())) #.keys()
            print("\tAutomationId: " + str(ctrl.tAutomationId) + "\n")
        '''

