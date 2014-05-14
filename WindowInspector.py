import sys
import os
import time

import clr
clr.AddReference('UIAutomationClient')
clr.AddReference('UIAutomationTypes')
from System.Windows.Automation import AutomationElement, PropertyCondition, TreeScope, Condition, Automation
from System.Diagnostics import Process
from System.Threading import Thread

from ironpywinauto.automation_element import PythonicAutomationElement

clr.AddReferenceToFile('uiauto.dll')
import iprcs

filename = sys.argv[1]
if filename and os.path.exists(filename) and os.path.isfile(filename):
    print "Starting application..."
    proc = Process.Start(filename)
    time.sleep(3)
else:
    print filename + " not found"
    exit()

print "Finding window with process id=" + str(proc.Id) + "..."
condition = PropertyCondition(AutomationElement.ProcessIdProperty, proc.Id)
mainWindow = PythonicAutomationElement(iprcs.uiauto().RootElement()).FindFirst(TreeScope.Children, condition);
allElements = mainWindow.FindAll(TreeScope.Descendants, Condition.TrueCondition);

inspectFile = os.getcwd() + r"\output.txt"
if not os.path.exists(inspectFile):
    print "file not found, creating new"
f = open(inspectFile, "w")

f.write('Main window attributes:\n\n')
for attr in dir(mainWindow):
    try:
        f.write(attr + ' = "' + repr(mainWindow.GetCurrentPropertyValue(getattr(AutomationElement, attr))) + '"\n')
    except:
        try:
            f.write(attr + ' = "' + repr(getattr(mainWindow, attr)) + '"\n')
        except:
            f.write(attr + '\n')
f.write('\n\n')

print "Gathering controls information..."
print mainWindow.PrintControlIdentifiers()

f.close()
print "Killing application..."
proc.Kill()

print "Done"