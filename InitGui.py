#***************************************************************************
#*    Copyright (C) 2023 
#*    This library is free software
#***************************************************************************
import inspect
import os
import sys
import FreeCAD
import FreeCADGui

class screwCvShowCommand:
    def GetResources(self):
        file_path = inspect.getfile(inspect.currentframe())
        module_path=os.path.dirname(file_path)
        return { 
          'Pixmap': os.path.join(module_path, "icons", "screw.svg"),
          'MenuText': "screwCv",
          'ToolTip': "Show/Hide screwCv"}

    def IsActive(self):
        import screwCvAssy
        screwCvAssy
        return True

    def Activated(self):
        try:
          import screwCvAssy
          screwCvAssy.main.d.show()
        except Exception as e:
          FreeCAD.Console.PrintError(str(e) + "\n")

    def IsActive(self):
        import screwCvAssy
        return not FreeCAD.ActiveDocument is None

class screwCvWB(FreeCADGui.Workbench):
    def __init__(self):
        file_path = inspect.getfile(inspect.currentframe())
        module_path=os.path.dirname(file_path)
        self.__class__.Icon = os.path.join(module_path, "icons", "screw.svg")
        self.__class__.MenuText = "screwCv"
        self.__class__.ToolTip = "screwCv by Pascal"

    def Initialize(self):
        self.commandList = ["screwCv_Show"]
        self.appendToolbar("&screwCv", self.commandList)
        self.appendMenu("&screwCv", self.commandList)

    def Activated(self):
        import screwCvAssy
        screwCvAssy
        return

    def Deactivated(self):
        return

    def ContextMenu(self, recipient):
        return

    def GetClassName(self): 
        return "Gui::PythonWorkbench"
FreeCADGui.addWorkbench(screwCvWB())
FreeCADGui.addCommand("screwCv_Show", screwCvShowCommand())

