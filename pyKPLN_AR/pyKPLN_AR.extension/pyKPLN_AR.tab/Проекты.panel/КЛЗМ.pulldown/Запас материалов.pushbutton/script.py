# -*- coding: utf-8 -*-
"""
KPLN

"""
__author__ = 'Gyulnara Fedoseeva - gyulnaraf@gmail.com'
__title__ = "Запас материалов"
__doc__ = 'Добавляет указанный процент запаса к площади и объему элементов категорий стен, перекрытий, потолков' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr
import wpf
from System import Guid
from System.Windows import Application, Window
from System.Collections.ObjectModel import ObservableCollection
import math
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, FamilyInstance,\
                              RevitLinkInstance
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit import revit
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert
import datetime
from Autodesk.Revit.DB import *
from Autodesk.Revit.ApplicationServices import Application
from System.Security.Principal import WindowsIdentity
from rpw.ui.forms import*
from System import Enum, Guid
from rpw.ui.forms import CommandLink, TaskDialog, Alert
import datetime
import webbrowser


doc = revit.doc
output = script.get_output()


param_area = BuiltInParameter.HOST_AREA_COMPUTED
param_volume = BuiltInParameter.HOST_VOLUME_COMPUTED



class MyWindow(Window):
    def __init__(self):
        wpf.LoadComponent(self, 'X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Проекты.panel\\КЛЗМ.pulldown\\Запас материалов.pushbutton\\Form.xaml')
        self.Value_walls = 0
        self.Value_floors = 0
        self.Value_ceilings = 0

    def Update(self):
        try:
            self.Value_walls = int(self.tbStartNumber_walls.Text)
            self.Value_floors = int(self.tbStartNumber_floors.Text)
            self.Value_ceilings = int(self.tbStartNumber_ceilings.Text)
            self.OnOk.IsEnabled = True
        except:
            self.OnOk.IsEnabled = False

    def OnTextChanged(self, sender, e):
        self.Update()

    def OnButtonApply(self, sender, e):
        with db.Transaction(name = "Расчитать с запасом"):
            walls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
            ceilings = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements()
            floors = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements()
            for wall in walls:
                #добавление запаса к площади стен
                walls_area_value = wall.get_Parameter(param_area).AsValueString()
                if ' ' in walls_area_value:
                    walls_area_value = walls_area_value.split(' ')[0]
                if 'м'.lower() in walls_area_value.lower():
                    walls_area_value = walls_area_value.split('м')[0]
                if ',' in walls_area_value:
                    walls_area_value = float(walls_area_value.split(',')[0]+'.'+ walls_area_value.split(',')[1])
                else:
                    walls_area_value = float(walls_area_value)
                if self.Value_walls != 0: 
                    new_walls_area_value = walls_area_value * (1+float(self.Value_walls)/100)
                    try:
                        wall.LookupParameter('КП_Площадь с запасом').Set(new_walls_area_value)
                    except:
                        output.print_md("Стена не имеет площадь: {}".format(output.linkify(wall.Id)))
                #добавление запаса к объему стен
                walls_volume_value = wall.get_Parameter(param_volume).AsValueString()
                if ' ' in walls_volume_value:
                    walls_volume_value = walls_volume_value.split(' ')[0]
                if 'м'.lower() in walls_volume_value.lower():
                    walls_volume_value = walls_volume_value.split('м')[0]
                if ',' in walls_volume_value:
                    walls_volume_value = float(walls_volume_value.split(',')[0]+'.'+ walls_volume_value.split(',')[1])
                else:
                    walls_volume_value = float(walls_volume_value)
                if self.Value_walls != 0: 
                    new_walls_volume_value = walls_volume_value * (1+float(self.Value_walls)/100)
                    try:
                        wall.LookupParameter('КП_Объем с запасом').Set(new_walls_volume_value)
                    except:
                        output.print_md("Стена не имеет объем: {}".format(output.linkify(wall.Id)))
            for ceiling in ceilings:
                #добавление запаса к площади потолков
                ceilings_area_value = ceiling.get_Parameter(param_area).AsValueString()
                if ' ' in ceilings_area_value:
                    ceilings_area_value = ceilings_area_value.split(' ')[0]
                if 'м'.lower() in ceilings_area_value.lower():
                    ceilings_area_value = ceilings_area_value.split('м')[0]
                if ',' in ceilings_area_value:
                    ceilings_area_value = float(ceilings_area_value.split(',')[0]+'.'+ ceilings_area_value.split(',')[1])
                else:
                    ceilings_area_value = float(ceilings_area_value)
                if self.Value_ceilings != 0: 
                    new_ceilings_area_value = ceilings_area_value * (1+float(self.Value_ceilings)/100)
                    try:
                        ceiling.LookupParameter('КП_Площадь с запасом').Set(new_ceilings_area_value)
                    except:
                        output.print_md("Потолок не имеет площадь: {}".format(output.linkify(ceiling.Id)))
            for floor in floors:
                #добавление запаса к площади перекрытий
                floors_area_value = floor.get_Parameter(param_area).AsValueString()
                if ' ' in floors_area_value:
                    floors_area_value = floors_area_value.split(' ')[0]
                if 'м'.lower() in floors_area_value.lower():
                    floors_area_value = floors_area_value.split('м')[0]
                if ',' in floors_area_value:
                    floors_area_value = float(floors_area_value.split(',')[0]+'.'+ floors_area_value.split(',')[1])
                else:
                    floors_area_value = float(floors_area_value)
                if self.Value_floors != 0: 
                    new_floors_area_value = floors_area_value * (1+float(self.Value_floors)/100)
                    try:
                        floor.LookupParameter('КП_Площадь с запасом').Set(new_floors_area_value)
                    except:
                        output.print_md("Перекрытие не имеет площадь: {}".format(output.linkify(floor.Id)))
                #добавление запаса к объему перекрытий
                floors_volume_value = floor.get_Parameter(param_volume).AsValueString()
                if ' ' in floors_volume_value:
                    floors_volume_value = floors_volume_value.split(' ')[0]
                if 'м'.lower() in floors_volume_value.lower():
                    floors_volume_value = floors_volume_value.split('м')[0]
                if ',' in floors_volume_value:
                    floors_volume_value = float(floors_volume_value.split(',')[0]+'.'+ floors_volume_value.split(',')[1])
                else:
                    floors_volume_value = float(floors_volume_value)
                if self.Value_floors != 0: 
                    new_floors_volume_value = floors_volume_value * (1+float(self.Value_floors)/100)
                    try:
                        floor.LookupParameter('КП_Объем с запасом').Set(new_floors_volume_value)
                    except:
                        output.print_md("Перекрытие не имеет объем: {}".format(output.linkify(floor.Id)))
        self.Close()
window = MyWindow()
window.ShowDialog()