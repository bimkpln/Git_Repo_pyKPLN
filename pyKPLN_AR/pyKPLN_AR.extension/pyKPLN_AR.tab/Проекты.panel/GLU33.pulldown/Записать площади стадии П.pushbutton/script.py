# -*- coding: utf-8 -*-
"""
GLU33

"""
__author__ = 'Gyulnara Fedoseeva - gyulnaraf@gmail.com'
__title__ = "Записать площади стадии П"
__doc__ = '' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr
import wpf
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



out = script.get_output()
out.set_title("Импортировать значения (import)")

def in_list(element, list):
    for i in list:
        if element == i:
            return True
    return False


class MyWindow(Window):
    def __init__(self):
        wpf.LoadComponent(self, 'Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Проекты.panel\\GLU33.pulldown\\Записать площади стадии П.pushbutton\\Form.xaml')
        self.Value = 0

    def Update(self):
        try:
            self.Value = str(self.tbStartNumber_K.Text)
            self.OnOk.IsEnabled = True
        except:
            self.OnOk.IsEnabled = False

    def OnTextChanged(self, sender, e):
        self.Update()

    def OnButtonApply(self, sender, e):
        if self.Value == "К0":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К0\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К0\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К0\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К0\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К0\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К0\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К0\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К0\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К0\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К1":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К1\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К1\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К1\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К1\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К1\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К1\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К1\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К1\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К1\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К2":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К2\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К2\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К2\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К2\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К2\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К2\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К2\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К2\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К2\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К3":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К3\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К3\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К3\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К3\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К3\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К3\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К3\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К3\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К3\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К4":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К4\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К4\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К4\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К4\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К4\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К4\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К4\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К4\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К4\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К5":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К5\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К5\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К5\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К5\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К5\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К5\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К5\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К5\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К5\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К6":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К6\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К6\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К6\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К6\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К6\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К6\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К6\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К6\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К6\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К7":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К7\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К7\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К7\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К7\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К7\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К7\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К7\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К7\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К7\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К8":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К8\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К8\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К8\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К8\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К8\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К8\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К8\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К8\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К8\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К9":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К9\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К9\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К9\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К9\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К9\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К9\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К9\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К9\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К9\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К10":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К10\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К10\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К10\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К10\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К10\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К10\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К10\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К10\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К10\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К11":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К11\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К11\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К11\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К11\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К11\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К11\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К11\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К11\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К11\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "К12":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К12\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К12\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К12\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К12\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К12\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К12\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К12\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К12\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\К12\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        elif self.Value == "ДОУ":
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ДОУ\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ДОУ\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ДОУ\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ДОУ\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ДОУ\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ДОУ\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ДОУ\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ДОУ\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ДОУ\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        else:
            with db.Transaction(name = "Запись результатов"):
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ПАР\\param_0.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Жилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ПАР\\param_1.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Летние').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ПАР\\param_2.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Нежилая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ПАР\\param_3.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ПАР\\param_4.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Общая_К').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ПАР\\param_5.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_КВ_Площадь_Отапливаемые').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ПАР\\param_6.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ПАР\\param_7.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь').Set(float(readLineSplitedSplited[0]))
                with open('Y:\\Жилые Здания\\СПБ Глухарская 33 участок\\12. BIM\\08_3.АР\\1. RVT\\Площади Стадии П\\ПАР\\param_8.txt', 'r') as docum:
                        readLine = docum.read()
                readLineSplited = readLine.split(']')
                for i in readLineSplited:
                    readLineSplitedSplited = i[1:].split('/')
                    if len(readLineSplitedSplited) < 2:
                        break
                    S = float(readLineSplitedSplited[0])
                    Ids = int(readLineSplitedSplited[1])
                    room = doc.GetElement(ElementId(int(readLineSplitedSplited[1])))
                    room.LookupParameter('G33Ex_ПОМ_Площадь_К').Set(float(readLineSplitedSplited[0]))
        self.Close()
window = MyWindow()
window.ShowDialog()