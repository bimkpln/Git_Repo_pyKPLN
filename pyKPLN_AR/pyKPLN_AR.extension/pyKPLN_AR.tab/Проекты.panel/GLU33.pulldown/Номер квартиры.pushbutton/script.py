# -*- coding: utf-8 -*-
"""
KPLN:GLU33:ROOM:COUNTER

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Номер квартиры"
__doc__ = 'Увиличивает или уменьшает номера всех квартир (кроме номера «0») на заданное значение. Необходимо для сквозной нумерации между отдельными проектами rvt' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr

import wpf
from System.Windows import Application, Window
from System.Collections.ObjectModel import ObservableCollection

import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert
import datetime
import webbrowser

class MyWindow(Window):
    def __init__(self):
        wpf.LoadComponent(self, 'Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Проекты.panel\\GLU33.pulldown\\Номер квартиры.pushbutton\\Form.xaml')
        self.Parameters = []
        self.Value = 0
        room = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().FirstElement()
        for j in room.Parameters:
            if j.IsShared and j.StorageType == DB.StorageType.String:
                self.Parameters.append(j)
        self.Parameters.sort(key=lambda x: x.Definition.Name)
        self.cbxPameter.ItemsSource = self.Parameters

    def Update(self):
        if self.cbxPameter.SelectedIndex == -1:
            self.OnOk.IsEnabled = False
            return
        try:
            self.Value = int(self.tbStartNumber.Text)
            self.OnOk.IsEnabled = True
        except:
            self.OnOk.IsEnabled = False

    def OnSelectionChanged(self, sender, e):
        self.Update()
    
    def OnTextChanged(self, sender, e):
        self.Update()

    def OnButtonApply(self, sender, e):
        with db.Transaction(name = "Нумерация квартир"):
            try:
                for room in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements():
                    value = ""
                    param = None
                    for j in room.Parameters:
                        if j.IsShared and j.Definition.Name == self.cbxPameter.SelectedItem.Definition.Name and j.StorageType == DB.StorageType.String:
                            value = j.AsString()
                            param = j
                            break
                    if value != "" and value != "0" and value != None and param != None:
                        try:
                            number = int(value)
                            param.Set(str(number + self.Value))
                        except :
                            pass
            except Exception as e:
                print(str(e));

        self.Close()

MyWindow().ShowDialog()