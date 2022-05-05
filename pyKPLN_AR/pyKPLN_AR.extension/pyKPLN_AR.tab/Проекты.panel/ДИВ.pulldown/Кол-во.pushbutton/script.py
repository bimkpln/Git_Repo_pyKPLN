# -*- coding: utf-8 -*-
"""
KPLN:DIV:AMMOUNT:COUNTER

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Кол-во на этаж"
__doc__ = 'Подсчет кол-ва элементов на этаж с записью значений в выбранные параметры' \

"""
Архитекурное бюро KPLN

"""
import math
import re
from pyrevit.framework import clr
from System.Collections.Generic import *
from System.Windows import Application, Window
from System.Collections.ObjectModel import ObservableCollection
from rpw import doc, uidoc, DB, UI, db, ui, revit
from rpw.ui.forms import CommandLink, TaskDialog, Alert
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
import datetime
import webbrowser
import wpf

class Param():
    def __init__(self, parameter):
        if parameter == None:
            self.Name = "<Нет>"
            self.Parameter = None
        else:
            self.Name = parameter.Definition.Name
            self.Parameter = parameter
    def GetValue(self, element):
        for j in element.Parameters:
            if j.Id.IntegerValue == self.Parameter.Id.IntegerValue:
                return str(j.AsString())
        try:
            for j in element.Symbol.Parameters:
                if j.Id.IntegerValue == self.Parameter.Id.IntegerValue:
                    return str(j.AsString())
        except :
            pass
        try:
            for j in doc.GetElement(element.GetTypeId()).Parameters:
                if j.Id.IntegerValue == self.Parameter.Id.IntegerValue:
                    return str(j.AsString())
        except :
            pass
        return None
class Collection():
    def __init__(self, element, levels, filters, p, l):
        self.Elements = []
        self.AllElements = []
        self.Levels = []
        self.Type = None
        try:
            for j in doc.GetElement(element.GetTypeId()).Parameters:
                if j.Id.IntegerValue == p.Id.IntegerValue:
                    self.Type = j.AsString()
        except: pass
        try:
            for j in element.Symbol.Parameters:
                if j.Id.IntegerValue == p.Id.IntegerValue:
                    self.Type = j.AsString()
        except: pass
        try:
            for j in element.Parameters:
                if j.Id.IntegerValue == p.Id.IntegerValue:
                    self.Type = j.AsString()
        except: pass
        self.Key = self.Type
        for i in filters:
            self.Key += "_" + i.GetValue(element)
        self.Append(element, l)

    def Append(self, element, level):
        for i in range(0, len(self.Levels)):
            if self.Levels[i].SelectedIndex == level.SelectedIndex:
                self.Elements[i].append(element)
                self.AllElements.append(element)
                return
        self.AllElements.append(element)
        self.Elements.append([element])
        self.Levels.append(level)

class ProjectLevel():
    def __init__(self, level, bool, parent):
        self.Elements = []
        self.Id = level.Id
        self.IsEnabled = bool
        self.Elevation = level.Elevation
        self.LevelName = level.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString() + " : "
        self.Parameters = []
        if parent.ComboBoxCategory.SelectedIndex != -1:
            cat = parent.Categories[parent.ComboBoxCategory.SelectedIndex]
            el = DB.FilteredElementCollector(doc).OfCategoryId(cat.Id).WhereElementIsNotElementType().FirstElement()
            try:
                for j in el.Parameters:
                    if j.StorageType == DB.StorageType.String:
                        if(not self.Exist(j, self.Parameters)):
                            self.Parameters.append(j)
            except:
                pass

            try:
                for j in el.Symbol.Parameters:
                    if j.StorageType == DB.StorageType.String:
                        if(not self.Exist(j, self.Parameters)):
                            self.Parameters.append(j)
            except:
                pass
            type = DB.FilteredElementCollector(doc).OfCategoryId(cat.Id).WhereElementIsElementType().FirstElement()
            try:
                for j in type.Parameters:
                    if j.StorageType == DB.StorageType.String:
                        if(not self.Exist(j, self.Parameters)):
                            self.Parameters.append(j)
            except :
                pass

            self.Parameters.sort(key=lambda x: x.Definition.Name)
        self.SelectedIndex = -1

    def Exist(self, parameter, list):
        for i in list:
            if i.Definition.Name == parameter.Definition.Name and i.IsShared == parameter.IsShared:
                return True
        return False
class MyWindow(Window):
    def __init__(self):
        wpf.LoadComponent(self, 'Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Проекты.panel\\ДИВ.pulldown\\Кол-во.pushbutton\\Form.xaml')
        self.LevelCollection = []
        self.Categories = []
        self.Coll = []
        for c in doc.Settings.Categories:
            if not c.IsTagCategory and c.CategoryType == DB.CategoryType.Model and DB.FilteredElementCollector(doc).OfCategoryId(c.Id).OfClass(DB.FamilyInstance).WhereElementIsNotElementType().ToElements().Count != 0:
                self.Categories.append(c)
        self.Categories.sort(key=lambda x: x.Name)
        self.ComboBoxCategory.ItemsSource = self.Categories
        for i in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements():
            self.LevelCollection.append(ProjectLevel(i, self.ComboBoxCategory.SelectedIndex != -1, self))
        self.LevelCollection.sort(key=lambda x: x.Elevation)
        self.iControll.ItemsSource = self.LevelCollection
        self.UpdateGroupParameters()
    def Update(self):
        self.LevelCollection = []
        for i in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements():
            self.LevelCollection.append(ProjectLevel(i, self.ComboBoxCategory.SelectedIndex != -1, self))
        self.LevelCollection.sort(key=lambda x: x.Elevation)
        self.iControll.ItemsSource = self.LevelCollection
    def OnSelectedParameterChanged(self, sender, e):
        try:
            sender.DataContext.SelectedIndex = sender.SelectedIndex
            if self.ComboBoxCategory.SelectedIndex == -1:
                self.btnApply.IsEnabled = False
            else:
                self.btnApply.IsEnabled = True
        except :
            pass
    def UpdateGroupParameters(self):
        params = []
        sorted_params = []
        p = []
        sorted_params.append(Param(None))
        if self.ComboBoxCategory.SelectedIndex != -1:
            cat = self.Categories[self.ComboBoxCategory.SelectedIndex]
            type = DB.FilteredElementCollector(doc).OfCategoryId(cat.Id).OfClass(DB.FamilyInstance).WhereElementIsElementType().FirstElement()
            try:
                for j in type.Parameters:
                    if j.StorageType == DB.StorageType.String:
                        if(not self.Exist(j, p)):
                            params.append(Param(j))
                            p.append(j)
            except :
                pass
            el = DB.FilteredElementCollector(doc).OfCategoryId(cat.Id).OfClass(DB.FamilyInstance).WhereElementIsNotElementType().FirstElement()
            try:
                for j in el.Parameters:
                    if j.StorageType == DB.StorageType.String:
                        if(not self.Exist(j, p)):
                            params.append(Param(j))
                            p.append(j)
            except :
                pass
            try:
                for j in el.Symbol.Parameters:
                    if j.StorageType == DB.StorageType.String:
                        if(not self.Exist(j, p)):
                            params.append(Param(j))
                            p.append(j)
            except :
                pass
            self.gp1.IsEnabled = True
            self.gp2.IsEnabled = True
            self.gp3.IsEnabled = True
            self.gp0.IsEnabled = True
            params.sort(key=lambda x: x.Name, reverse=False)
            for p in params:
                sorted_params.append(p)
            self.gp1.ItemsSource = sorted_params
            self.gp2.ItemsSource = sorted_params
            self.gp3.ItemsSource = sorted_params
            self.gp1.SelectedIndex = 0
            self.gp2.SelectedIndex = 0
            self.gp3.SelectedIndex = 0
        else:
            self.gp1.ItemsSource = None
            self.gp2.ItemsSource = None
            self.gp3.ItemsSource = None
            self.gp1.IsEnabled = False
            self.gp2.IsEnabled = False
            self.gp3.IsEnabled = False
            self.gp0.IsEnabled = False
    def Exist(self, parameter, list):
        for i in list:
            if i.Definition.Name == parameter.Definition.Name and i.IsShared == parameter.IsShared:
                return True
        return False
    def OnSelectedCategoryChanged(self, sender, e):
        self.gp0.ItemsSource = None
        self.Update()
        self.UpdateGroupParameters()
        if self.ComboBoxCategory.SelectedIndex == -1:
            self.btnApply.IsEnabled = False
        else:
            self.btnApply.IsEnabled = True
        parameters = []
        type = DB.FilteredElementCollector(doc).OfCategoryId(self.Categories[self.ComboBoxCategory.SelectedIndex].Id).WhereElementIsElementType().FirstElement()
        try:
            for j in type.Parameters:
                if j.StorageType == DB.StorageType.String:
                    if(not self.Exist(j, parameters)):
                        parameters.append(j)
        except:
            pass
        element = DB.FilteredElementCollector(doc).OfCategoryId(self.Categories[self.ComboBoxCategory.SelectedIndex].Id).WhereElementIsNotElementType().FirstElement()
        try:
            for j in element.Parameters:
                if j.StorageType == DB.StorageType.String:
                    if(not self.Exist(j, parameters)):
                        parameters.append(j)
        except:
            pass
        try:
            for j in element.Symbol.Parameters:
                if j.StorageType == DB.StorageType.String:
                    if(not self.Exist(j, parameters)):
                        parameters.append(j)
        except:
            pass
        parameters.sort(key=lambda x: x.Definition.Name, reverse=False)
        num = 0
        pick = 0
        for j in parameters:
            if not j.IsShared and j.Definition.Name == "Маркировка типоразмера":
                pick = num
            else:
                num += 1
        self.gp0.ItemsSource = parameters
        self.gp0.SelectedIndex = pick
    def GetLevel(self, e):
        level = None
        try:
            level = doc.GetElement(e.LevelId)
            if level == None:
                try:
                    level = doc.GetElement(e.Host.LevelId)
                except :
                    level = doc.GetElement(e.Room.LevelId)
                
        except:
            try:
                level = doc.GetElement(e.Host.LevelId)
            except:
                pass
        return level
    def OnButtonApply(self, sender, e):
        cat = self.Categories[self.ComboBoxCategory.SelectedIndex]
        filters = []
        if self.gp1.SelectedIndex > 0:
            filters.append(self.gp1.SelectedItem)
        if self.gp2.SelectedIndex > 0:
            filters.append(self.gp2.SelectedItem)
        if self.gp3.SelectedIndex > 0:
            filters.append(self.gp3.SelectedItem)
        with db.Transaction(name = "write_a"):
            for e in DB.FilteredElementCollector(doc).OfCategoryId(cat.Id).OfClass(DB.FamilyInstance).WhereElementIsNotElementType().ToElements():
                for level in self.LevelCollection:
                    if level.SelectedIndex != -1:
                        try:
                            e.LookupParameter(level.Parameters[level.SelectedIndex].Definition.Name).Set("")
                        except :
                            pass
                try:
                    level = self.GetLevel(e)
                    if level == None:
                        continue
                    collectionLevel = self.GetLevelFromCollection(level.Id)
                    if self.GetTypeCollection(e, filters) != None:
                        self.GetTypeCollection(e, filters).Append(e, collectionLevel)
                    else:
                        self.Coll.append(Collection(e, self.LevelCollection, filters, self.gp0.SelectedItem, collectionLevel))
                except Exception as e:
                    print(str(e))
                    pass
        with db.Transaction(name = "write_b"):
            try:
                for c in self.Coll:
                    for z in range(0, len(c.Levels)):
                        if c.Levels[z].SelectedIndex != -1 and 1 == 0:
                            print("\tLEVEL Name=" + c.Levels[z].Parameters[c.Levels[z].SelectedIndex].Definition.Name)
                            for e in c.Elements[z]:
                                print("\t\tELEMENT Id=" + e.Id.ToString())
                        value = ""
                        if len(c.Elements[z]) > 0:
                            value = str(len(c.Elements[z]))
                        if c.Levels[z].SelectedIndex != -1:
                            for e in c.AllElements:
                                try:
                                    e.LookupParameter(c.Levels[z].Parameters[c.Levels[z].SelectedIndex].Definition.Name).Set(value)
                                except: pass
            except Exception as e:
                print(str(e))

        self.Close()

    def OnButtonClose(self, sender, e):
    	self.Close()   
    def OnButtonHelp(self, sender, e):
    	webbrowser.open('https://kpln.kdb24.ru/article/86270/')
    def GetLevelFromCollection(self, levelId):
        for l in self.LevelCollection:
            if levelId.IntegerValue == l.Id.IntegerValue:
                return l
        return None
    def GetTypeCollection(self, e, filters):
        type = None
        param = self.gp0.SelectedItem
        try:
            for j in doc.GetElement(e.GetTypeId()).Parameters:
                if j.Id.IntegerValue == param.Id.IntegerValue:
                    type = j.AsString()
        except :
            pass
        try:
            for j in e.Symbol.Parameters:
                if j.Id.IntegerValue == param.Id.IntegerValue:
                    type = j.AsString()
        except :
            pass
        try:
            for j in e.Parameters:
                if j.Id.IntegerValue == param.Id.IntegerValue:
                    type = j.AsString()
        except :
            pass
        key = type
        for i in filters:
            key += "_" + i.GetValue(e)
        for i in self.Coll:
            if i.Key == key:
                return i
        return None
MyWindow().ShowDialog()