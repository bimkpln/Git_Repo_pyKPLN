# -*- coding: utf-8 -*-
"""
Маркировка проемов

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Маркировать проемы"
__doc__ = 'Присвоение марок для стеновых проемов «199_Отверстие в стене прямоугольное»' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from rpw.ui.forms import CommandLink, TaskDialog, Alert
from System.Collections.Generic import *
from rpw.ui.forms import TextInput
import System
from System import Enum, Guid
from System.Windows.Forms import *
from System.Drawing import *
from itertools import chain
from collections import OrderedDict

def AddParameter(par_name):
    try:
        param_found = False
        app = doc.Application
        category_set_elements = app.Create.NewCategorySet()
        insert_cat_elements = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_MechanicalEquipment)
        category_set_elements.Insert(insert_cat_elements)
        originalFile = app.SharedParametersFilename
        app.SharedParametersFilename = "Z:\\Отдел BIM\\02_Внутренняя документация\\05_ФОП\\ФОП2019_КП_АР.txt"
        SharedParametersFile = app.OpenSharedParameterFile()
        map = doc.ParameterBindings
        it = map.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            d_Definition = it.Key
            d_Name = it.Key.Name
            d_Binding = it.Current
            d_catSet = d_Binding.Categories
            if d_Name == par_name:
                if d_Binding.GetType() == DB.InstanceBinding:
                    if str(d_Definition.ParameterType) == "Text":
                        if d_Definition.VariesAcrossGroups:
                            if d_catSet.Contains(DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_MechanicalEquipment)):
                                param_found = True
        if not param_found:
            with db.Transaction(name = "AddSharedParameter"):
                for dg in SharedParametersFile.Groups:
                    if dg.Name == "АРХИТЕКТУРА - Дополнительные":
                        externalDefinition = dg.Definitions.get_Item(par_name)
                        newIB = app.Create.NewInstanceBinding(category_set_elements)
                        doc.ParameterBindings.Insert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
                        doc.ParameterBindings.ReInsert(externalDefinition, newIB, DB.BuiltInParameterGroup.PG_DATA)
            map = doc.ParameterBindings
            it = map.ForwardIterator()
            it.Reset()
            with db.Transaction(name = "SetAllowVaryBetweenGroups"):
                while it.MoveNext():
                    d_Definition = it.Key
                    d_Name = it.Key.Name
                    d_Binding = it.Current
                    if d_Name == par_name:
                        d_Definition.SetAllowVaryBetweenGroups(doc, True)
    except: pass

class PickParameter(Form):
    def __init__(self): 
        self.Name = "Параметр"
        self.Text = "Выберите параметр"
        self.Size = Size(205, 110)
        self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
        self.button_apply = Button(Text = "Применить")
        self.combo_box = ComboBox()
        self.ControlBox = False
        self.TopMost = True
        self.MinimumSize = Size(205, 110)
        self.MaximumSize = Size(205, 110)
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow
        self.ShowInTaskbar = False
        self.CenterToScreen()

        self.combo_box.Parent = self
        self.combo_box.Items.Add("<по умолчанию>")
        self.combo_box.Text = "<по умолчанию>"
        self.combo_box.DropDownStyle = ComboBoxStyle.DropDownList
        self.combo_box.Location = Point(12, 12)
        self.combo_box.Size = Size(166, 21)
        self.paramlist = []
        for element in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements():
            fam_name = element.Symbol.FamilyName
            if fam_name.startswith("199_Отверстие"):
                for j in element.Parameters:
                    if j.IsShared and j.UserModifiable and j.StorageType == DB.StorageType.String and not j.IsReadOnly:
                        self.paramlist.append(j.Definition.Name)
                break
        self.paramlist.sort()
        for i in self.paramlist:
                self.combo_box.Items.Add(i)
        self.button_apply.Parent = self
        self.button_apply.Location = Point(12, 40)
        self.button_apply.Size = Size(75, 23)
        self.button_apply.Text = "Применить"
        self.button_apply.Click += self.Commit

    def Commit(self, sender, args):
        global write_parameter
        write_parameter = self.combo_box.Text
        self.Close()

class CreateWindow(Form):
    def __init__(self): 
        self.Name = "Связи для маркировки"
        self.Text = "Выберите связи с общей маркировкой"
        self.Size = Size(418, 608)
        self.Icon = Icon("Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
        self.button_create = Button(Text = "Ок")
        self.ControlBox = True
        self.TopMost = True
        self.MinimumSize = Size(418, 480)
        self.MaximumSize = Size(418, 480)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.CenterToScreen()

        self.listbox = ListView()
        self.c_cb = ColumnHeader()
        self.c_cb.Text = ""
        self.c_cb.Width = -2
        self.c_cb.TextAlign = HorizontalAlignment.Center
        self.c_name = ColumnHeader()
        self.c_name.Text = "Путь"
        self.c_name.Width = -2
        self.c_name.TextAlign = HorizontalAlignment.Left

        self.SuspendLayout()
        self.listbox.Dock = DockStyle.Fill
        self.listbox.View = View.Details

        self.listbox.Parent = self
        self.listbox.Size = Size(400, 400)
        self.listbox.Location = Point(1, 1)
        self.listbox.FullRowSelect = True
        self.listbox.GridLines = True
        self.listbox.AllowColumnReorder = True
        self.listbox.Sorting = SortOrder.Ascending
        self.listbox.Columns.Add(self.c_cb)
        self.listbox.Columns.Add(self.c_name)
        self.listbox.LabelEdit = True
        self.listbox.CheckBoxes = True
        self.listbox.MultiSelect = True

        self.button_ok = Button(Text = "Ok")
        self.button_ok.Parent = self
        self.button_ok.Location = Point(10, 410)
        self.button_ok.Click += self.OnOk

        self.button_ok = Button(Text = "Отмена")
        self.button_ok.Parent = self
        self.button_ok.Location = Point(100, 410)
        self.button_ok.Click += self.OnCancel

        self.item = []


        for link in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements():
            try:
                document = link.GetLinkDocument()
                transform = link.GetTransform()
                title = document.Title
                self.item.append(ListViewItem())
                self.item[len(self.item)-1].Text = ""
                self.item[len(self.item)-1].Checked = False
                self.item[len(self.item)-1].SubItems.Add("{} ({})".format(document.Title, document.PathName))
                self.listbox.Items.Add(self.item[len(self.item)-1])
            except: pass

    def OnOk(self, sender, args):
        global paths
        list_of_levels = []
        for i in self.item:
            if i.Checked:
                si = i.SubItems
                paths.append(si[1].Text)
        self.Close()

    def OnCancel(self, sender, args):
        global next
        next = False
        self.Close()


class SSymbol():
    def __init__(self, name, width, height, offset, elevation, element, host, islink):
        self.Type = name
        self.Width = self.zero(width)
        self.Height = self.zero(height)
        self.Offset = self.zero(offset)
        self.LevelElev = elevation
        self.Host = ""
        if host:
            self.Host = "1"
        else:
            self.Host = "0"
        self.Element = element
        self.Key = "{}_{}_{}_{}_{}".format(self.Host, self.Type, self.Height, self.Width, self.Offset)
        self.Link = islink


    def IsLink(self):
        return self.Link

    def SetKey(self, k):
        self.Key = k

    def zero(self, integer):
        newint = math.fabs(integer)
        z = ""
        for i in range(0, 10 - len(str(newint))):
            z += "0"
        z += str(newint)
        if integer < 0:
            z = "-{}".format(z)
        return z

    @property
    def Key(self):
        return self.Key

    @property
    def Host(self):
        return self.Host

    @property
    def Offset(self):
        return self.Offset

    @property
    def Type(self):
        return self.Type

    @property
    def Width(self):
        return self.Width

    @property
    def Height(self):
        return self.Height

    @property
    def LevelElev(self):
        return self.LevelElev

    @property
    def Element(self):
        return self.Element


def InList(item, list):
    try:
        for i in list:
            if i == item:
                return True
        return False
    except: return False


# Запись в класс SSymbol
def seterSSymbol(doc, isLink):
    collector_elements = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()
    for element in collector_elements:
        fam_name = element.Symbol.FamilyName
        if fam_name.startswith("199_Отверстие"):
            try:
                name = "{}_{}".format(element.Symbol.get_Parameter(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString(), element.Symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString())
                width = int(round(element.get_Parameter(guid_width).AsDouble() * 3.048, 1) * 100)
                height = int(round(element.get_Parameter(guid_height).AsDouble() * 3.048, 1) * 100)
                offset = int(round(element.get_Parameter(DB.BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsDouble() * 3.048, 1) * 100)
                elevation = int(round(doc.GetElement(element.LevelId).Elevation * 3.048, 1) * 100)
                host = element.Host.Name.startswith("00_")
                symbols.append(SSymbol(name, width, height, offset, elevation, element, host, isLink))
            except Exception as e_add:
                print(str(e_add) + "!!!")

paths = []
documents = []
# Ширина
guid_width = Guid("8f2e4f93-9472-4941-a65d-0ac468fd6a5d")
# Высота
guid_height = Guid("da753fe3-ecfa-465b-9a2c-02f55d0c2ff1")


write_parameter = "None"
commands = [CommandLink('Да', return_value=True), CommandLink('Нет', return_value=False), CommandLink('Отмена', return_value="Отмена")]
dialog = TaskDialog('Учитывать маркировку подгруженных связей во время маркировки?',
                    title = "Учет связей",
                    title_prefix=False,
                    content="Опция позволяет создавать общую систему марок для нескольких проектов.",
                    commands=commands,
                    footer='',
                    show_close=False)

dialog_par = TaskDialog('Использовать параметр «по умолчанию» для записи значения марки?',
                    title = "Параметр для записи",
                    title_prefix=False,
                    content="",
                    commands=commands,
                    footer='',
                    show_close=False)
ShowForm = dialog.show()
if ShowForm != "Отмена":
    next = True
    if ShowForm:
        form = CreateWindow()
        Application.Run(form)
    if next:
        ParamForm = dialog_par.show()
        if not ParamForm:
            form2 = PickParameter()
            Application.Run(form2)
        if len(paths) != 0:
            for name in paths:
                for link in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements():
                    try:
                        document = link.GetLinkDocument()
                        if "{} ({})".format(document.Title, document.PathName) == name:
                            documents.append(document)
                    except:
                        pass

        AddParameter("00_Комментарий")
        AddParameter("00_Фасад")
        # collector_elements = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements()
        symbols = []
        link_symbols = []

        value = str(TextInput('Префикс для маркировки', default="", description = "«[префикс][марка]»", exit_on_close = False))
        dict = []
        sub_elements = []

        if len(documents) != 0:
            for document in documents:
                try:
                    seterSSymbol(document, True)
                except Exception as e:
                    continue

        seterSSymbol(doc, False)
        # Добавли сортировку по уровню
        symbols = sorted (symbols, key=lambda s: s.LevelElev)

        # Упорядоченный словарь с экземплярами класса SSymbol, разделенные
        # по заданным условиям. Ключ - условия группирования, значения - 
        # все отверстия под данной группой
        typesDict = OrderedDict()
        for symbol in symbols:
            dictKey = str(symbol.LevelElev) +\
                "~" +\
                symbol.Width.ToString() +\
                "~" +\
                symbol.Height.ToString() +\
                "~" +\
                symbol.Offset.ToString()
            try:
                typesDict[dictKey].append(symbol)
            except:
                typesDict[dictKey] = [symbol]

        with db.Transaction("KPLN. Нумерация отверстий"):
            tempKey = "~"
            for key, symbols in typesDict.items():
                # Нумерация должна быть ВНУТРИ этажа. Из-за структуры (словарь)
                # нужно обнулять счетчик при смене этажа в key (часть имени ключа)
                if key.split("~")[0] != tempKey:
                    cnt = 0
                    tempKey = key.split("~")[0]
                cnt += 1
                for symbol in symbols:
                    if not symbol.IsLink():
                        if write_parameter == "None" or write_parameter == "<по умолчанию>":
                            symbol.Element.get_Parameter(DB.BuiltInParameter.DOOR_NUMBER).Set(value + str(cnt))
                        else:
                            symbol.Element.LookupParameter(write_parameter).Set(value + str(cnt))
                        if symbol.Element.Host.Name.startswith("00_"):
                            symbol.Element.LookupParameter("00_Комментарий").Set("см. раздел КЖ")
                        else:
                            symbol.Element.LookupParameter("00_Комментарий").Set("")
                        symbol.Element.LookupParameter("00_Фасад").Set(symbol.LevelElev.ToString())