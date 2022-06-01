# -*- coding: utf-8 -*-
"""
SelectByLevel

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Выбрать по уровню"
__doc__ = 'Выбирает все элементы на выбранных уровнях' \

"""
Архитекурное бюро KPLN

"""
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, \
    ElementId
from rpw import doc, uidoc
from System.Collections.Generic import List
from System.Windows.Forms import Button, FormBorderStyle, ListView,\
    ColumnHeader, HorizontalAlignment, DockStyle, View, SortOrder,\
    ListViewItem, Application, Form
from System.Drawing import Size, Icon, Point

class CreateWindow(Form):
    def __init__(self):
        self.Name = "Редактор наборов"
        self.Text = "Выберите уровни"
        self.Size = Size(418, 608)
        self.Icon = Icon("X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
        self.button_create = Button(Text="Ок")
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
        self.c_name.Text = "Имя"
        self.c_name.Text = "Имя"
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

        self.button_ok = Button(Text="Ok")
        self.button_ok.Parent = self
        self.button_ok.Location = Point(10, 410)
        self.button_ok.Click += self.OnOk

        self.button_ok = Button(Text="Закрыть")
        self.button_ok.Parent = self
        self.button_ok.Location = Point(100, 410)
        self.button_ok.Click += self.OnCancel

        self.Text = "Выберите уровни для выбора"
        self.levels = []
        self.levels_sorted = []
        self.levels_sort = []
        self.levels_sort_elevation = []
        self.levels_names = []
        self.collector_levels = FilteredElementCollector(doc).\
            OfCategory(BuiltInCategory.OST_Levels).\
            WhereElementIsNotElementType().\
            ToElements()
        self.item = []
        for level in self.collector_levels:
            self.levels_sort.append(level)
            self.levels_sort_elevation.append(level.Elevation)
        self.levels_sort_elevation.sort()
        self.levels_sort_elevation.reverse()
        for i in self.levels_sort_elevation:
            for level in self.levels_sort:
                if i == level.Elevation:
                    self.levels_sorted.append(level)
                    break
        for level in self.levels_sorted:
            try:
                self.levels.append(level)
                name = level.Name
                self.levels_names.append(name)
                self.item.append(ListViewItem())
                self.item[len(self.item)-1].Text = ""
                self.item[len(self.item)-1].Checked = False
                self.item[len(self.item)-1].SubItems.Add(name)
                self.listbox.Items.Add(self.item[len(self.item)-1])
            except Exception: 
                pass

    def OnOk(self, sender, args):
        self.select_lvl_elements(self.define_levels())
        self.Close()

    def define_levels(self):
        list_of_levels = []
        for i in self.item:
            if i.Checked:
                si = i.SubItems
                viewname = si[1].Text
                for level in self.levels:
                    if viewname == level.Name:
                        list_of_levels.append(level.Id)
        return list_of_levels

    def OnCancel(self, sender, args):
        self.Close()

    def select_lvl_elements(self, levels):
        if len(levels) != 0:
            elements = []
            active_view = doc.ActiveView.Id
            collector = FilteredElementCollector(doc, active_view).\
                WhereElementIsNotElementType().\
                ToElements()
            for element in collector:
                for id in levels:
                    if str(element.LevelId) == str(id):
                        elements.append(element)
                        break
            group_collector = FilteredElementCollector(doc, active_view).\
                OfCategory(BuiltInCategory.OST_IOSModelGroups).\
                WhereElementIsNotElementType().\
                ToElements()
            for group in group_collector:
                group_elements = group.GetMemberIds()
                for element_id in group_elements:
                    element = doc.GetElement(element_id)
                    for id in levels:
                        try:
                            if str(element.LevelId) == str(id):
                                elements.append(element)
                                break
                        except Exception:
                            pass
                        try:
                            if str(element.Host.Id) == str(id):
                                elements.append(element)
                                break
                        except Exception: 
                            pass
            if len(elements) != 0:
                selection = uidoc.Selection
                collection = List[ElementId](
                    [element.Id for element in elements]
                )
                selection.SetElementIds(collection)


form = CreateWindow()
Application.Run(form)
