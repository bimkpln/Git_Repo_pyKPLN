# -*- coding: utf-8 -*-
"""
CopyCropView

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Копировать\nподрезку"
__doc__ = 'Копировать подрезку активного вида на выбранные виды' \

"""
Архитекурное бюро KPLN

"""
from pyrevit.framework import clr
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import forms
from pyrevit import revit, DB, UI
from System.Collections.Generic import *
from System.Collections.Generic import *
from System.Windows.Forms import *
from System.Drawing import *
import datetime
import datetime


class CreateWindow(Form):
    def __init__(self): 
        self.Name = "Подрезать виды"
        self.Text = "Выберите виды для подрезки"
        self.Size = Size(418, 608)
        self.Icon = Icon("X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
        self.button_create = Button(Text = "Ок")
        self.ControlBox = True
        self.TopMost = True
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.MinimumSize = Size(418, 480)
        self.MaximumSize = Size(418, 480)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.CenterToScreen()
        self.views = ""

        self.selection = revit.get_selection()

        self.listbox = ListView()
        self.c_cb = ColumnHeader()
        self.c_cb.Text = ""
        self.c_cb.Width = -2
        self.c_cb.TextAlign = HorizontalAlignment.Center
        self.c_name = ColumnHeader()
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

        self.active_view = doc.ActiveView
        self.crop_boundary = active_view.GetCropRegionShapeManager().GetCropShape()

        self.button_ok = Button(Text = "Ok")
        self.button_ok.Parent = self
        self.button_ok.Location = Point(10, 410)
        self.button_ok.Click += self.OnOk

        self.button_cancel = Button(Text = "Закрыть")
        self.button_cancel.Parent = self
        self.button_cancel.Location = Point(100, 410)
        self.button_cancel.Click += self.OnCancel

        self.view_type = str(self.active_view.ViewType)
        self.Text = "Выберите виды для подрезки (" + self.view_type + ")"
        self.plans = []
        self.plans_names = []
        self.collector_views = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Views)
        self.item = []
        for view in self.collector_views:
            if str(view.ViewType) == self.view_type and not view.IsTemplate and str(view.Name) != str(self.active_view.Name) and str(view.Name) not in self.plans_names:
                self.plans.append(view)
                self.plans_names.append(view.Name)
                self.item.append(ListViewItem())
                self.item[len(self.item)-1].Text = ""
                self.item[len(self.item)-1].Checked = False
                self.item[len(self.item)-1].SubItems.Add(view.Name)
                self.listbox.Items.Add(self.item[len(self.item)-1])

    def OnOk(self, sender, args):
        with db.Transaction(name='КП_Обрезка видов'):
            self.define_views()
        self.Close()
    def OnCancel(self, sender, args):
        self.Close()

    def define_views(self):
        for i in self.item:
            if i.Checked:
                si = i.SubItems
                viewname = si[1].Text
                for v in self.plans:
                    if viewname == v.Name:
                        self.run(v)

    def run(self, view):
        view.CropBoxActive = True
        view.GetCropRegionShapeManager().SetCropShape(self.crop_boundary[0])
        view.CropBoxActive = False
        self.views += "     {};\n".format(view.Title)

active_view = doc.ActiveView
if str(active_view.ViewType) == "FloorPlan" or str(active_view.ViewType) == "EngineeringPlan" or str(active_view.ViewType) == "AreaPlan":
    form = CreateWindow()
    Application.Run(form)
else:
    forms.alert("Ошибка: активный вид не является планом!\nТип активного вида: <{}>".format(str(active_view.ViewType)))
