# -*- coding: utf-8 -*-
"""
MarkAuto

"""
__author__ = 'GF'
__title__ = "Нумерация по сплайну"
__doc__ = 'Маркировка машиномест по сплайну с указанием префикса, постфикса и начального номера' \

"""
Архитекурное бюро KPLN

"""
import clr
import math
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from rpw.ui import Pick
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import TextInput, Alert, TaskDialog, CommandLink
import datetime
from Autodesk.Revit.DB import *
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import *
clr.AddReference('System.Drawing')
from System.Drawing import *

# Окно выбора
commands= [CommandLink('Помещения', return_value='rooms'),
            CommandLink('Парковочные места', return_value='parking')]
dialog = TaskDialog('Выберите категорию элементов, которые требуется пронумеровать:',
               title="KPLN Нумерация по сплайну",
               title_prefix=False,
               commands=commands,
               show_close=True)
dialog_out = dialog.show()

if dialog_out == 'rooms':
    category = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().FirstElement()
if dialog_out == 'parking':
    category = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Parking).WhereElementIsNotElementType().FirstElement()

class CreateWindow(Form):
    def __init__(self): 
        self.Name = "KPLN Нумерация по сплайну"
        self.Text = "KPLN Нумерация по сплайну"
        self.Size = Size(418, 608)
        self.Icon = Icon("X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
        self.button_create = Button(Text = "Ок")
        self.ControlBox = True
        self.TopMost = True
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.MinimumSize = Size(418, 240)
        self.MaximumSize = Size(418, 240)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.CenterToScreen()
        self.BackColor = Color.FromArgb(255, 255, 255)
        self.value = "Выберите параметр"
        self.el = category
        self.params = []
        for j in self.el.Parameters:
            # Исключение для Revit 2023
            try:
                if str(j.Definition.ParameterType) == "Text":
                    self.params.append(j.Definition.Name)
            except:
                if j.Definition.GetDataType() == SpecTypeId.String.Text:
                    self.params.append(j.Definition.Name)
        self.params.sort()
        self.lb1 = Label(Text = "Выберите параметр для записи марки")
        self.lb1.Parent = self
        self.lb1.Size = Size(343, 15)
        self.lb1.Location = Point(30, 30)

        self.lb2 = Label(Text = "Необходимо выбрать текстовый параметр.\nРекомендуется использовать системный параметр «Марка».")
        self.lb2.Parent = self
        self.lb2.Size = Size(343, 40)
        self.lb2.Location = Point(30, 75)

        self.btn_confirm = Button(Text = "Далее")
        self.btn_confirm.Parent = self
        self.btn_confirm.Location = Point(30, 150)
        self.btn_confirm.Enabled = False
        self.btn_confirm.Click += self.OnOk

        self.cb = ComboBox()
        self.cb.Parent = self
        self.cb.Size = Size(343, 100)
        self.cb.Location = Point(30, 50)
        self.cb.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb.SelectedIndexChanged += self.allow_ok
        self.cb.Items.Add("Выберите параметр")
        self.cb.Text = "Выберите параметр"
        for par in self.params:
            self.cb.Items.Add(par)

    def OnOk(self, sender, args):
        self.Close()

    def allow_ok(self, sender, event):
        if self.cb.Text != "Выберите параметр":
            self.btn_confirm.Enabled = True
        else:
            self.btn_confirm.Enabled = False
        self.value = self.cb.Text

class AutoMobileSelectionFilter (UI.Selection.ISelectionFilter):
    def AllowElement(self, element = DB.Element):
        if element.Category.Id == DB.ElementId(-2001180) and ("парков" in str(element.Symbol.FamilyName).lower() or "парков" in str(element.Name).lower()):
            return True
        return False
    def AllowReference(self, refer, point):
        return False

class RoomSelectionFilter (UI.Selection.ISelectionFilter):
    def AllowElement(self, element = DB.Element):
        if element.Category.Id == DB.ElementId(-2000160):
            return True
        return False
    def AllowReference(self, refer, point):
        return False

class SplineSelectionFilter (UI.Selection.ISelectionFilter):
    def AllowElement(self, element = DB.Element):
        if element.Category.Id == DB.ElementId(-2000051):
            return True
        return False
    def AllowReference(self, refer, point):
        return False

def GetPoints(spline):
    spline_points = []
    line = doc.GetElement(spline)
    nurbSpline = line.Location.Curve
    spline_points = nurbSpline.CtrlPoints
    return spline_points

def GetFaces(elements):
    elements_faces = []
    for e in elements:
        element = doc.GetElement(e)
        opt = Options()
        opt.DetailLevel = ViewDetailLevel.Fine
        elemGeom = element.get_Geometry(opt)
        for gObj in elemGeom:
            if type(gObj) == Solid:
                solid = gObj
            else:
                geomInsts = gObj.GetInstanceGeometry()
                for solid in geomInsts:
                    if solid.Volume == 0: continue
            faces = solid.Faces
            for face in faces:
                if type(face) == PlanarFace:
                    if face.FaceNormal.X == 0 and face.FaceNormal.Y == 0 and face.FaceNormal.Z == -1:
                        elements_faces.append([element, face])
    return elements_faces

list_elements = []
list_elements_sorted = []
list_parameters = []
list_parameters_sorted = []
points = []
faces = []
parameter_name = "Марка"
testList = []



try:
    Alert("1. Выбрать линию\n\n2. Выбрать машиноместа\n\n3. Указать остальные необходимые параметры:\n   - префикс, постфикс, начало нумерации", title="KPLN Нумерация по сплайну", header = "Инструкция:")
    otype = UI.Selection.ObjectType.Element
    filter_spline = SplineSelectionFilter()
    if dialog_out == 'rooms':
        selectionFilter = RoomSelectionFilter()
        text = "Веберите помещения"
    if dialog_out == 'parking':
        selectionFilter = AutoMobileSelectionFilter()
        text = "Веберите машиноместа"
    spline = uidoc.Selection.PickObject(otype, filter_spline, "KPLN : Веберите сплайн")
    Alert("и нажмите кнопку «Готово»", title="KPLN Нумерация по сплайну", header = text)
    elements = uidoc.Selection.PickObjects(otype, selectionFilter, "KPLN : " + text)
    if spline and len(elements) != 0:
        try:
            with db.Transaction(name = "mark"):
                points = GetPoints(spline)
                faces = GetFaces(elements)
                for i in range(0, len(points)):
                    index = i+1
                    for f in faces:
                        s = f[1].GetSurface()
                        uv = s.Project(points[i])[0]
                        if f[1].IsInside(uv):
                            list_elements.append([i, f[0]])
        except Exception as e:
            print(e)
        if len(points) < len(list_elements):
            Alert("Точек в сплайне меньше, чем выбранных элементов. Возможна некорректная последовательность в нумерации.", title="KPLN Нумерация по сплайну", header = "Внимание!")
        elif len(points) > len(list_elements):
            Alert("Точек в сплайне больше, чем выбранных элементов. Возможна некорректная последовательность в нумерации.", title="KPLN Нумерация по сплайну", header = "Внимание!")
        list_elements.sort()
        num = 1
        # Возможность добавления префикса и постфикса (отключено, т.к. не используется)
        # prefix = TextInput('Префикс для маркировки', default = "", description = "Укажите префикс: <префикс><марка>", exit_on_close = False)
        # postfix = TextInput('Постфикс для маркировки', default = "", description = "Укажите постфикс: <марка><постфикс>", exit_on_close = False)
        startnumber = TextInput('Начать нумерацию со значения', default = "0", description = "Примечание: только числовое значение!", exit_on_close = False)
        try:
            num = int(startnumber)
        except:
            Alert("Не удалось определить стартовый номер, будет использовано значение по умолчанию «1»", title="KPLN Нумерация по сплайну", header = "Ошибка")
            num = 1
        try:
            form = CreateWindow()
            Application.Run(form)
            parameter_name = form.value
        except Exception as e:
            print(e)
        with db.Transaction(name = "mark"):
            for element in list_elements:
                try:
                    # element[1].LookupParameter(parameter_name).Set(prefix + str(num) + postfix)
                    element[1].LookupParameter(parameter_name).Set(str(num))
                    num += 1
                except Exception as e:
                    print(e)
except Exception as e:
    print(e)