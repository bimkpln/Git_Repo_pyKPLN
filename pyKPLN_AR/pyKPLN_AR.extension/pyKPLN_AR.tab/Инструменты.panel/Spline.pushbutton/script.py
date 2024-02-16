# -*- coding: utf-8 -*-
"""
SplineNumber

"""
__author__ = 'GF'
__title__ = "Нумерация\nпо сплайну"
__doc__ = 'Нумерация помещений или машиномест по сплайну с указанием префикса, постфикса и начального номера' \

"""
Архитекурное бюро KPLN

"""
import clr
import webbrowser
from rpw import doc, uidoc, DB, UI, db
from pyrevit import DB, UI
from System.Collections.Generic import *
from rpw.ui.forms import Alert, TaskDialog, CommandLink
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
    recomm_param = '"ПОМ_Номер помещения"'
if dialog_out == 'parking':
    category = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Parking).WhereElementIsNotElementType().FirstElement()
    recomm_param = '"КП_О_Позиция" (ADSK_Позиция)'

class Input(Form):
    UserInput = ''
    
    def __init__(self):
        self.Size = Size(400, 200)
        self.Text = 'KPLN Нумерация по сплайну'
        self.Icon = Icon('X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico')
        self.CenterToScreen()

        self._lbl = Label(Text = 'Укажите порядковый номер параметра "ПОМ_Номер помещения"\n(см. настройки квартирографии - "Формула номера помещения")')
        self._lbl.Size = Size(380, 60)
        self._lbl.Location = Point(20, 20)
        self._lbl.Parent = self

        self._textbox = TextBox()
        self._textbox.Size = Size(200, 30)
        self._textbox.Location = Point(100, 80)
        self.Controls.Add(self._textbox)

        self._btn_ok = Button()
        self._btn_ok.Size = Size(80, 25)
        self._btn_ok.Location = Point(160, 120)
        self._btn_ok.Text = 'OK'
        self.Controls.Add(self._btn_ok)
        self._btn_ok.Click += self._on_button_ok_click

    def _on_button_ok_click(self, sender, event_args):
        self.UserInput = self._textbox.Text
        self.Close()

class CreateWindow(Form):
    Prefix = ''
    Postfix = ''
    StartNumber = ''
    LabelText = '<Префикс><Стартовый номер><Постфикс>'

    def __init__(self): 
        self.Name = "KPLN Нумерация по сплайну"
        self.Text = "KPLN Нумерация по сплайну"
        self.Size = Size(418, 608)
        self.Icon = Icon("X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
        self.button_create = Button(Text = "Ок")
        self.tooltip = ToolTip()
        self.ControlBox = True
        self.TopMost = True
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.MinimumSize = Size(418, 540)
        self.MaximumSize = Size(418, 540)
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
        self.lb1 = Label(Text = "1. Выберите параметр для записи марки")
        self.lb1.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.lb1.Parent = self
        self.lb1.Size = Size(343, 20)
        self.lb1.Location = Point(30, 30)

        self.lb2 = Label(Text = "Рекомендуется использовать " + recomm_param)
        self.lb2.Parent = self
        self.lb2.Size = Size(343, 30)
        self.lb2.Location = Point(30, 85)

        self.lb3 = Label(Text = "2. Укажите префикс (при необходимости):")
        self.lb3.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.lb3.Parent = self
        self.lb3.Size = Size(343, 20)
        self.lb3.Location = Point(30, 125)
        self.textbox1 = TextBox()
        self.textbox1.Parent = self
        self.textbox1.Size = Size(120, 30)
        self.textbox1.Location = Point(30, 150)
        self.textbox1.TextChanged += self.updateLabelText

        self.lb4 = Label(Text = "3. Укажите стартовый номер:")
        self.lb4.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.lb4.Parent = self
        self.lb4.Size = Size(343, 20)
        self.lb4.Location = Point(30, 195)
        self.textbox2 = TextBox()
        self.textbox2.Parent = self
        self.textbox2.Size = Size(120, 30)
        self.textbox2.Location = Point(30, 220)
        self.textbox2.TextChanged += self.updateLabelText

        self.lb5 = Label(Text = "4. Укажите постфикс (при необходимости):")
        self.lb5.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.lb5.Parent = self
        self.lb5.Size = Size(343, 20)
        self.lb5.Location = Point(30, 265)
        self.textbox3 = TextBox()
        self.textbox3.Parent = self
        self.textbox3.Size = Size(120, 30)
        self.textbox3.Location = Point(30, 290)
        self.textbox3.TextChanged += self.updateLabelText

        self.lb6 = Label()
        self.lb6.Text = self.labelText()
        self.lb6.Parent = self
        self.lb6.Size = Size(343, 50)
        self.lb6.Location = Point(40, 355)
        self.lb6.Font = Font("Segoe UI", 11)
        self.lb6.ForeColor = Color.FromName('HotPink')

        self.btn_confirm = Button(Text = "Запуск")
        self.btn_confirm.Parent = self
        self.btn_confirm.Location = Point(110, 440)
        self.btn_confirm.Size = Size(140, 40)
        self.btn_confirm.Enabled = False
        self.btn_confirm.Click += self.OnOk

        self.btn_i = Button(Text = "?")
        self.btn_i.Parent = self
        self.btn_i.Location = Point(255, 440)
        self.btn_i.Size = Size(40, 40)
        self.btn_i.Font = Font("Arial", 14, FontStyle.Bold)
        self.btn_i.BackColor = Color.FromArgb(255,218,185)
        self.btn_i.ForeColor = Color.FromArgb(250,0,0)
        self.btn_i.Click += self.go_to_help
        self.tooltip.SetToolTip(self.btn_i, "Открыть инструкцию к плагину")

        self.cb = ComboBox()
        self.cb.Parent = self
        self.cb.Size = Size(343, 100)
        self.cb.Location = Point(30, 55)
        self.cb.DropDownStyle = ComboBoxStyle.DropDownList
        self.cb.SelectedIndexChanged += self.allow_ok
        self.cb.Items.Add("Выберите параметр")
        self.cb.Text = "Выберите параметр"
        for par in self.params:
            self.cb.Items.Add(par)

    def updateLabelText(self, sender, args):
        if sender == self.textbox1:
            self.Prefix = sender.Text
        if sender == self.textbox2:
            self.StartNumber = sender.Text
        if sender == self.textbox3:
            self.Postfix = sender.Text
        self.lb6.Text = self.labelText()
        if self.labelText() == '<Префикс><Стартовый номер><Постфикс>':
            self.lb6.Font = Font("Segoe UI", 11)
        else:
            self.lb6.Font = Font("Segoe UI", 11, FontStyle.Bold)

    def labelText(self):
        if self.Prefix + self.StartNumber + self.Postfix != '':
            self.LabelText = self.Prefix + self.StartNumber + self.Postfix
        else:
            self.LabelText = '<Префикс><Стартовый номер><Постфикс>'
        return self.LabelText

    def go_to_help(self, sender, args):
        webbrowser.open('http://moodle/mod/book/view.php?id=502&chapterid=1286/')

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
        normal_faces = []
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
                    if face.FaceNormal.X == 0 and face.FaceNormal.Y == 0:
                        normal_faces.append(face)
        elements_faces.append([element, normal_faces])
    return elements_faces

def ReNumber(room, index, number):
    numParam = room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
    numParamList = numParam.split(".")
    newNum = ""
    strLen = len(numParamList)
    if index < strLen:
        for i in range(strLen):
            if i == index:
                newNum += str(number)
                if i == strLen-1:
                    continue
                else:
                    newNum += "."
            elif i < strLen-1:
                newNum += str(numParamList[i])
                newNum += "."
            else:
                newNum += str(numParamList[i])
        room.get_Parameter(BuiltInParameter.ROOM_NUMBER).Set(newNum)
    else:
        raise Exception("Индекс выше положенного! Работа остановлена!")

list_elements = []
list_elements_sorted = []
list_parameters = []
list_parameters_sorted = []
points = []
faces = []
parameter_name = "Марка"

try:
    otype = UI.Selection.ObjectType.Element
    filter_spline = SplineSelectionFilter()
    if dialog_out == 'rooms':
        selectionFilter = RoomSelectionFilter()
        text = "Выберите помещения"
    if dialog_out == 'parking':
        selectionFilter = AutoMobileSelectionFilter()
        text = "Выберите машиноместа"
    spline = uidoc.Selection.PickObject(otype, filter_spline, "KPLN : Веберите сплайн")
    Alert("и нажмите кнопку «Готово»", title="KPLN Нумерация по сплайну", header = text)
    elements = uidoc.Selection.PickObjects(otype, selectionFilter, "KPLN : " + text)
    if spline and len(elements) != 0:
        points = GetPoints(spline)
        element_faces = GetFaces(elements)
        for i in range(0, len(points)):
            for faces in element_faces:
                for f in faces[1]:
                    s = f.GetSurface()
                    uv = s.Project(points[i])[0]
                    if f.IsInside(uv):
                        list_elements.append([i, faces[0]])
                        break # Если хотя бы на одну плоскость удается спроецировать точку - цикл прерывается

        if len(points) < len(list_elements):
            Alert("Точек в сплайне меньше, чем выбранных элементов. Возможна некорректная последовательность в нумерации.", title="KPLN Нумерация по сплайну", header = "Внимание!")
        elif len(points) > len(list_elements):
            Alert("Точек в сплайне больше, чем выбранных элементов. Возможна некорректная последовательность в нумерации.", title="KPLN Нумерация по сплайну", header = "Внимание!")
        list_elements.sort()

        form = CreateWindow()
        Application.Run(form)
        prefix = form.Prefix
        postfix = form.Postfix
        startnumber = form.StartNumber
        try:
            num = int(startnumber)
        except ValueError:
            Alert("Не удалось определить стартовый номер, будет использовано значение по умолчанию «1».", title="KPLN Нумерация по сплайну", header = "Ошибка")
        parameter_name = form.value

        index = None
        if parameter_name == "ПОМ_Номер помещения":
            commands= [CommandLink('Да', return_value=True),
                        CommandLink('Отмена', return_value=False)]
            dialog = TaskDialog('Перезаписать параметр "Номер"?',
                        title="KPLN Нумерация по сплайну",
                        title_prefix=False,
                        commands=commands,
                        show_close=True)
            dialog_out = dialog.show()
            if dialog_out:
                input = Input()
                input.ShowDialog()
                index = int(input.UserInput) - 1

        with db.Transaction(name = "mark"):
            for element in list_elements:
                value = prefix + str(num) + postfix
                element[1].LookupParameter(parameter_name).Set(value)
                num += 1
                if index != None:
                    ReNumber(element[1], index, value)
except Exception as e:
    print(e)