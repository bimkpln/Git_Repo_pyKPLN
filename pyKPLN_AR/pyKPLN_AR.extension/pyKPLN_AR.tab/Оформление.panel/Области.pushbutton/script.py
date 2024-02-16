# -*- coding: utf-8 -*-
"""
FilledRegionCreation

"""
__author__ = 'Gyulnara Fedoseeva'
__title__ = "Цветовые\nобласти"
__doc__ = 'Создание цветовых областей на основе выбранной палитры'\

"""
Архитекурное бюро KPLN

"""

import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import *
clr.AddReference('System.Drawing')
import webbrowser
from System.Drawing import *
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import revit as REVIT
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from Autodesk.Revit.DB import BuiltInParameter,\
    FilteredElementCollector, BuiltInCategory, Transaction,\
    Category, ElementLevelFilter, XYZ, Line, ElementTransformUtils, LinePatternElement, GraphicsStyleType, FilledRegion, FilledRegionType, ElementType, Element
from Autodesk.Revit.DB import Color as RColor
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import forms
from pyrevit import script
from libKPLN import color_palette

succes = bool
categories = doc.Settings.Categories
category = categories.get_Item(BuiltInCategory.OST_Lines)
detail_lines = DB.FilteredElementCollector(doc).OfClass(DB.LinePatternElement).ToElements()

#Параметры для логирования в Extensible Storage. Не менять по ходу изменений кода
extStorage_guid = "720080C5-DA99-40D7-9445-E53F288AA180"
extStorage_name = "kpln_ar_color_palette"

try:
    extStorage = color_palette.create_obj(extStorage_guid, extStorage_name)
    extText = color_palette.read_log(extStorage)
    extColors = extText.split('|')[1:]
except:
    extStorage = None

example = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType).FirstElement()
elemId = example.Id
elemType = doc.GetElement(elemId)

fillpatterns = DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement).ToElements()

class DblButton(Button):
    def __init__(self):
        self.SetStyle(ControlStyles.StandardClick, True)
        self.SetStyle(ControlStyles.StandardDoubleClick, True)

class ColorButton():
    Btn = None
    color = Color
    name = ''

    def __init__(self, btn, name, color):
        self.Btn = btn
        self.name = name
        self.color = color

    def initBtn(self, click_event, dblclick_event):
        self.Btn.Size = Size(164, 30)
        self.Btn.Text = self.name
        self.Btn.BackColor = self.color
        self.Btn.Click += click_event
        self.Btn.DoubleClick += dblclick_event

class Input(Form):
    UserInput = ''
    
    def __init__(self):
        self.Size = Size(400, 200)
        self.Text = 'KPLN Создание цветовых областей'
        self.Icon = Icon('X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico')
        self.CenterToScreen()

        self._lbl = Label(Text = 'Введите название цвета (пример: Красный)')
        self._lbl.Size = Size(300, 20)
        self._lbl.Location = Point(10,10)
        self._lbl.Parent = self

        self._textbox = TextBox()
        self._textbox.Size = Size(200, 40)
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

class Window(Form):
    def __init__(self):
        self.Size = Size(730, 500)
        self.Text = 'KPLN Создание цветовых областей'
        self.Icon = Icon('X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico')
        self.CenterToScreen()

        self._tooltip = ToolTip()

        self.Buttons = [] #Список кнопок для удаления
        self.Colors = [] #Список цветов [Name, Color]
        self.Patterns = [] #Список образцов заливки (штриховки)
        self.Weights = [] #Список толщин линий

        self._btn_help = Button()
        self._btn_help.Size = Size(50, 50)
        self._btn_help.Location = Point(50, 20)
        self._btn_help.Text = '?'
        self._btn_help.Font = Font("Arial", 14, FontStyle.Bold)
        self.Controls.Add(self._btn_help)
        self._tooltip.SetToolTip(self._btn_help, "Открыть инструкцию к плагину")
        self._btn_help.Click += self._help

        self.info = Label()
        self.info.Text = 'Плагин создает в проекте цветовые области на основе выбранных свойств.\nНазвание свойств отображается в имени создаваемых областей.\n\nПРИМЕР: Желтый_Сплошная_1'
        self.info.Size = Size(400, 100)
        self.info.Location = Point(110, 20)
        self.info.Parent = self

        # COLORS
        self._btn_palette = Button()
        self._btn_palette.Size = Size(25, 25)
        self._btn_palette.Location = Point(195, 120)
        self._btn_palette.Image = Image.FromFile("X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Оформление.panel\\Палитра.pushbutton\\btn.png")
        self.Controls.Add(self._btn_palette)
        self._tooltip.SetToolTip(self._btn_palette, "Загрузить палитру")
        self._btn_palette.Click += self._on_button_palette_click

        self._layout_1 = FlowLayoutPanel()
        self._layout_1.AutoScroll = True
        self._layout_1.Location = Point(50, 150)
        self._layout_1.BackColor = Color.FromName('White')
        self._layout_1.Size = Size(170, 240)
        self.Controls.Add(self._layout_1)

        self._lbl_1 = Label(Text = 'ЦВЕТ ЗАЛИВКИ')
        self._lbl_1.Size = Size(170, 20)
        self._lbl_1.Location = Point(90, 130)
        self._lbl_1.Parent = self

        self._btn_1_add = Button()
        self._btn_1_add.Size = Size(80, 25)
        self._btn_1_add.Location = Point(50, 400)
        self._btn_1_add.Text = 'Добавить'
        self.Controls.Add(self._btn_1_add)
        self._btn_1_add.Click += self._on_button_color_add_click

        self._btn_1_del = Button()
        self._btn_1_del.Size = Size(80, 25)
        self._btn_1_del.Location = Point(140, 400)
        self._btn_1_del.Text = 'Удалить'
        self.Controls.Add(self._btn_1_del)
        self._btn_1_del.Click += self._on_button_color_del_click

        # PATTERNS
        self._layout_2 = FlowLayoutPanel()
        self._layout_2.Location = Point(270, 150)
        self._layout_2.BackColor = Color.FromName('White')
        self._layout_2.Size = Size(170, 240)
        self.Controls.Add(self._layout_2)

        self._lbl_2 = Label(Text = 'ОБРАЗЕЦ ЗАЛИВКИ')
        self._lbl_2.Size = Size(170, 20)
        self._lbl_2.Location = Point(305, 130)
        self._lbl_2.Parent = self

        self._btn_2_add = Button()
        self._btn_2_add.Size = Size(80, 25)
        self._btn_2_add.Location = Point(270, 400)
        self._btn_2_add.Text = 'Добавить'
        self.Controls.Add(self._btn_2_add)
        self._btn_2_add.Click += self._on_button_pattern_add_click

        self._btn_2_del = Button()
        self._btn_2_del.Size = Size(80, 25)
        self._btn_2_del.Location = Point(360, 400)
        self._btn_2_del.Text = 'Удалить'
        self.Controls.Add(self._btn_2_del)
        self._btn_2_del.Click += self._on_button_pattern_del_click

        # WEIGHTS
        self._layout_3 = FlowLayoutPanel()
        self._layout_3.Location = Point(490, 150)
        self._layout_3.BackColor = Color.FromName('White')
        self._layout_3.Size = Size(170, 240)
        self.Controls.Add(self._layout_3)

        self._lbl_3 = Label(Text = 'ВЕС ЛИНИИ')
        self._lbl_3.Size = Size(170, 20)
        self._lbl_3.Location = Point(540, 130)
        self._lbl_3.Parent = self

        self._btn_3_add = Button()
        self._btn_3_add.Size = Size(80, 25)
        self._btn_3_add.Location = Point(490, 400)
        self._btn_3_add.Text = 'Добавить'
        self.Controls.Add(self._btn_3_add)
        self._btn_3_add.Click += self._on_button_weight_add_click

        self._btn_3_del = Button()
        self._btn_3_del.Size = Size(80, 25)
        self._btn_3_del.Location = Point(580, 400)
        self._btn_3_del.Text = 'Удалить'
        self.Controls.Add(self._btn_3_del)
        self._btn_3_del.Click += self._on_button_weight_del_click

        # CREATE FILLED REGION TYPES
        self._btn_run = Button()
        self._btn_run.Size = Size(150, 40)
        self._btn_run.Location = Point(510, 20)
        self._btn_run.Text = 'Создать цветовые области'
        self.Controls.Add(self._btn_run)
        self._btn_run.Click += self.create

    def _help(self, sender, event_args):
        webbrowser.open('http://moodle/mod/book/view.php?id=502&chapterid=1255')

    def _on_button_palette_click(self, sender, event_args):
        if extStorage != None:
            for i in extColors:
                name = i.split('*')[0]
                color = i.split('*')[1]
                R = int(color.split('-')[0])
                G = int(color.split('-')[1])
                B = int(color.split('-')[2])
                self._btn = ColorButton(DblButton(), name, Color.FromArgb(R,G,B))
                self._btn.initBtn(self._get_selected_button, self._on_button_color_click)
                self._layout_1.Controls.Add(self._btn.Btn)
        else:
            pass

    def _on_button_color_click(self, sender, event_args):
        self.Buttons = []
        self._window = ColorDialog()
        self._window.AllowFullOpen = True
        self._window.ShowHelp = True
        self._window.ShowDialog()
        sender.BackColor = self._window.Color
        colorNameWind = Input()
        colorNameWind.ShowDialog()
        sender.Text = colorNameWind.UserInput

    def _on_button_color_add_click(self, sender, event_args):
        self._btn = ColorButton(DblButton(), 'Выберите цвет заливки', Color.FromName('White'))
        self._btn.initBtn(self._get_selected_button, self._on_button_color_click)
        self._layout_1.Controls.Add(self._btn.Btn)
        self._tooltip.SetToolTip(self._btn.Btn, "Для выбора цвета нажмите 2 раза")

    def _get_selected_button(self, sender, event_args):
        self.Buttons = []
        self.Buttons.append(sender)

    def _on_button_color_del_click(self, sender, event_args):
        for btn in self.Buttons:
            self._layout_1.Controls.Remove(btn)
        self.Buttons = []

    def _on_button_pattern_add_click(self, sender, event_args):
        self._cbbox_pattern = ComboBox()
        self._cbbox_pattern.Size = Size(164, 30)
        self._cbbox_pattern.Text = '<Выберите тип заливки>'
        self._cbbox_pattern.Items.Add('<Выберите тип заливки>')
        for i in fillpatterns:
            self._cbbox_pattern.Items.Add(i.Name)
        self._layout_2.Controls.Add(self._cbbox_pattern)
        self._cbbox_pattern.GotFocus += self._get_selected_button

    def _on_button_pattern_del_click(self, sender, event_args):
        for btn in self.Buttons:
            self._layout_2.Controls.Remove(btn)
        self.Buttons = []

    def _on_button_weight_add_click(self, sender, event_args):
        self._cbbox_weight = ComboBox()
        self._cbbox_weight.Size = Size(164, 30)
        self._cbbox_weight.Text = '<Выберите вес линии>'
        self._cbbox_weight.Items.Add('<Выберите вес линии>')
        for i in range(1, 17):
            self._cbbox_weight.Items.Add(i)
        self._layout_3.Controls.Add(self._cbbox_weight)
        self._cbbox_weight.GotFocus += self._get_selected_button

    def _on_button_weight_del_click(self, sender, event_args):
        for btn in self.Buttons:
            self._layout_3.Controls.Remove(btn)
        self.Buttons = []

    def create(self, sender, event_args):
        for color in self._layout_1.Controls:
            if color.Text != 'Выберите цвет заливки':
                self.Colors.append([color.Text, color.BackColor])
        for item in self._layout_2.Controls:
            if item.Text != '<Выберите тип заливки>':
                for i in fillpatterns:
                    text = item.Text
                    if item.Text == '<Сплошная заливка>':
                        text = 'Сплошная'
                    if item.Text == i.Name:
                        pattern = i.Id
                self.Patterns.append([text, pattern])
        for weight in self._layout_3.Controls:
            if weight.Text != '<Выберите вес линии>':
                self.Weights.append(weight.Text)

        with Transaction(doc, 'Создать цветовые области') as t:
            t.Start()
            for c in self.Colors:
                for p in self.Patterns:
                    for w in self.Weights:
                        Name = 'Х_' + c[0] + '_' + p[0] + '_' + w
                        newFilledRegionType = elemType.Duplicate(Name)
                        newFilledRegionType.ForegroundPatternColor = RColor(c[1].R, c[1].G, c[1].B)
                        newFilledRegionType.ForegroundPatternId = p[1]
                        newFilledRegionType.LineWeight = int(w)
            self.Close()
            t.Commit()

class Presenter(object):
    def __init__(self):
        self._window = Window()

    def run(self):
        self._window.ShowDialog()

if __name__ == '__main__':
    p = Presenter()
    p.run()