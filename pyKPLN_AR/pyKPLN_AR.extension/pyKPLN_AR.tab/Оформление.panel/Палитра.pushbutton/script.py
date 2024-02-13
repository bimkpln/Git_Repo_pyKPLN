# -*- coding: utf-8 -*-
"""
ColorPalette

"""
__author__ = 'Gyulnara Fedoseeva'
__title__ = "Палитра"
__doc__ = 'Создание палитры с сохранением в текущий проект'\

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
from pyrevit import revit as REVIT
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from Autodesk.Revit.DB import BuiltInParameter,\
    FilteredElementCollector, BuiltInCategory, Transaction,\
    Category, ElementLevelFilter, XYZ, Line, ElementTransformUtils, LinePatternElement, GraphicsStyleType
from Autodesk.Revit.DB import Color as RColor
from Autodesk.Revit.DB.Structure import StructuralType
from libKPLN import color_palette

#Параметры для логирования в Extensible Storage. Не менять по ходу изменений кода
extStorage_guid = "720080C5-DA99-40D7-9445-E53F288AA180"
extStorage_name = "kpln_ar_color_palette"

if __shiftclick__:
    try:
        obj = color_palette.create_obj(extStorage_guid, extStorage_name)
        text = color_palette.read_log(obj)
        print(text.split('|')[0])
    except:
        print("Записи отсутствуют")
    script.exit()

try:
    extStorage = color_palette.create_obj(extStorage_guid, extStorage_name)
    extText = color_palette.read_log(extStorage)
    extColors = extText.split('|')[1:]
except:
    extStorage = None

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
        self.Btn.Size = Size(120, 50)
        self.Btn.Text = self.name
        self.Btn.BackColor = self.color
        self.Btn.Click += click_event
        self.Btn.DoubleClick += dblclick_event

class Input(Form):
    UserInput = ''

    def __init__(self):
        self.Size = Size(400, 200)
        self.Text = 'KPLN Создание палитры'
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
    PaletteButtons = []

    def __init__(self):
        self.Size = Size(730, 500)
        self.Text = 'KPLN Создание палитры'
        self.Icon = Icon('X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico')
        self.CenterToScreen()

        self._tooltip = ToolTip()

        self.Buttons = [] #Список кнопок для удаления
        self.Colors = [] #Список цветов [Name, Color]

        self._btn_help = Button()
        self._btn_help.Size = Size(50, 50)
        self._btn_help.Location = Point(50, 20)
        self._btn_help.Text = '?'
        self._btn_help.Font = Font("Arial", 14, FontStyle.Bold)
        self.Controls.Add(self._btn_help)
        self._tooltip.SetToolTip(self._btn_help, "Открыть инструкцию к плагину")
        self._btn_help.Click += self._help

        self.info = Label()
        self.info.Text = 'Плагин позволяет задать цветовую палитру RGB и зафиксировать ее в системных настройках текущего проекта с возможностью дальнейшего редактирования'
        self.info.Font = Font('Century Gothic', 10)
        self.info.Size = Size(400, 100)
        self.info.Location = Point(110, 20)
        self.info.Parent = self

        # COLORS
        self._layout_1 = FlowLayoutPanel()
        self._layout_1.AutoScroll = True
        self._layout_1.Location = Point(40, 150)
        self._layout_1.BackColor = Color.FromName('White')
        self._layout_1.Size = Size(630, 240)
        self.Controls.Add(self._layout_1)

        self._lbl_1 = Label(Text = 'ТЕКУЩАЯ ПАЛИТРА:')
        self._lbl_1.Size = Size(170, 20)
        self._lbl_1.Location = Point(45, 130)
        self._lbl_1.Parent = self

        self._btn_1_add = Button()
        self._btn_1_add.Size = Size(80, 25)
        self._btn_1_add.Location = Point(40, 400)
        self._btn_1_add.Text = 'Добавить'
        self.Controls.Add(self._btn_1_add)
        self._btn_1_add.Click += self._on_button_color_add_click

        self._btn_1_del = Button()
        self._btn_1_del.Size = Size(80, 25)
        self._btn_1_del.Location = Point(130, 400)
        self._btn_1_del.Text = 'Удалить'
        self.Controls.Add(self._btn_1_del)
        self._btn_1_del.Click += self._on_button_color_del_click

        # ADD COLORS FROM PALETTE
        if extStorage != None:
            for i in extColors:
                color = i.split('*')[1]
                name = i.split('*')[0] + '\n' + color
                R = int(color.split('-')[0])
                G = int(color.split('-')[1])
                B = int(color.split('-')[2])
                self._btn = ColorButton(DblButton(), name, Color.FromArgb(R,G,B))
                self._btn.initBtn(self._get_selected_button, self._on_button_color_click)
                self._layout_1.Controls.Add(self._btn.Btn)

        # CREATE PALETTE
        self._btn_run = Button()
        self._btn_run.Size = Size(150, 40)
        self._btn_run.Location = Point(510, 20)
        if extStorage == None:
            self._btn_run.Text = 'Создать палитру'
        else:
            self._btn_run.Text = 'Перезаписать палитру'
        self.Controls.Add(self._btn_run)
        self._btn_run.Click += self.createPalette

    def _help(self, sender, event_args):
        webbrowser.open('http://moodle/mod/book/view.php?id=502&chapterid=1255')

    def _on_button_color_click(self, sender, event_args):
        self.Buttons = []
        self._window = ColorDialog()
        self._window.AllowFullOpen = True
        self._window.ShowHelp = True
        self._window.ShowDialog()
        sender.BackColor = self._window.Color
        colorNameWind = Input()
        colorNameWind.ShowDialog()
        sender.Text = colorNameWind.UserInput + ' \n' + self._window.Color.R.ToString() + " - " + self._window.Color.G.ToString() + " - " + self._window.Color.B.ToString()

    def _on_button_color_add_click(self, sender, event_args):
        self._btn = ColorButton(DblButton(), 'Выберите цвет', Color.FromName('White'))
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

    def createPalette(self, sender, event_args):
        for color in self._layout_1.Controls:
            if color.Text != 'Выберите цвет':
                colorName = str(color.Text).split(' ')
                self.Colors.append([colorName[0], color.BackColor])

        with Transaction(doc, 'Создать палитру') as t:
            t.Start()
            extString = ''
            for c in self.Colors:
                name = c[0]
                colr = c[1].R.ToString() + '-' + c[1].G.ToString() + '-' + c[1].B.ToString()
                text = '|' + name + '*' + colr
                extString += text
            self.Close()
            if extStorage == None:
                try:
                    obj = color_palette.create_obj(extStorage_guid, extStorage_name)
                    color_palette.write_log(obj, extString)
                except:
                    print("Ошибка записи. Обратитесь в BIM - отдел!")
            t.Commit()

class Presenter(object):
    def __init__(self):
        self._window = Window()

    def run(self):
        self._window.ShowDialog()

if __name__ == '__main__':
    p = Presenter()
    p.run()