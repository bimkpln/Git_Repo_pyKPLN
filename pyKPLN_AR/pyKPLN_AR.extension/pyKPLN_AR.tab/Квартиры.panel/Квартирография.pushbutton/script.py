# -*- coding: utf-8 -*-
"""
Квартирография

"""
__helpurl__ = "http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=681/"
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Квартирография"
__doc__ = 'KPLN Квартирография разработана для автоматической нумерации помещений и определения показетелей площади. См. инструкцию по использованию скрипта внутри (значок «?»)\n' \
          'ВАЖНО: Для корректной работы файл должен быть настроен для совместной работы\n' \
          'Квартиры определяются по параметру «Назначение» (см. вкладка «Параметры» скрипта) и значение у всек квартир должно быть равно «Квартира»;\n\n' \
          'Функции:\n' \
          '   - Сохранение результатов работы скрипта и настроек пользователя - для корректной работы не рекомендуется удалять файл (...\\путь к хранилищу\\\\data\\\\rooms_settings.txt)\n\n' \
          '   - Запись значений ТЭП (сведения о проекте):\n' \
		  '      ТЭП_Площадь_Жилая\n' \
		  '      ТЭП_Площадь_Общая\n' \
		  '      ТЭП_Площадь_Общая_К\n' \
		  '      ТЭП_Площадь_МОП\n' \
		  '      ТЭП_Площадь_Технические помещения\n' \
		  '      ТЭП_Площадь_Технические пространства\n' \
		  '      ТЭП_Площадь_Аренда\n' \
		  '      ТЭП_Площадь_Кладовые\n' \
		  '      ТЭП_Площадь_ДОУ\n' \
		  '      ТЭП_Количество_Квартиры\n' \
		  '      ТЭП_Количество_Кладовые\n' \
		  '      С созданием текстовой таблицы для импорта в Excell\n\n' \
          '   - Проверка на ошибки и уведомление пользователя (см. log-файл)\n\n' \
          '   - ЛОГ - автоматическая генерация в папку проекта (...путь к файлу хранилищу\\\\data\\\\log.txt)' \
		  'версия: 2019/08/24' \

"""
KPLN

"""
import clr
clr.AddReference('RevitAPI')
import math
import webbrowser
import re
import os
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import revit as REVIT
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection

from System.Collections.Generic import *
import System
from System.Windows.Forms import *
from System.Drawing import *
import re
from itertools import chain
import datetime
from rpw.ui.forms import TextInput, Alert, select_folder

class CreateWindow(Form):
    def __init__(self): 
        #INIT
        self.Name = "KPLN_AR_Квартирография"
        self.Text = "KPLN Квартирография"
        self.Size = Size(800, 600)
        self.MinimumSize = Size(800, 600)
        self.MaximumSize = Size(800, 600)
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ControlBox = True
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.TopMost = True
        self.Icon = Icon("X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
        self.initialisation_completed = False #BLOCKS METHODS UNTIL INITIALISATION IS DONE
        self.alphabet = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
        self.alphabet.sort()
        self.CBImage = Image.FromFile('X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\CB000001.png')
        self.out = script.get_output()

        #TEP
        self.tep = ""
        self.tep_flats = 0.00
        self.tep_flats_k = 0.00
        self.tep_flats_living = 0.00
        self.tep_flats_balcony = 0.00
        self.tep_flats_terrass = 0.00
        self.tep_flats_wet = 0.00
        self.tep_mop = 0.00
        self.tep_com = 0.00
        self.tep_tech = 0.00
        self.tep_tech_space = 0.00
        self.tep_all = 0.00
        self.tep_all_k = 0.00
        self.tep_flat_amount = 0
        self.tep_store_amount = 0

        #BUILT-IN-PARAMETERS
        self.builtin_room_name_par = DB.BuiltInParameter.ROOM_NAME
        self.builtin_room_department_par = DB.BuiltInParameter.ROOM_DEPARTMENT

        #DEBUG&HISTORY
        self.errors_amount_former = 0
        self.errors_amount = 0
        self.commonarea = 0.00
        self.commonarea_former = 0.00	
        
        #ROOMDATA
        self.result = ""
        self.department_dict = ["Квартира", "МОП"]
        self.abs_numeration = []
        self.rooms_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
        self.rooms_names = []
        self.levels_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
        self.levels_id = []
        self.levels_height = []
        self.room_parameters_names = []
        self.room_parameters_names_area = []
        self.room_parameters_names_string = []
        self.allrooms = []
        for room in self.rooms_collector:
            for j in room.Parameters:
                if str(j.StorageType) == "String":
                    self.room_parameters_names.append("Помещение: {}".format(j.Definition.Name))
                    self.room_parameters_names_string.append(j.Definition.Name)
                elif str(j.StorageType) == "Double":
                    try:
                        if str(j.DisplayUnitType) == "DUT_SQUARE_METERS":
                            self.room_parameters_names_area.append(j.Definition.Name)
                    except:
                        if str(j.GetUnitTypeId().TypeId) == "autodesk.unit.unit:squareMeters-1.0.1":
                            self.room_parameters_names_area.append(j.Definition.Name)

            break

        for room in self.rooms_collector:
            try:
                if room.Area > 0:
                    self.allrooms.append(room)
                    if not self.in_list("{}+{}".format(room.get_Parameter(self.builtin_room_department_par).AsString(), room.get_Parameter(self.builtin_room_name_par).AsString()), self.rooms_names):
                        self.rooms_names.append("{}+{}".format(room.get_Parameter(self.builtin_room_department_par).AsString(), room.get_Parameter(self.builtin_room_name_par).AsString()))
            except :
                pass
        for level in self.levels_collector:
            try:
                for j in level.Parameters:
                    if str(j.StorageType) == "String":
                        self.room_parameters_names.append("Связанный уровень: {}".format(j.Definition.Name))
                break
            except :
                pass
        #ELEMENTS UI
        self.tooltip = ToolTip() #ToolTips
        self.cb = [] #ComboBoxes
        self.chbx_settings = [] #ChekBoxes
        self.cb_parameters = [] #ComboBoxes
        self.lb = [] #ListBoxes
        self.lbl = [] #Label
        self.lbl_names = [] #Label
        self.lbl_department = [] #Label
        self.lbl_parameters = [] #Label
        self.tb = [] #TextBoxes
        self.gb = [] #GroupBoxes
        self.bt = [] #Buttons
        self.tp = [] #TabPages
        self.pnl = [] #Panel
        self.tc = TabControl()
        self.status = ["не запускался", "user", "неизвестно"]

        #DRAG&DROP
        self.formula_parts = []
        self.formula = Label(Text = "[Drag and Drop]")
        self.formula.DragDrop += self.OnDragDrop
        self.formula.DragEnter += self.OnDragEnter
        self.formula.AllowDrop = True
        self.font = Font("Arial", 12, FontStyle.Bold)
        self.formula.Font = self.font
        self.formula.MouseClick += self.Clear_formula

        self.formula_parts_o = []
        self.formula_o = Label(Text = "[Drag and Drop]")
        self.formula_o.DragDrop += self.OnDragDrop
        self.formula_o.DragEnter += self.OnDragEnter
        self.formula_o.AllowDrop = True
        self.font = Font("Arial", 12, FontStyle.Bold)
        self.formula_o.Font = self.font
        self.formula_o.MouseClick += self.Clear_formula

        #DICTIONARIES
        self.default_parameters = ["Помещение: Назначение",
                                "Помещение: Корпус (номер)",
                                "Помещение: Секция (номер)",
                                "Помещение: Этаж (номер)",
                                "Квартира: Номер (на этаже)",
                                "Помещение: Внутриквартирный номер",
                                "Помещение: Номер",
                                "Помещение: S факт.",
                                "Помещение: S коэфф.",
                                "Квартира: Наименование",
                                "Квартира: Номер (сквозной)",
                                "Квартира: Код",
                                "Квартира: S факт.",
                                "Квартира: S жилая",
                                "Квартира: S лоджий",
                                "Квартира: S нежилых помещений",
                                "Квартира: S коэфф.",
                                "Квартира: S без лоджий"]
        #TOOLTIP
        self.default_parameters_tt = ["Для квартир должно быть установлено\nзначение «Кв». Параметр отвечает\nза внутриквартирную нумерацию\nи нумерацию внутри групп остальных\nпомещений по назначению",
                                "Если в проекте нет деления на корпуса\nможно не заполнять значения\nлибо установить одно значение\nдля всех помещений",
                                "Исходный параметр: Должно быть заполнено числовым значением",
                                "Исходный параметр: Должно быть заполнено числовым значением. Рекомендуется взять из параметра связанного уровня",
                                "Номер квартиры на этаже (по часовой стрелке\nот 1 и далее по прядку)\nВажно: значение «0» для неквартирных\nпомещений",
                                "Помещение: Внутриквартирный номер",
                                "В данный параметр будет записан номер\nпомещения по формуле (см. «Главная»)",
                                "Помещение: S факт.",
                                "Помещение: S коэфф.",
                                "напр. «Двухкомнатная квартира»",
                                "Абсолютный номер квартиры по порядку от «1»",
                                "напр. «К2» для двухкомнатной квартиры",
                                "Фактическая общая площадь квартиры\n(без понижающих коэффициентов)",
                                "Общая площадь жилых помещений квартиры",
                                "Площади балконов, террасс и лоджий\n(с понижающим коэффициентом)",
                                "Общая площадь нежилых помещений квартиры",
                                "Общая площадь квартиры\n(с понижающими коэффициентами)",
                                "Общая площадь квартиры без учета летних помещений"]

        #DEFAULT VALUES
        self.default_parameters_values = ["Помещение: Назначение",
                                "Помещение: ПОМ_Корпус",
                                "Помещение: ПОМ_Секция",
                                "Помещение: ПОМ_Номер этажа",
                                "Помещение: ПОМ_Номер квартиры",
                                "ПОМ_Номер помещения",
                                "Номер",
                                "ПОМ_Площадь",
                                "ПОМ_Площадь_К",
                                "КВ_Наименование",
                                "КВ_Номер",
                                "КВ_Код",
                                "КВ_Площадь_Общая",
                                "КВ_Площадь_Жилая",
                                "КВ_Площадь_Летние",
                                "КВ_Площадь_Нежилая",
                                "КВ_Площадь_Общая_К",
                                "КВ_Площадь_Отапливаемые"]

        self.default_project_parameters = ["ТЭП_Площадь_Жилая",
                                        "ТЭП_Площадь_Общая",
                                        "ТЭП_Площадь_Общая_К",
                                        "ТЭП_Площадь_МОП",
                                        "ТЭП_Площадь_Технические помещения",
                                        "ТЭП_Площадь_Технические пространства",
                                        "ТЭП_Площадь_Аренда",
                                        "ТЭП_Площадь_Кладовые",
                                        "ТЭП_Количество_Квартиры",
                                        "ТЭП_Количество_Кладовые",
                                        "ТЭП_Площадь_ДОУ"]

        self.default_project_values = []	

        self.default_categories = ["Кв: Жилое пом.",
                                "Кв: Лоджия",
                                "Кв: Терраса",
                                "Кв: Балкон",
                                "Кв: Нежилое пом.",
                                "Кв: Нежилое пом. (мокрое)",
                                "МОП",
                                "Тех.помещения",
                                "Тех.пространства",
                                "Аренда",
                                "Кладовые",
                                "ДОУ",
                                "Прочее"]

        self.default_settings = {"Округлять площадь до «0.1»" : True,
                                "Записать ТЭП" : False,
                                "Показывать log" : True,
                                "Расчитывать квартиры" : True,
                                "Расчитывать неквартирные помещения" : True,
                                "Автоматический разделитель в формуле" : True}

        self.default_settings_names = ["Округлять площадь до «0.1»", "Записать ТЭП", "Показывать log",
                                    "Расчитывать квартиры", "Расчитывать неквартирные помещения",
                                    "Автоматический разделитель в формуле"]

        #CM
        self.contextmenu = ContextMenuStrip()
        self.cm_target = ""
        self.contextmenu.ItemClicked += self.item_clicked
        for i in self.default_categories:
            self.contextmenu.Items.Add(i)
            if i == "Кв: Нежилое пом. (мокрое)":
                self.contextmenu.Items.Add(ToolStripSeparator())

        #TABCONTROL
        self.tp.append(TabPage())
        self.tp.append(TabPage())
        self.tc.Parent = self
        self.tc.Size = Size(786, 563)
        self.tp[0].Text = "Главная"
        self.tp[1].Text = "Параметры"
        self.tc.TabPages.Add(self.tp[0])
        self.tc.TabPages.Add(self.tp[1])

        #TABPAGE_1
        self.lbl.append(Label(Text = "Дата последнего расчета: {} (user: {})".format(self.status[0], self.status[1])))
        self.lbl[0].Parent = self.tp[0]
        self.lbl[0].Location = Point(5,10)
        self.lbl[0].Size = Size(500, 20)
        self.lbl[0].ForeColor = Color.FromArgb(0,150,150,150)

        self.lbl.append(Label(Text = "Статус: {}".format(self.status[2])))
        self.lbl[1].Parent = self.tp[0]
        self.lbl[1].Location = Point(5,35)
        self.lbl[1].Size = Size(500, 20)

        self.lbl.append(Label())
        self.lbl[2].Parent = self.tp[0]
        self.lbl[2].Location = Point(5,60)
        self.lbl[2].Size = Size(765, 2)
        self.lbl[2].BorderStyle = BorderStyle.Fixed3D

        self.bt.append(Button(Text = "?"))
        self.bt[0].Parent = self.tp[0]
        self.bt[0].Location = Point(730,10)
        self.bt[0].Size = Size(40, 40)
        self.bt[0].BackColor = Color.FromArgb(255,218,185)
        self.bt[0].Font = Font("Arial", 14, FontStyle.Bold)
        self.bt[0].ForeColor = Color.FromArgb(250,0,0)
        self.bt[0].Click += self.go_to_help
        self.tooltip.SetToolTip(self.bt[0], "Открыть инструкцию к плагину")

        self.bt.append(Button(Text = "Запуск"))
        self.bt[1].Parent = self.tp[0]
        self.bt[1].Location = Point(580,10)
        self.bt[1].Size = Size(140, 40)
        self.bt[1].Click += self.run

        self.gb.append(GroupBox())
        self.gb[0].Text = "Формула номера помещения (квартиры):"
        self.gb[0].Size = Size(370, 70)
        self.gb[0].Location = Point(5, 75)
        self.gb[0].Parent = self.tp[0]

        self.gb.append(GroupBox())
        self.gb[1].Text = "Формула номера помещения (остальные помещения):"
        self.gb[1].Size = Size(370, 70)
        self.gb[1].Location = Point(5, 155)
        self.gb[1].Parent = self.tp[0]

        self.formula.Parent = self.gb[0]
        self.formula.Location = Point(15, 15)
        self.formula.ForeColor = Color.FromArgb(0,0,0,255)
        self.formula.TextAlign = ContentAlignment.MiddleCenter
        self.formula.Size = Size(340,40)

        self.formula_o.Parent = self.gb[1]
        self.formula_o.Location = Point(15, 15)
        self.formula_o.ForeColor = Color.FromArgb(0,0,0,255)
        self.formula_o.TextAlign = ContentAlignment.MiddleCenter
        self.formula_o.Size = Size(340,40)

        self.gb.append(GroupBox())
        self.gb[2].Text = "Доступные параметры:"
        self.gb[2].Size = Size(370, 297)
        self.gb[2].Location = Point(5, 235)
        self.gb[2].Parent = self.tp[0]

        self.gb.append(GroupBox())
        self.gb[3].Text = "Категории помещений:"
        self.gb[3].Size = Size(370, 457)
        self.gb[3].Location = Point(400, 75)
        self.gb[3].Parent = self.tp[0]

        self.pnl.append(Panel())
        self.pnl[0].Size = Size(360, 430)
        self.pnl[0].Location = Point(3, 20)
        self.pnl[0].Parent = self.gb[3]
        self.pnl[0].AutoScroll = True

        self.lb.append(ListBox())
        self.lb[0].Parent = self.gb[2]
        self.lb[0].AllowDrop = True
        self.lb[0].Size = Size(350, 270)
        self.lb[0].Location = Point(10, 20)
        self.room_parameters_names.sort()
        self.lb[0].Items.Add("« . »")
        self.lb[0].Items.Add("« _ »")
        self.lb[0].Items.Add("«   »")
        self.lb[0].Items.Add("« - »")
        self.lb[0].Items.Add("« , »")
        self.lb[0].Items.Add("« = »")
        self.lb[0].Items.Add("« + »")
        self.lb[0].Items.Add("« ( »")
        self.lb[0].Items.Add("« ) »")
        self.lb[0].Items.Add("« [ »")
        self.lb[0].Items.Add("« ] »")
        self.lb[0].Items.Add("Нумерация: номер по назначению")
        self.lb[0].Items.Add("Нумерация: номер внутри назначения")
        self.lb[0].Items.Add("Системный: id помещения")
        self.lb[0].Items.Add("Системный: рабочий набор")
        for parameter in self.room_parameters_names:
            self.lb[0].Items.Add(parameter)
        self.lb[0].MouseDown += self.OnMousDown

        self.rooms_names.sort()
        for i in range(0, len(self.rooms_names)):
            self.cb.append(ComboBox())
            self.cb[i].Parent = self.pnl[0]
            self.cb[i].Size = Size(180, 10)
            self.cb[i].Location = Point(150, 50 * i + 20)
            self.cb[i].DropDownStyle = ComboBoxStyle.DropDownList
            self.cb[i].SelectedIndexChanged += self.CblOnChanged
            self.cb[i].FlatStyle = FlatStyle.Flat
            self.temp_list = self.rooms_names[i].split("+")
            self.temp_text = "{}: {}".format(self.temp_list[0], self.temp_list[1])
            for t in self.default_categories:
                self.cb[i].Items.Add(t)
            if self.temp_list[1].lower() in "лоджиялоджии":
                self.cb[i].Text = self.default_categories[1]
            elif self.temp_list[1].lower() in "ванная комнатапостирочная комнатадушдушеваясанузелтуалетс/ууборная":
                self.cb[i].Text = self.default_categories[5]
            elif self.temp_list[1].lower() in "прихожаягардеробнаякухня-нишакухня-столоваякладовая":
                self.cb[i].Text = self.default_categories[4]
            elif self.temp_list[1].lower() in "моплклестничная клеткалифтовой холллифтовый холлтамбурвходной тамбурколясочнаякоридорлк-0лк-1лк-2лк-3лк-4лк-5лк-6лк-7лк-8лк-9лк-10кл-11лк-12вестибюль":
                self.cb[i].Text = self.default_categories[6]
            elif self.temp_list[1].lower() in "террасатерасатеррасса":
                self.cb[i].Text = self.default_categories[2]
            elif self.temp_list[1].lower() in "балконы":
                self.cb[i].Text = self.default_categories[3]
            elif self.temp_list[1].lower() in "гостиннаягостинаякухня-гостинаяжилая комнатаспальняспальная комната":
                self.cb[i].Text = self.default_categories[0]
            elif self.temp_list[1].lower() in "пуиитпкроссоваяврупомещение по сбору мусора":
                self.cb[i].Text = self.default_categories[7]
            else:
                self.cb[i].Text = self.default_categories[9]
                
        for i in range(0, len(self.rooms_names)):
            self.temp_list = self.rooms_names[i].split("+")
            self.temp_text = "{}: {}".format(self.temp_list[0], self.temp_list[1])
            self.lbl_names.append(Label(Text = self.temp_list[1]))
            self.lbl_names[i].Parent = self.pnl[0]
            self.lbl_names[i].Size = Size(145, 50)
            self.lbl_names[i].Location = Point(0, 50 * i + 20)
            self.lbl_names[i].ForeColor = Color.FromArgb(0,0,0,255)
            self.lbl_names[i].TextAlign = ContentAlignment.TopRight
            self.lbl_department.append(Label(Text = self.temp_list[0]))
            self.lbl_department[i].Parent = self.pnl[0]
            self.lbl_department[i].Size = Size(180, 15)
            self.lbl_department[i].Location = Point(150, 50 * i - 0)
            self.lbl_department[i].ForeColor = Color.FromArgb(0,255,0,0)
            self.lbl_department[i].TextAlign = ContentAlignment.TopLeft
            self.lbl_department[i].MouseEnter += self.mouse_enter
            self.lbl_department[i].MouseLeave += self.mouse_exit
            self.lbl_department[i].MouseClick += self.mouse_click_r

        #TABPAGE_2
        self.lbl.append(Label(Text = "Дата последнего расчета: {} (user: {})".format(self.status[0], self.status[1])))
        self.lbl[3].Parent = self.tp[1]
        self.lbl[3].Location = Point(5,10)
        self.lbl[3].Size = Size(500, 20)
        self.lbl[3].ForeColor = Color.FromArgb(0,150,150,150)

        self.lbl.append(Label())
        self.lbl[4].Parent = self.tp[1]
        self.lbl[4].Location = Point(5,60)
        self.lbl[4].Size = Size(765, 2)
        self.lbl[4].BorderStyle = BorderStyle.Fixed3D

        self.bt.append(Button(Text = "?"))
        self.bt[2].Parent = self.tp[1]
        self.bt[2].Location = Point(730,10)
        self.bt[2].Size = Size(40, 40)
        self.bt[2].BackColor = Color.FromArgb(255,218,185)
        self.bt[2].Font = Font("Arial", 14, FontStyle.Bold)
        self.bt[2].ForeColor = Color.FromArgb(250,0,0)
        self.bt[2].Click += self.go_to_help
        self.tooltip.SetToolTip(self.bt[2], "Открыть инструкцию к плагину")

        self.bt.append(Button(Text = "Сохранить"))
        self.bt[3].Parent = self.tp[1]
        self.bt[3].Location = Point(580,10)
        self.bt[3].Size = Size(140, 40)
        self.bt[3].Enabled = False
        self.bt[3].Click += self.save_settings

        self.gb.append(GroupBox())
        self.gb[4].Text = "Настройки:"
        self.gb[4].Size = Size(370, 457)
        self.gb[4].Location = Point(5, 75)
        self.gb[4].Parent = self.tp[1]

        self.gb.append(GroupBox())
        self.gb[5].Text = "Рабочие параметры:"
        self.gb[5].Size = Size(370, 457)
        self.gb[5].Location = Point(400, 75)
        self.gb[5].Parent = self.tp[1]

        for i in range(0, len(self.default_settings.keys())):
            self.chbx_settings.append(CheckBox())
            self.chbx_settings[i].Text = self.default_settings_names[i]
            self.chbx_settings[i].Parent = self.gb[4]
            self.chbx_settings[i].Size = Size(300, 50)
            self.chbx_settings[i].Location = Point(25, 50 * i + 21)
            self.chbx_settings[i].FlatStyle = FlatStyle.Flat
            self.chbx_settings[i].CheckedChanged += self.activate_button
            self.chbx_settings[i].Checked = self.default_settings[self.default_settings_names[i]]

        self.pnl.append(Panel())
        self.pnl[1].Size = Size(360, 430)
        self.pnl[1].Location = Point(3, 20)
        self.pnl[1].Parent = self.gb[5]
        self.pnl[1].AutoScroll = True

        for i in range(0, len(self.default_parameters)):
            self.cb_parameters.append(ComboBox())
            self.cb_parameters[i].DropDownWidth = 400
            self.cb_parameters[i].Parent = self.pnl[1]
            self.cb_parameters[i].Size = Size(180, 10)
            self.cb_parameters[i].Location = Point(150, 50 * i + 5)
            self.cb_parameters[i].DropDownStyle = ComboBoxStyle.DropDownList
            self.cb_parameters[i].SelectedIndexChanged += self.CblOnChanged
            self.cb_parameters[i].Items.Add("<Выберите параметр>")
            self.cb_parameters[i].Text = "<Выберите параметр>"
            self.room_parameters_names_area.sort()
            self.room_parameters_names_string.sort()
            if i < 5:
                for t in self.room_parameters_names:
                    self.cb_parameters[i].Items.Add(t)
            else:
                self.cb_parameters[i].FlatStyle = FlatStyle.Flat
                if "S" in self.default_parameters[i]:
                    for t in self.room_parameters_names_area:
                        self.cb_parameters[i].Items.Add(t)
                else:
                    for t in self.room_parameters_names_string:
                        self.cb_parameters[i].Items.Add(t)
            try:
                self.cb_parameters[i].Text = self.default_parameters_values[i]
            except:
                pass
        self.cb_parameters[0].Items.Clear()
        self.cb_parameters[0].Items.Add("<Системный: Назначение>")
        self.cb_parameters[0].Text = "<Системный: Назначение>"
        self.cb_parameters[0].Enabled = False #ВРЕМЕННАЯ бЛОКИРОВКА ПАРАМЕТРА

        for i in range(0, len(self.default_parameters)):
            self.lbl_parameters.append(Label(Text = self.default_parameters[i]))
            self.lbl_parameters[i].Parent = self.pnl[1]
            self.lbl_parameters[i].Size = Size(145, 50)
            self.lbl_parameters[i].Location = Point(0, 50 * i + 5)
            self.tooltip.SetToolTip(self.lbl_parameters[i], self.default_parameters_tt[i])
            if i < 5:
                self.lbl_parameters[i].ForeColor = Color.FromArgb(0,255,0,0)		
            else:
                self.lbl_parameters[i].ForeColor = Color.FromArgb(0,0,0,0)
            self.lbl_parameters[i].TextAlign = ContentAlignment.TopRight

        #LOAD DEFAULT SETTINGS
        self.load_settings()
        self.CenterToScreen()
        self.initialisation_completed = True

    def CblOnChanged(self, sender, event):
        if self.initialisation_completed:
            self.bt[3].Enabled = True

    def Clear_formula(self, sender, args):
        self.bt[3].Enabled = True
        if sender == self.formula:
            self.font = Font("Arial", 12, FontStyle.Bold)
            self.formula_parts = []
            self.formula.Font = self.font
            self.formula.Text = "[Drag and Drop]"
        elif sender == self.formula_o:
            self.font = Font("Arial", 12, FontStyle.Bold)
            self.formula_parts_o = []
            self.formula_o.Font = self.font
            self.formula_o.Text = "[Drag and Drop]"

    def OnMousDown(self, sender, event):
        sender.DoDragDrop(sender.Text, DragDropEffects.Copy)

    def OnDragEnter(self, sender, event):
        event.Effect = DragDropEffects.Copy

    def item_clicked(self, sender, event):
        for i in range(0, len(self.lbl_department)):
            if self.lbl_department[i].Text == self.cm_target:
                self.cb[i].Text = event.ClickedItem.Text

    def mouse_click_r(self, sender, event):
        self.cm_target = sender.Text
        self.contextmenu.Show(sender, Point(event.X, event.Y))

    def mouse_enter(self, sender, event):
        for i in range(0, len(self.lbl_department)):
            if self.lbl_department[i].Text == sender.Text:
                self.lbl_department[i].ForeColor = Color.FromArgb(0,0,0,255)
                self.lbl_names[i].BackgroundImage = self.CBImage

    def mouse_exit(self, sender, event):
        for i in range(0, len(self.lbl_department)):
            if self.lbl_department[i].Text == sender.Text:
                self.lbl_department[i].ForeColor = Color.FromArgb(0,255,0,0)
                self.lbl_names[i].BackgroundImage = None

    def OnDragDrop(self, sender, event):
        if sender == self.formula:
            self.bt[3].Enabled = True
            sender.ForeColor = Color.FromArgb(0,0,0,255)
            sender_text = event.Data.GetData(DataFormats.Text)
            if sender_text.startswith("Помещение: "): text = "[Пом] {}".format(sender_text[11:])
            elif sender_text.startswith("Связанный уровень: "): text = "[Ур] {}".format(sender_text[19:])
            else: text = sender_text
            if len(self.formula_parts) == 0:
                self.font = Font("Arial", 12, FontStyle.Bold)
                self.formula_parts.append(text)
                self.formula.Font = self.font
            elif len(self.formula_parts) == 1:
                self.font = Font("Arial", 10, FontStyle.Bold)
                self.formula_parts.append(text)
                self.formula.Font = self.font
            elif len(self.formula_parts) == 2:
                self.font = Font("Arial", 9, FontStyle.Bold)
                self.formula_parts.append(text)
                self.formula.Font = self.font
            elif len(self.formula_parts) == 3:
                self.font = Font("Arial", 8, FontStyle.Bold)
                self.formula_parts.append(text)
                self.formula.Font = self.font
            elif len(self.formula_parts) > 3 and len(self.formula_parts) < 10:
                self.font = Font("Arial", 6, FontStyle.Bold)
                self.formula_parts.append(text)
            else:
                self.font = Font("Arial", 12, FontStyle.Bold)
                self.formula_parts = []
                self.formula.Font = self.font
            if len(self.formula_parts) == 0:
                self.formula.Text = "[Drag and Drop]"
            else:
                self.formula.Text = ""
                for i in range(0, len(self.formula_parts)):
                    if i == 0: self.formula.Text += self.formula_parts[i]
                    else: self.formula.Text += " . {}".format(self.formula_parts[i])
        elif sender == self.formula_o:
            self.bt[3].Enabled = True
            sender.ForeColor = Color.FromArgb(0,0,0,255)
            sender_text = event.Data.GetData(DataFormats.Text)
            if sender_text.startswith("Помещение: "): text = "[Пом] {}".format(sender_text[11:])
            elif sender_text.startswith("Связанный уровень: "): text = "[Ур] {}".format(sender_text[19:])
            else: text = sender_text
            if len(self.formula_parts_o) == 0:
                self.font = Font("Arial", 12, FontStyle.Bold)
                self.formula_parts_o.append(text)
                self.formula_o.Font = self.font
            elif len(self.formula_parts_o) == 1:
                self.font = Font("Arial", 10, FontStyle.Bold)
                self.formula_parts_o.append(text)
                self.formula_o.Font = self.font
            elif len(self.formula_parts_o) == 2:
                self.font = Font("Arial", 9, FontStyle.Bold)
                self.formula_parts_o.append(text)
                self.formula_o.Font = self.font
            elif len(self.formula_parts_o) == 3:
                self.font = Font("Arial", 8, FontStyle.Bold)
                self.formula_parts_o.append(text)
                self.formula_o.Font = self.font
            elif len(self.formula_parts_o) > 3 and len(self.formula_parts_o) < 10:
                self.font = Font("Arial", 6, FontStyle.Bold)
                self.formula_parts_o.append(text)
            else:
                self.font = Font("Arial", 12, FontStyle.Bold)
                self.formula_parts_o = []
                self.formula_o.Font = self.font
            if len(self.formula_parts_o) == 0:
                self.formula_o.Text = "[Drag and Drop]"
            else:
                self.formula_o.Text = ""
                for i in range(0, len(self.formula_parts_o)):
                    if i == 0: self.formula_o.Text += self.formula_parts_o[i]
                    else: self.formula_o.Text += " . {}".format(self.formula_parts_o[i])

    def activate_button(self, sender, event):
        self.bt[3].Enabled = True

    def update_meta(self, result = "Неизвестно", reset = False):
        try:
            if reset:
                now = datetime.datetime.now()
                self.status[0] = "{}-{}-{} {}:{}:{}".format(now.year, now.month, now.day, now.hour, now.minute, now.second)
                self.status[1] = revit.username
                self.status[2] = result
            self.lbl[0].Text = "Дата последнего расчета: {} (user: {})".format(self.status[0], self.status[1])
            self.lbl[1].Text = "Статус: {}".format(result)
            self.lbl[3].Text = "Дата последнего расчета: {} (user: {})".format(self.status[0], self.status[1])
        except: pass

    def load_settings(self):
        try:
            parts = REVIT.query.get_central_path(doc=revit.doc).split("\\")
            path = ""
            for i in range(0, len(parts)-1):
                path += "{}\\".format(parts[i])
            path += "data"
            filepath = "{}\\rooms_settings.txt".format(path)
            if os.path.exists(filepath):
                file = open(filepath, "r")
                data = file.read().decode('utf-8')			
                values = data.split(";")
                file.close()
                v = []
                for i in range (0,6):
                    if str(values[i]) == str(True):
                        v.append(True)
                        self.chbx_settings[i].Checked = True
                    else:
                        v.append(False)
                        self.chbx_settings[i].Checked = False
                try:
                    self.commonarea_former = float(values[30])
                except: 
                    self.commonarea_former = 0.00

                try:
                    self.errors_amount_former = int(values[31])
                except: 
                    self.errors_amount_former = 0
                self.status[0] = values[6]
                self.status[1] = values[7]
                self.result_former = values[9]
                self.status[2] = self.check_state(errors = self.errors_amount_former)
                self.default_settings.clear()
                self.default_settings = {"Округлять площадь до «0.1»" : v[0],
                                    "Записать ТЭП" : v[1],
                                    "Показывать log" : v[2],
                                    "Расчитывать квартиры" : v[3],
                                    "Расчитывать неквартирные помещения" : v[4],
                                    "Автоматический разделитель в формуле" : v[5]}
                for b in range(0, len(self.cb_parameters)):
                    if  values[b + 10]:
                        if values[b + 10] > "":
                            self.cb_parameters[b].Text = values[b + 10]
                room_types_input = values[28].split("::")
                room_types = []
                for type in room_types_input:
                    room_types.append(type.split("_-_"))
                for t in room_types:
                    for i in range(0, len(self.rooms_names)):
                        if t[0] == self.rooms_names[i]:
                            self.cb[i].Text = t[1]
                self.formula_parts_input = values[29].split("::")
                self.formula_parts_o_input = values[32].split("::")
                self.formula_parts = []
                for fp in self.formula_parts_input:
                    if fp != "":
                        self.formula_parts.append(fp)
                self.formula_parts_o = []
                for fpo in self.formula_parts_o_input:
                    if fpo != "":
                        self.formula_parts_o.append(fpo)
                #UPDATE FORMULA
                if len(self.formula_parts) == 2:
                    self.font = Font("Arial", 10, FontStyle.Bold)
                    self.formula.Font = self.font
                elif len(self.formula_parts) == 3:
                    self.font = Font("Arial", 9, FontStyle.Bold)
                    self.formula.Font = self.font
                elif len(self.formula_parts) == 4:
                    self.font = Font("Arial", 8, FontStyle.Bold)
                    self.formula.Font = self.font
                elif len(self.formula_parts) > 4 and len(self.formula_parts) < 11:
                    self.font = Font("Arial", 8, FontStyle.Bold)
                    self.formula.Font = self.font
                else:
                    self.font = Font("Arial", 12, FontStyle.Bold)
                    self.formula.Font = self.font

                if len(self.formula_parts) == 0:
                    self.formula.Text = "[Drag and Drop]"
                else:
                    self.formula.Text = ""

                for i in range(0, len(self.formula_parts)):
                    if self.formula.Text == "": self.formula.Text += self.formula_parts[i]
                    else: self.formula.Text += " . {}".format(self.formula_parts[i])

                if len(self.formula_parts_o) == 2:
                    self.font = Font("Arial", 10, FontStyle.Bold)
                    self.formula_o.Font = self.font
                elif len(self.formula_parts_o) == 3:
                    self.font = Font("Arial", 9, FontStyle.Bold)
                    self.formula_o.Font = self.font
                elif len(self.formula_parts_o) == 4:
                    self.font = Font("Arial", 8, FontStyle.Bold)
                    self.formula_o.Font = self.font
                elif len(self.formula_parts_o) > 4 and len(self.formula_parts_o) < 11:
                    self.font = Font("Arial", 8, FontStyle.Bold)
                    self.formula_o.Font = self.font
                else:
                    self.font = Font("Arial", 12, FontStyle.Bold)
                    self.formula_o.Font = self.font

                if len(self.formula_parts_o) == 0:
                    self.formula_o.Text = "[Drag and Drop]"
                else:
                    self.formula_o.Text = ""

                for i in range(0, len(self.formula_parts_o)):
                    if self.formula_o.Text == "": self.formula_o.Text += self.formula_parts_o[i]
                    else: self.formula_o.Text += " . {}".format(self.formula_parts_o[i])
            self.update_meta(self.status[2], reset = False)
            self.bt[3].Enabled = False
        except:
            self.show_alert(text = "Во время загрузки предыдущих настроек произошла ошибка!\n\nПримечание: необходимо заново настроить текущие параметры либо сообщить о неполадке BIM-координатору проекта", head = "Ошибка при загрузке!")
            self.status = ["Не запускался", "Пользователь", "Неизвестно"]
            self.default_settings = {"Округлять площадь до «0.1»" : True,
                                    "Записать ТЭП" : False,
                                    "Показывать log" : True,
                                    "Расчитывать квартиры" : True,
                                    "Расчитывать неквартирные помещения" : True,
                                    "Автоматический разделитель в формуле" : True}

    def check_state(self, errors = 0, override = True):
        self.result = ""
        self.commonarea = 0.00
        collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
        if override:
            for room in collector:
                try:
                    if room.Area != 0.00:
                        self.commonarea += room.Area
                        self.result =+ "{}-{}-{}+".format(room.Get_Parameter(self.builtin_room_name_par).AsString(), str(room.LookupParameter(self.cb_parameters[10].Text).AsString()), str(room.LookupParameter(self.cb_parameters[7].Text).AsDouble()))
                except: pass
        else:
            self.result = "error"
            self.commonarea = 0.00
        if self.result == "error":
            return "расчет не завершен"
        if errors == 0:
            if str(self.commonarea_former) == str(0.00):
                return "расчет не проводился"
            elif str(self.commonarea) == str(self.commonarea_former):
                if self.result == self.result_former:
                    return "актуально"
                else:
                    return "требуется перерасчет (помещения изменены)"
            else:
                return "требуется перерасчет (площадь помещений изменена)"
        else:
            if str(self.commonarea_former) == str(0.00):
                return "расчет не завершен (ошибок - {})".format(str(errors))
            elif str(self.commonarea) == str(self.commonarea_former):
                if self.result == self.result_former:
                    return "требуется перерасчет (ошибок - {})".format(str(errors))
                else:
                    return "требуется перерасчет (помещения изменены, ошибок - {})".format(str(errors))
            else:
                return "требуется перерасчет (площадь помещений изменена, ошибок - {})".format(str(errors))

    def save_settings_manual(self, override = True):
        try:
            if override:
                self.status[2] = self.check_state(errors = self.errors_amount_former)
            else:
                self.status[2] = self.check_state(errors = self.errors_amount_former, override = False)
            parts = REVIT.query.get_central_path(doc=revit.doc).split("\\")
            path = ""
            for i in range(0, len(parts)-1):
                path += "{}\\".format(parts[i])
            path += "data"
            if not os.path.exists(path):
                os.makedirs(path)
            filepath = "{}\\rooms_settings.txt".format(path)
            file = open(filepath, "w+")
            t_pars = ""
            for i in range(0, len(self.rooms_names)):
                if i == 0:
                    t_pars += "{}_-_{}".format(self.rooms_names[i], self.cb[i].Text)
                else:
                    t_pars += "::{}_-_{}".format(self.rooms_names[i], self.cb[i].Text)
            f_pars = ""
            for j in self.formula_parts:
                if f_pars == "":
                    f_pars += "{}".format(j)
                else:
                    f_pars += "::{}".format(j)
            f_pars_o = ""
            for j in self.formula_parts_o:
                if f_pars_o == "":
                    f_pars_o += "{}".format(j)
                else:
                    f_pars_o += "::{}".format(j)
            text = "{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{}".format(str(self.chbx_settings[0].Checked),
                                                        str(self.chbx_settings[1].Checked),
                                                        str(self.chbx_settings[2].Checked),
                                                        str(self.chbx_settings[3].Checked),
                                                        str(self.chbx_settings[4].Checked),
                                                        str(self.chbx_settings[5].Checked),
                                                        str(self.status[0]),
                                                        str(self.status[1]),
                                                        str(self.status[2]),
                                                        str(self.result),
                                                        str(self.cb_parameters[0].Text),
                                                        str(self.cb_parameters[1].Text),
                                                        str(self.cb_parameters[2].Text),
                                                        str(self.cb_parameters[3].Text),
                                                        str(self.cb_parameters[4].Text),
                                                        str(self.cb_parameters[5].Text),
                                                        str(self.cb_parameters[6].Text),
                                                        str(self.cb_parameters[7].Text),
                                                        str(self.cb_parameters[8].Text),
                                                        str(self.cb_parameters[9].Text),
                                                        str(self.cb_parameters[10].Text),
                                                        str(self.cb_parameters[11].Text),
                                                        str(self.cb_parameters[12].Text),
                                                        str(self.cb_parameters[13].Text),
                                                        str(self.cb_parameters[14].Text),
                                                        str(self.cb_parameters[15].Text),
                                                        str(self.cb_parameters[16].Text),
                                                        str(self.cb_parameters[17].Text),
                                                        t_pars,
                                                        f_pars,
                                                        str(self.commonarea),
                                                        str(self.errors_amount),
                                                        f_pars_o)
            file.write(text.encode('utf-8'))
            file.close()
            self.bt[3].Enabled = False
            self.update_meta(self.status[2], reset = True)
        except: self.show_alert(text = "Ошибка при попытке сохранить текущие настройки квартирографии! Изменения не сохранены.\n\nПримечание: необходимо убедиться, что файл настроен для совместной работы или файл с настройками параметров не занят (см. папка хранилища\\\\data\\\\rooms_settings.txt)", head = "Ошибка при сохранении!")

    def save_settings(self, sender, args):
        try:
            self.status[2] = self.check_state(errors = self.errors_amount_former)
            parts = REVIT.query.get_central_path(doc=revit.doc).split("\\")
            path = ""
            for i in range(0, len(parts)-1):
                path += "{}\\".format(parts[i])
            path += "data"
            if not os.path.exists(path):
                os.makedirs(path)
            filepath = "{}\\rooms_settings.txt".format(path)
            file = open(filepath, "w+")
            t_pars = ""
            for i in range(0, len(self.rooms_names)):
                if i == 0:
                    t_pars += "{}_-_{}".format(self.rooms_names[i], self.cb[i].Text)
                else:
                    t_pars += "::{}_-_{}".format(self.rooms_names[i], self.cb[i].Text)
            f_pars = ""
            for j in self.formula_parts:
                if f_pars == "":
                    f_pars += "{}".format(j)
                else:
                    f_pars += "::{}".format(j)
            f_pars_o = ""
            for j in self.formula_parts_o:
                if f_pars_o == "":
                    f_pars_o += "{}".format(j)
                else:
                    f_pars_o += "::{}".format(j)
            text = "{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{};{}".format(str(self.chbx_settings[0].Checked),
                                                        str(self.chbx_settings[1].Checked),
                                                        str(self.chbx_settings[2].Checked),
                                                        str(self.chbx_settings[3].Checked),
                                                        str(self.chbx_settings[4].Checked),
                                                        str(self.chbx_settings[5].Checked),
                                                        str(self.status[0]),
                                                        str(self.status[1]),
                                                        str(self.status[2]),
                                                        str(self.result),
                                                        str(self.cb_parameters[0].Text),
                                                        str(self.cb_parameters[1].Text),
                                                        str(self.cb_parameters[2].Text),
                                                        str(self.cb_parameters[3].Text),
                                                        str(self.cb_parameters[4].Text),
                                                        str(self.cb_parameters[5].Text),
                                                        str(self.cb_parameters[6].Text),
                                                        str(self.cb_parameters[7].Text),
                                                        str(self.cb_parameters[8].Text),
                                                        str(self.cb_parameters[9].Text),
                                                        str(self.cb_parameters[10].Text),
                                                        str(self.cb_parameters[11].Text),
                                                        str(self.cb_parameters[12].Text),
                                                        str(self.cb_parameters[13].Text),
                                                        str(self.cb_parameters[14].Text),
                                                        str(self.cb_parameters[15].Text),
                                                        str(self.cb_parameters[16].Text),
                                                        str(self.cb_parameters[17].Text),
                                                        t_pars,
                                                        f_pars,
                                                        str(self.commonarea),
                                                        str(self.errors_amount),
                                                        f_pars_o)
            file.write(text.encode('utf-8'))
            file.close()
            self.bt[3].Enabled = False
            self.update_meta(self.status[2], reset = False)
        except: self.show_alert(text = "Ошибка при попытке сохранить текущие настройки квартирографии! Изменения не сохранены.\n\nПримечание: необходимо убедиться, что файл настроен для совместной работы или файл с настройками параметров не занят (см. папка хранилища\\\\data\\\\rooms_settings.txt)", head = "Ошибка при сохранении!")

    def show_alert(self, text = str(), head = "Ошибка!"):
        if self.Visible:
            self.Hide()
            Alert(text, title="KPLN Квартирография", header = head)
            self.Show()
        else:
            Alert(text, title="KPLN Квартирография", header = head)

    def go_to_help(self, sender, args):
        webbrowser.open('http://moodle/mod/book/view.php?id=502&chapterid=681/')

    def append_dict(self, key, room):
        bool = False
        for i in range(0, len(self.dict_keys)):
            if key == self.dict_keys[i]:
                self.dict_rooms[i].append(room)
                bool = True
        if not bool:
            self.dict_keys.append(key)
            self.dict_rooms.append([room])

    def check(self):
        self.Hide()
        par_dict = [self.cb_parameters[1].Text, self.cb_parameters[2].Text, self.cb_parameters[3].Text, self.cb_parameters[4].Text]
        lost_values = ""
        lost_department = ""
        not_digits = ""
        for room in self.allrooms:
            if self.get_parameter_def(room, self.cb_parameters[0].Text) == "??" or self.get_parameter_def(room, self.cb_parameters[0].Text) == "" or self.get_parameter_def(room, self.cb_parameters[0].Text) == "None":
                lost_department += "\n\t«{}» - ROOM_{}ID ({});".format(self.cb_parameters[0].Text, str(room.Id), room.get_Parameter(self.builtin_room_name_par).AsString())
            for par in par_dict:
                if self.get_parameter_def(room, par) == "??" or self.get_parameter_def(room, par) == "" or self.get_parameter_def(room, par) == "None":
                    lost_values += "\n\t«{}» - ROOM_{}ID ({});".format(par, str(room.Id), room.get_Parameter(self.builtin_room_name_par).AsString())
                if not self.get_parameter_def(room, par).isdigit():
                    if self.get_parameter_def(room, par)[1:].isdigit() and self.get_parameter_def(room, par)[0] == "-": pass
                    else:
                        not_digits += "\n\t«{}» : Значение: {} - ROOM_{}ID ({});".format(par, self.get_parameter_def(room, par), str(room.Id), room.get_Parameter(self.builtin_room_name_par).AsString())
        if lost_department != "":
            self.show_alert(text = "Отсутствует назначение:\n{}".format(lost_department), head = "Ошибка заполнения!")
            self.Show()
            return False
        if lost_values != "":
            self.show_alert(text = "Не заполнены параметры в помещениях:\n{}".format(lost_values), head = "Ошибка заполнения!")
            self.Show()
            return False
        if not_digits != "":
            self.show_alert(text = "Данные параметр должен содержать только цифры:\n{}\n\nПримечание: запрещены также знаки «,» и «.»".format(not_digits), head = "Ошибка заполнения!")
            self.Show()
            return False
        unused_params = ""
        for i in range(0,len(self.cb_parameters)):
            if self.cb_parameters[i].Text == "" or self.cb_parameters[i].Text == "<Выберите параметр>":
                unused_params += "     «{}» - не заполнен\n".format(self.default_parameters[i])
        for i in range(0,len(self.cb_parameters)):
            if self.cb_parameters[i].Text == "" or self.cb_parameters[i].Text == "<Выберите параметр>":
                self.show_alert(text = "Необходимо проверить наличие всех параметров исходных данных!\n{}\n\nПримечание: см.вкладка «Параметры», область «Рабочие параметры:»".format(unused_params), head = "Ошибка заполнения!")
                self.Show()
                return False
        if len(self.formula_parts) == 0 and self.chbx_settings[3].Checked:
            self.show_alert(text = "Необходимо заполнить формулу номера помещения!\n\nПримечание: для добавления значений в формулу необходимо перетянуть одно из доступных значений из поля доступных параметров в поле «Формула номера помещения» (см. «Drag and Drop»)", head = "Ошибка заполнения!")
            self.Show()
            return False
        if len(self.formula_parts_o) == 0 and self.chbx_settings[4].Checked:
            self.show_alert(text = "Необходимо заполнить формулу номера помещения!\n\nПримечание: для добавления значений в формулу необходимо перетянуть одно из доступных значений из поля доступных параметров в поле «Формула номера помещения» (см. «Drag and Drop»)", head = "Ошибка заполнения!")
            self.Show()
            return False
        return True

    def check_level(self, flats = list()):
        error = False
        levelid = str(flats[0].LevelId)
        flatnum = flats[0].LookupParameter(self.cb_parameters[10].Text).AsString()
        fl = []
        for room in flats:
            fl.append(self.out.linkify(room.Id))
            if str(room.LevelId) != levelid:
                error = True
        rooms = " ".join(fl)
        for room in flats:
            if str(room.LevelId) != levelid:
                error = True
        if error:
            if self.chbx_settings[2].Checked: print("\tПредупреждение: В квартире #{} - помещения находятся на разных уровнях\n\t{}".format(flatnum, rooms))

    def is_studio(self, flat):
        for r in flat:
            if r.get_Parameter(self.builtin_room_name_par).AsString() == "Кухня-гостиная" or r.get_Parameter(self.builtin_room_name_par).AsString() == "Кухня-ниша":
                return True
        return False

    def to_abc(self, value):
        if value.isdigit():
            try:
                num = int(value)
                return self.alphabet[num]
            except:
                for char in [" ", "_", "-"]:
                    try:
                        num = int(value.split(char)[0])
                        return self.alphabet[num]
                    except: pass
                return value
        else:
            return value

    def numerate(self, roomo):
        n_kor = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[1].Text))
        n_sec = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[2].Text))
        n_elev = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[3].Text))
        n_num = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[4].Text))
        department = self.get_parameter_def(roomo, self.cb_parameters[0].Text)
        code = "{}.{}.{}".format(n_kor, n_sec, n_elev)
        for i in range(0, len(self.noflatroomsdict)):
            if code == self.noflatroomsdict[i]:
                self.noflatrooms[i].append(roomo)
                return True
        self.noflatrooms.append([roomo])
        self.noflatroomsdict.append(code)
        return False

    def update_numeration(self):
        mass_dict = []
        for dict in self.noflatrooms:
            extendet_dict = []
            extendet_dict_sorted = []
            extendet_dict_keys = []
            extendet_dict_keys_sorted = []
            for roomo in dict:
                n_kor = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[1].Text))
                n_sec = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[2].Text))
                n_elev = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[3].Text))
                n_num = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[4].Text))
                department = self.get_parameter_def(roomo, self.cb_parameters[0].Text)
                code = "{}.{}.{}.{}.{}".format(n_kor, n_sec, n_elev, department, n_num)
                for i in range(0, len(extendet_dict_keys)):
                    if code == extendet_dict_keys[i]:
                        extendet_dict[i].append(roomo)
                        break
                    extendet_dict.append([roomo])
                    extendet_dict_keys.append(code)
                    extendet_dict_keys_sorted.append(code)
            extendet_dict_keys_sorted.sort()
            for i in range(0, len(extendet_dict_keys_sorted)):
                for z in range(0, len(extendet_dict_keys)):
                    if extendet_dict_keys_sorted[i] == extendet_dict_keys[z]:
                        for r in extendet_dict[z]:
                            extendet_dict_sorted.append(r)
            mass_dict.append(extendet_dict_sorted)
        self.noflatrooms = []
        for i in mass_dict:
            temp = []
            for z in i:
                temp.append(z)
                print(str(z))
            self.noflatrooms.append(temp)

    def dep_numerate(self, roomo):
        n_kor = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[1].Text))
        n_sec = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[2].Text))
        n_elev = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[3].Text))
        n_num = self.to_abc(self.get_parameter_def(roomo, self.cb_parameters[4].Text))
        n_id = str(roomo.Id)
        department = self.get_parameter_def(roomo, self.cb_parameters[0].Text)
        code = "{}.{}.{}.{}.{}".format(n_kor, n_sec, n_elev, n_num, n_id)
        code_2 = "{}.{}.{}".format(n_elev, n_sec, department)
        for i in range(0, len(self.depflatrooms_dep)):
            if code_2 == self.depflatrooms_dep[i]:
                self.depflatrooms[i].append(roomo)
                self.depflatroomsdict[i].append(code)
                self.depflatroomsdict_sorted[i].append(code)
                return True
        self.depflatrooms.append([roomo])
        self.depflatroomsdict.append([code])
        self.depflatroomsdict_sorted.append([code])
        self.depflatrooms_dep.append(code_2)
        return False

    def generate_tep(self, otherrooms, flatrooms):
        self.tep_flats = 0.00
        self.tep_flats_k = 0.00
        self.tep_flats_living = 0.00
        self.tep_flats_balcony = 0.00
        self.tep_flats_terrass = 0.00
        self.tep_flats_wet = 0.00
        self.tep_mop = 0.00
        self.tep_com = 0.00
        self.tep_tech = 0.00
        self.tep_tech_space = 0.00
        self.tep_all = 0.00
        self.tep_all_k = 0.00
        self.tep_store = 0.00
        self.tep_dou = 0.00
        self.tep_flat_amount = len(self.abs_numeration)
        self.tep_store_amount = 0
        #DEPARTMENTDICT
        tep_department_dict = []
        tep_department_dict_area = []
        #NO FLATS
        if self.chbx_settings[4].Checked:
            for room in otherrooms:
                try:
                    tep_department = self.get_parameter_def(room, self.cb_parameters[0].Text)
                    if not self.in_list(tep_department, tep_department_dict):
                        tep_department_dict.append(tep_department)
                        tep_department_dict_area.append(round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1))
                    else:
                        for i in range(0, len(tep_department_dict)):
                            if tep_department == tep_department_dict[i]:
                                tep_department_dict_area[i] += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                    self.tep_all += round(room.LookupParameter(self.cb_parameters[7].Text).AsDouble() * 0.09290304, 1)
                    self.tep_all_k += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                    name = room.get_Parameter(self.builtin_room_name_par).AsString()
                    if name and name != "":
                        for n in range(0, len(self.rooms_names)):
                            if name == self.rooms_names[n].split("+")[1]:
                                if self.cb[n].Text == "Аренда":
                                    self.tep_com += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                                elif self.cb[n].Text == "Тех.помещения":
                                    self.tep_tech += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                                elif self.cb[n].Text == "Тех.пространства":
                                    self.tep_tech_space += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                                elif self.cb[n].Text == "МОП":
                                    self.tep_mop += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                                elif self.cb[n].Text == "Кладовые":
                                    self.tep_store += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                                    self.tep_store_amount += 1
                                elif self.cb[n].Text == "ДОУ":
                                    self.tep_dou += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                                else: pass
                except: break
        #FLATS
        if self.chbx_settings[3].Checked:
            for room in flatrooms:
                try:
                    self.tep_all += round(room.LookupParameter(self.cb_parameters[7].Text).AsDouble() * 0.09290304, 1)
                    self.tep_all_k += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                    self.tep_flats += round(room.LookupParameter(self.cb_parameters[7].Text).AsDouble() * 0.09290304, 1)
                    self.tep_flats_k += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                    name = room.get_Parameter(self.builtin_room_name_par).AsString()
                    for n in range(0, len(self.rooms_names)):
                        if name == self.rooms_names[n].split("+")[1]:
                            if self.cb[n].Text == "Кв: Жилое пом.":
                                self.tep_flats_living += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                            elif self.cb[n].Text == "Кв: Нежилое пом. (мокрое)":
                                self.tep_flats_wet += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                            elif self.cb[n].Text == "Кв: Лоджия" or self.cb[n].Text == "Кв: Балкон":
                                self.tep_flats_balcony += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                            elif self.cb[n].Text == "Кв: Терраса":
                                self.tep_flats_terrass += round(room.LookupParameter(self.cb_parameters[8].Text).AsDouble() * 0.09290304, 1)
                            else: pass
                except: break
        #WRITE DATA
        self.default_project_values = [self.tep_flats_living,
                                    self.tep_flats,
                                    self.tep_flats_k,
                                    self.tep_mop,
                                    self.tep_tech,
                                    self.tep_tech_space,
                                    self.tep_com,
                                    self.tep_store,
                                    self.tep_flat_amount,
                                    self.tep_store_amount,
                                    self.tep_dou]
        with db.Transaction(name = "write_tep"):
            for b in range(0, len(self.default_project_parameters)):
                try:
                    if str(type(self.default_project_values[b])) == str(type(int())):
                        doc.ProjectInformation.LookupParameter(self.default_project_parameters[b]).Set(self.default_project_values[b])
                    elif str(type(self.default_project_values[b])) == str(type(float())):
                        doc.ProjectInformation.LookupParameter(self.default_project_parameters[b]).Set(self.default_project_values[b] / 0.09290304)
                except:
                    if self.chbx_settings[2].Checked: print("\t\tОшибка: В сведениях о проекте отсутствует параметр «{}»\n".format(self.default_project_parameters[b]))
        if self.chbx_settings[2].Checked:
            print("\t\tПлощадь квартир (жилая)\t{}\n".format(str(self.tep_flats_living)))
            print("\t\tПлощадь лоджий и балконов\t{}\n".format(str(self.tep_flats_balcony)))
            print("\t\tПлощадь мокрых помещений\t{}\n".format(str(self.tep_flats_wet)))
            print("\t\tПлощадь террасс\t{}\n".format(str(self.tep_flats_terrass)))
            print("\t\tОбщая площадь квартир:\t{}\n".format(str(self.tep_flats)))
            print("\t\tОбщая площадь квартир с коэфф.:\t{}\n".format(str(self.tep_flats_k)))
            print("\t\tКол-во квартир:\t{}\n".format(self.tep_flat_amount))
            print("\t\tМОП\t{}\n".format(str(self.tep_mop)))
            print("\t\tДОУ\t{}\n".format(str(self.tep_dou)))
            print("\t\tАренда\t{}\n".format(str(self.tep_com)))
            print("\t\tТехнические помещения\t{}\n".format(str(self.tep_tech)))
            print("\t\tТехнические пространства\t{}\n".format(str(self.tep_tech_space)))
            print("\t\tОбщая площадь\t{}\n".format(str(self.tep_all)))
            print("\t\tОбщая площадь с коэфф.\t{}\n".format(str(self.tep_all_k)))
            print("\t\tОбщая площадь кладовых.\t{}\n".format(str(self.tep_store)))
            print("\t\tКол-во кладовых:\t{}\n".format(str(self.tep_store_amount)))
            self.out.print_md('###DEPERTMENT AREA')
            for i in range(0, len(tep_department_dict)):
                print("\t\t{} с коэфф.:\t{}".format(tep_department_dict[i], tep_department_dict_area[i]))
        self.tep = "Наименование\tПлощадь\n"
        self.tep += "Площадь квартир (жилая)\t{}\n".format(str(self.tep_flats_living))
        self.tep += "Площадь лоджий и балконов\t{}\n".format(str(self.tep_flats_balcony))
        self.tep += "Площадь мокрых помещений\t{}\n".format(str(self.tep_flats_wet))
        self.tep += "Площадь террасс\t{}\n".format(str(self.tep_flats_terrass))
        self.tep += "Общая площадь квартир:\t{}\n".format(str(self.tep_flats))
        self.tep += "Общая площадь квартир с коэфф.:\t{}\n".format(str(self.tep_flats_k))
        self.tep += "Кол-во квартир:\t{}\n".format(self.tep_flat_amount)
        self.tep += "МОП\t{}\n".format(str(self.tep_mop))
        self.tep += "ДОУ\t{}\n".format(str(self.tep_dou))
        self.tep += "Аренда\t{}\n".format(str(self.tep_com))
        self.tep += "Технические помещения\t{}\n".format(str(self.tep_tech))
        self.tep += "Технические пространства\t{}\n".format(str(self.tep_tech_space))
        self.tep += "Общая площадь\t{}\n".format(str(self.tep_all))
        self.tep += "Общая площадь с коэфф.\t{}\n".format(str(self.tep_all_k))
        self.tep += "Общая площадь кладовых.\t{}\n".format(str(self.tep_store))
        self.tep += "Кол-во кладовых:\t{}\n".format(str(self.tep_store_amount))
        dialog = FolderBrowserDialog()
        dialog.Description = "Выберите папку для сохранения таблицы ТЭП"
        if (dialog.ShowDialog(self) == DialogResult.OK):
            try:
                filepath = "{}\\tep.txt".format(dialog.SelectedPath)
                file = open(filepath, "w+")
                file.write(self.tep.encode('utf-8'))
                file.close()
            except: pass
    def in_list(self, element, list):
        for i in list:
            if i == element:
                return True
        return False

    def get_parameter_def(self, roomp, parameter):
        if parameter == self.cb_parameters[0].Text:
            return str(roomp.get_Parameter(self.builtin_room_department_par).AsString())
        if parameter.startswith("Помещение: "):
            return str(roomp.LookupParameter(parameter[11:]).AsString())
        if parameter.startswith("Связанный уровень: "):
            level = doc.GetElement(roomp.LevelId)
            return str(level.LookupParameter(parameter[19:]).AsString())
        try:
            value = str(roomp.Lookupparameter(parameter).AsString())
            if value != "" and value.lower() != "none" and value.lower() != "null":
                return str(roomp.Lookupparameter(parameter).AsString())
            else:
                return "??"
        except: return "??"

    def get_ingroup_num(self, room):
        for department_rooms in self.noflatrooms:
            number = 0
            for r in department_rooms:
                if str(room.Id) == str(r.Id):
                    return number
        return 0

    def run(self, sender, args):
        with forms.ProgressBar(title='Инициализация . . .') as pb:
            try:
                pb.update_progress(1, max_value = 8)
                pb.title = 'Сбор информации . . .'
                if self.check() and str(REVIT.query.get_central_path(doc=revit.doc)) > "":
                    #self.Hide()
                    if self.chbx_settings[2].Checked:
                        self.out = script.get_output()
                    #VARIABLES
                    pb.update_progress(2, max_value = 8)
                    self.department_dict = ["Квартира", "МОП"]
                    self.dict_keys = []
                    self.dict_keys_sorted = []
                    self.dict_rooms = []
                    self.dict_rooms_sorted = []
                    missing_values_rooms = []
                    unidentified_rooms = []
                    error_param_rooms = []
                    unplaced_rooms = []
                    write_errors = []
                    flat_rooms = []
                    other_rooms = []
                    self.noflatrooms = []
                    self.noflatroomsdict = []
                    self.depflatrooms = []
                    self.depflatrooms_dep = []
                    self.depflatrooms_sorted = []
                    self.depflatroomsdict = []
                    self.depflatroomsdict_sorted = []
                    par_rooms_korpus = self.cb_parameters[1].Text
                    par_rooms_section = self.cb_parameters[2].Text
                    par_rooms_elevate = self.cb_parameters[3].Text
                    par_rooms_flatnum = self.cb_parameters[4].Text
                    par_rooms_roomnum = self.cb_parameters[6].Text #НОМЕР ДЛЯ ЗАПИСИ НОМЕРА ПОМЕЩЕНИЯ
                    par_rooms_inflat_num = self.cb_parameters[5].Text
                    par_flat_num_abs = self.cb_parameters[10].Text
                    par_flat_name = self.cb_parameters[11].Text
                    par_flat_description = self.cb_parameters[9].Text
                    par_area_flat_fact = self.cb_parameters[12].Text
                    par_area_flat_k = self.cb_parameters[16].Text
                    par_area_flat_living = self.cb_parameters[13].Text
                    par_area_room_fact = self.cb_parameters[7].Text
                    par_area_room_k = self.cb_parameters[8].Text
                    par_area_flat_balcony = self.cb_parameters[14].Text
                    par_area_flat_unliving = self.cb_parameters[15].Text
                    par_area_flat_heat = self.cb_parameters[17].Text
                    #GATHERINFORMATION
                    for room in self.allrooms:
                        self.dep_numerate(room)
                        dep = self.get_parameter_def(room, self.cb_parameters[0].Text)
                        if not self.in_list(dep, self.department_dict):
                            self.department_dict.append(dep)
                        try:
                            if self.get_parameter_def(room, self.cb_parameters[0].Text) != "":
                                if room.Area > 0:
                                    if self.get_parameter_def(room, self.cb_parameters[0].Text) == "Квартира":
                                        if self.chbx_settings[3].Checked:
                                            flat_rooms.append(room)
                                    else:
                                        if self.chbx_settings[4].Checked:
                                            other_rooms.append(room)
                                else:
                                    unplaced_rooms.append(room)
                            else:
                                unidentified_rooms.append(room)
                                if self.chbx_settings[2].Checked: print("\tПредупреждение: {}- отсутствует назначение. Помещение будет исключено из расчета".format(self.out.linkify(room.Id)))
                        except:
                            if self.chbx_settings[2].Checked: print("\t{}- ошибка чтения параметров".format(self.out.linkify(room.Id)))
                            error_param_rooms.append(room)
                    for i in self.depflatroomsdict_sorted:
                        i.sort()
                    for n in range(0, len(self.depflatroomsdict_sorted)):
                        list = []
                        for m in range(0, len(self.depflatroomsdict_sorted[n])):
                            for o in range(0, len(self.depflatroomsdict[n])):
                                if self.depflatroomsdict_sorted[n][m] == self.depflatroomsdict[n][o]:
                                    list.append(self.depflatrooms[n][o])
                        self.depflatrooms_sorted.append(list)
                    if len(unplaced_rooms) != 0:
                        self.out.print_html("<font color=#ff6666>\t\t<b>Неразмещенных / избыточных / неокруженных помещений</b> - {}</font>".format(str(len(unplaced_rooms))))
                    if len(unidentified_rooms) != 0:
                        if self.chbx_settings[2].Checked: self.out.print_html("<font color=#ff6666>\t\t<b>Помещений без назначения</b> - {}</font>".format(str(len(unidentified_rooms))))
                    if len(error_param_rooms) != 0:
                        if self.chbx_settings[2].Checked: self.out.print_html("<font color=#ff6666>\t\t<b>Ошибок</b> - {}</font>".format(str(len(error_param_rooms))))
                    if self.chbx_settings[2].Checked: self.out.print_html("\n\t\t<b>Квартирных помещений</b> - {}".format(str(len(flat_rooms))))
                    if self.chbx_settings[2].Checked: self.out.print_html("\t\t<b>Прочих помещений</b> - {}".format(str(len(other_rooms))))
                    if self.chbx_settings[2].Checked: self.out.print_html("\t\t<b>Всего помещений:</b> {}".format(str(len(self.allrooms))))
                    #NOT FLAT AREA
                    pb.title = 'Расчет неквартирных помещений . . .'
                    pb.update_progress(3, max_value = 8)
                    if self.chbx_settings[4].Checked:
                        with db.Transaction(name = "ta"):
                            for r in other_rooms:
                                for n in range(0, len(self.rooms_names)):
                                    if self.rooms_names[n] == "{}+{}".format(r.get_Parameter(self.builtin_room_department_par).AsString(), r.get_Parameter(self.builtin_room_name_par).AsString()):
                                        self.numerate(r)
                                        if self.chbx_settings[0].Checked:
                                            if self.cb[n].Text == "Кв: Лоджия":
                                                s_revit = round(r.Area * 0.09290304 * 0.5, 1)
                                                par = r.LookupParameter(par_area_room_k)
                                                par.Set(s_revit / 0.09290304)
                                                par = r.LookupParameter(par_area_room_fact)
                                                par.Set(s_revit * 2 / 0.09290304)
                                            elif self.cb[n].Text == "Кв: Балкон" or self.cb[n].Text == "Кв: Терраса":
                                                s_revit = round(r.Area * 0.09290304, 1)
                                                par = r.LookupParameter(par_area_room_fact)
                                                par.Set(s_revit / 0.09290304)
                                                par = r.LookupParameter(par_area_room_k)
                                                par.Set(round(s_revit * 0.3, 2) / 0.09290304)
                                            else:
                                                s_revit = round(r.Area * 0.09290304, 1)
                                                par = r.LookupParameter(par_area_room_k)
                                                par.Set(s_revit / 0.09290304)
                                                par = r.LookupParameter(par_area_room_fact)
                                                par.Set(s_revit / 0.09290304)
                                        else:
                                            if self.cb[n].Text == "Кв: Лоджия":
                                                s_revit = r.Area / 2
                                                par = r.LookupParameter(par_area_room_k)
                                                par.Set(s_revit)
                                                par = r.LookupParameter(par_area_room_fact)
                                                par.Set(s_revit * 2)
                                            elif self.cb[n].Text == "Кв: Балкон" or self.cb[n].Text == "Кв: Терраса":
                                                s_revit = r.Area
                                                par = r.LookupParameter(par_area_room_k)
                                                par.Set(s_revit * 0.3)
                                                par = r.LookupParameter(par_area_room_fact)
                                                par.Set(s_revit)
                                            else:
                                                s_revit = r.Area
                                                par = r.LookupParameter(par_area_room_k)
                                                par.Set(s_revit)
                                                par = r.LookupParameter(par_area_room_fact)
                                                par.Set(s_revit)
                                        par = r.LookupParameter(par_area_flat_fact)
                                        par.Set(0)
                                        par = r.LookupParameter(par_area_flat_k)
                                        par.Set(0)
                                        par = r.LookupParameter(par_area_flat_living)
                                        par.Set(0)
                                        par = r.LookupParameter(par_area_flat_heat)
                                        par.Set(0)
                            for department_rooms in self.noflatrooms:
                                number = 1
                                for room in department_rooms:
                                    par = room.LookupParameter(par_rooms_inflat_num)
                                    par.Set(str(number))
                                    number += 1
                    self.update_numeration()
                    #GET DICTS
                    pb.title = 'Создание словарей . . .'
                    pb.update_progress(4, max_value = 8)
                    if self.chbx_settings[3].Checked:
                        for r in flat_rooms:
                            kor = self.to_abc(self.get_parameter_def(r, par_rooms_korpus))
                            sec = self.to_abc(self.get_parameter_def(r, par_rooms_section))
                            elev = self.to_abc(self.get_parameter_def(r, par_rooms_elevate))
                            num = self.to_abc(self.get_parameter_def(r, par_rooms_flatnum))
                            k = "{}.{}.{}.{}".format(kor, sec, elev, num)
                            if kor and kor != "" and sec and sec != "" and elev and elev != "" and num and num != "":
                                self.append_dict(key = k, room = r)
                            else:
                                missing_values_rooms.append(r)
                            if not self.in_list(k, self.abs_numeration):
                                self.abs_numeration.append(k)
                        self.abs_numeration.sort()
                        if self.chbx_settings[2].Checked: 
                            self.out.print_html("\t\t<b>Количество квартир:</b> {}".format(str(len(self.abs_numeration))))
                        self.dict_keys_sorted = self.dict_keys
                        self.dict_keys_sorted.sort()
                        for i in range(0, len(self.dict_keys_sorted)):
                            for z in range(0, len(self.dict_keys)):
                                if self.dict_keys_sorted[i] == self.dict_keys[z]:
                                    self.dict_rooms_sorted.append(self.dict_rooms[z])
                    #GET FLAT AREA
                    pb.title = 'Расчет квартирных помещений'
                    pb.update_progress(5, max_value = 8)
                    if self.chbx_settings[3].Checked:
                        #ERR
                        with db.Transaction(name = "ta"):
                            for i in range(0, len(self.dict_rooms_sorted)):
                                s_flat_fact = 0.00
                                s_flat_k = 0.00
                                s_flat_living = 0.00
                                s_flat_unliving = 0.00
                                s_flat_balcony_k = 0.00
                                s_flat_heat = 0.00
                                s_living_rooms_count = 0
                                rooms_living = []
                                rooms_notliving = []
                                rooms_balcony = []
                                rooms_sorting = []
                                for r in self.dict_rooms_sorted[i]:
                                    name = r.get_Parameter(self.builtin_room_name_par).AsString()
                                    s_revit = round(r.Area * 0.09290304, 1)
                                    for n in range(0, len(self.rooms_names)):
                                        if self.rooms_names[n].split("+")[0] == r.get_Parameter(self.builtin_room_department_par).AsString():
                                            if name == self.rooms_names[n].split("+")[1]:
                                                if self.cb[n].Text == "Кв: Лоджия":
                                                    rooms_balcony.append(r)
                                                    if self.chbx_settings[0].Checked: 
                                                        s_revit = round(r.Area * 0.09290304 * 0.5, 1)
                                                        par = r.LookupParameter(par_area_room_k)
                                                        par.Set(s_revit / 0.09290304)
                                                        s_flat_k += s_revit / 0.09290304
                                                        s_flat_balcony_k  += s_revit / 0.09290304
                                                        par = r.LookupParameter(par_area_room_fact)
                                                        par.Set(s_revit / 0.09290304 * 2)
                                                        s_flat_fact += s_revit / 0.09290304 * 2
                                                    else:
                                                        s_revit = r.Area
                                                        par = r.LookupParameter(par_area_room_k)
                                                        par.Set(s_revit * 0.5)
                                                        s_flat_k += s_revit * 0.5
                                                        s_flat_balcony_k += s_revit * 0.5
                                                        par = r.LookupParameter(par_area_room_fact)
                                                        par.Set(s_revit)
                                                        s_flat_fact += s_revit
                                                elif self.cb[n].Text == "Кв: Балкон" or self.cb[n].Text == "Кв: Терраса":
                                                    rooms_balcony.append(r)
                                                    if self.chbx_settings[0].Checked: 
                                                        s_revit = round(r.Area * 0.09290304 * 0.3, 1)
                                                        par = r.LookupParameter(par_area_room_k)
                                                        par.Set(s_revit / 0.09290304)
                                                        s_flat_k += s_revit / 0.09290304
                                                        s_flat_balcony_k += s_revit / 0.09290304
                                                        s_revit = round(r.Area * 0.09290304, 1)
                                                        par = r.LookupParameter(par_area_room_fact)
                                                        par.Set(s_revit / 0.09290304)
                                                        s_flat_fact += s_revit / 0.09290304
                                                    else:
                                                        s_revit = r.Area
                                                        par = r.LookupParameter(par_area_room_k)
                                                        par.Set(s_revit * 0.3)
                                                        s_flat_k += s_revit * 0.3
                                                        s_flat_balcony_k += s_revit * 0.3
                                                        par = r.LookupParameter(par_area_room_fact)
                                                        par.Set(s_revit)
                                                        s_flat_fact += s_revit
                                                else:
                                                    if self.chbx_settings[0].Checked: 
                                                        s_revit = round(r.Area * 0.09290304, 1)
                                                        if self.cb[n].Text == "Кв: Жилое пом.":
                                                            rooms_living.append(r)
                                                            s_flat_living += s_revit / 0.09290304
                                                        elif self.cb[n].Text.startswith("Кв: Нежилое"):
                                                            s_flat_unliving += s_revit / 0.09290304
                                                            rooms_notliving.append(r)
                                                        par = r.LookupParameter(par_area_room_k)
                                                        par.Set(s_revit / 0.09290304)
                                                        s_flat_k += s_revit / 0.09290304
                                                        par = r.LookupParameter(par_area_room_fact)
                                                        par.Set(s_revit / 0.09290304)
                                                        s_flat_fact += s_revit / 0.09290304
                                                    else:
                                                        s_revit = r.Area
                                                        if self.cb[n].Text == "Кв: Жилое пом.":
                                                            rooms_living.append(r)
                                                            s_flat_living += s_revit
                                                            s_living_rooms_count += 1
                                                        elif self.cb[n].Text.startswith("Кв: Нежилое"):
                                                            s_flat_unliving += s_revit
                                                            rooms_notliving.append(r)
                                                        par = r.LookupParameter(par_area_room_k)
                                                        par.Set(s_revit)
                                                        s_flat_k += s_revit
                                                        par = r.LookupParameter(par_area_room_fact)
                                                        par.Set(s_revit)
                                                        s_flat_fact += s_revit
                                for r in self.dict_rooms_sorted[i]:
                                    par = r.LookupParameter(par_area_flat_fact)
                                    par.Set(s_flat_fact)
                                    par = r.LookupParameter(par_area_flat_k)
                                    par.Set(s_flat_k)
                                    par = r.LookupParameter(par_area_flat_living)
                                    par.Set(s_flat_living)
                                    par = r.LookupParameter(par_area_flat_balcony)
                                    par.Set(s_flat_balcony_k)
                                    par = r.LookupParameter(par_area_flat_unliving)
                                    par.Set(s_flat_unliving)
                                    s_flat_heat = s_flat_k - s_flat_balcony_k
                                    par = r.LookupParameter(par_area_flat_heat)
                                    par.Set(s_flat_heat)
                                    kor = self.to_abc(self.get_parameter_def(r, par_rooms_korpus))
                                    sec = self.to_abc(self.get_parameter_def(r, par_rooms_section))
                                    elev = self.to_abc(self.get_parameter_def(r, par_rooms_elevate))
                                    num = self.to_abc(self.get_parameter_def(r, par_rooms_flatnum))
                                    
                                    if kor and kor != "" and sec and sec != "" and elev and elev != "" and num and num != "":
                                        k = "{}.{}.{}.{}".format(kor, sec, elev, num)
                                        for g in range(0, len(self.abs_numeration)):
                                            
                                            if k == self.abs_numeration[g]:
                                                par = r.LookupParameter(par_flat_num_abs)
                                                par.Set(str(g+1))
                                #Проверка на ошибки и предупреждения в квартирах
                                pb.title = 'Поиск ошибок . . .'
                                pb.update_progress(6, max_value = 8)
                                self.check_level(self.dict_rooms_sorted[i])
                                op_rooms = []
                                for r in rooms_living:
                                    try:
                                        op_rooms.append(self.out.linkify(r.Id))
                                    except :
                                        pass
                                if len(rooms_living) == 0:
                                    print(str(self.dict_rooms_sorted[i][0]))
                                    if self.chbx_settings[2].Checked: self.out.print_html('<font color=#ff6666><b>Ошибка:</b></font> В квартире <b>#{}</b> - отсутствуют жилые помещения!'.format(self.dict_rooms_sorted[i][0].LookupParameter(par_flat_num_abs).AsString()))
                                    par = r.LookupParameter(par_flat_name)
                                    par.Set("NA")
                                    par = r.LookupParameter(par_flat_description)
                                    par.Set("Без жилых помещений")
                                elif len(rooms_living) == 1:
                                    if self.is_studio(self.dict_rooms_sorted[i]):
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("C")
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("Студия")
                                    else:
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("1К")
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("Однокомнатная квартира")
                                elif len(rooms_living) == 2:
                                    if not self.is_studio(self.dict_rooms_sorted[i]):
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("2К")
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("Двухкомнатная квартира")
                                    else:
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("2Е")
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("Двухкомнатная квартира (евро)")
                                elif len(rooms_living) == 3:
                                    if not self.is_studio(self.dict_rooms_sorted[i]):
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("3К")
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("Трехкомнатная квартира")
                                    else:
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("3Е")
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("Трехкомнатная квартира (евро)")
                                elif len(rooms_living) == 4:
                                    if not self.is_studio(self.dict_rooms_sorted[i]):
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("4К")
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("Четырехкомнатная квартира")
                                    else:
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("4Е")
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("Четырехкомнатная квартира (евро)")
                                elif len(rooms_living) == 5:
                                    if not self.is_studio(self.dict_rooms_sorted[i]):
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("5К")
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("Пятикомнатная квартира")
                                        if self.chbx_settings[2].Checked: self.out.print_html('\n<font color=#edbe00><b>Предупреждение:</b></font> В квартире <b>#{}</b> - 5 жилых помещений {}'.format(self.dict_rooms_sorted[i][0].LookupParameter(par_flat_num_abs).AsString(), " ".join(op_rooms)))
                                    else:
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("5Е")
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("Пятикомнатная квартира (евро)")
                                        if self.chbx_settings[2].Checked: self.out.print_html('\n<font color=#edbe00><b>Предупреждение:</b></font> В квартире <b>#{}</b> - 5 жилых помещений {}'.format(self.dict_rooms_sorted[i][0].LookupParameter(par_flat_num_abs).AsString(), " ".join(op_rooms)))
                                elif len(rooms_living) > 5:
                                    if not self.is_studio(self.dict_rooms_sorted[i]):
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("{}К".format(str(len(rooms_living))))
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("{}-комнатная квартира".format(str(len(rooms_living))))
                                        if self.chbx_settings[2].Checked: self.out.print_html('\n<font color=#edbe00><b>Предупреждение:</b></font> В квартире #{} - больше 5-ти жилых помещений {}'.format(self.dict_rooms_sorted[i][0].LookupParameter(par_flat_num_abs).AsString(), " ".join(op_rooms)))
                                    else:
                                        for r in self.dict_rooms_sorted[i]:
                                            par = r.LookupParameter(par_flat_name)
                                            par.Set("{}Е".format(str(len(rooms_living))))
                                            par = r.LookupParameter(par_flat_description)
                                            par.Set("{}-тикомнатная квартира (евро)".format(str(len(rooms_living))))
                                        if self.chbx_settings[2].Checked: self.out.print_html('\n<font color=#edbe00><b>Предупреждение:</b></font> В квартире <b>#{}</b> - больше 5-ти жилых помещений {}'.format(self.dict_rooms_sorted[i][0].LookupParameter(par_flat_num_abs).AsString(), " ".join(op_rooms)))
                                for roms in rooms_living:
                                    rooms_sorting.append(roms)
                                for roms in rooms_notliving:
                                    rooms_sorting.append(roms)
                                for roms in rooms_balcony:
                                    rooms_sorting.append(roms)
                                number = 1
                                for roms in rooms_sorting:
                                    try:
                                        par = roms.LookupParameter(par_rooms_inflat_num)
                                        par.Set(str(number))
                                    except :
                                        pass
                                    number += 1
                    #WRITE NUMBER
                    pb.title = 'Запись значений'
                    pb.update_progress(7, max_value = 8)
                    k = "??"
                    with db.Transaction(name = "write_final_number"):
                        if self.chbx_settings[3].Checked:
                            for room in flat_rooms:
                                value = ""
                                for v in range(0, len(self.formula_parts)):
                                    if self.formula_parts[v] == "Нумерация: номер по назначению":
                                        item = self.get_department_number(room)
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                value += ".{}".format(item)
                                            else:
                                                value += item
                                        else: value += item
                                    elif self.formula_parts[v] == "Нумерация: номер внутри назначения":
                                        pass
                                        for department_rooms in self.depflatrooms_sorted:
                                            for j in range(0, len(department_rooms)):
                                                if str(department_rooms[j].Id) == str(room.Id):
                                                    k = str(j)
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                value += ".{}".format(k)
                                            else:
                                                value += k
                                        else: value += k
                                    elif self.formula_parts[v].startswith("«"):
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                value += ".{}".format(self.formula_parts[v][2])
                                            else:
                                                value += self.formula_parts[v][2]
                                        else: value += self.formula_parts[v][2]
                                    elif self.formula_parts[v] == "Системный: id помещения":
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                value += ".{}".format(str(room.Id))
                                            else:
                                                value += str(room.Id)
                                        else: value += str(room.Id)
                                    elif self.formula_parts[v] == "Системный: рабочий набор":
                                        if self.chbx_settings[5].Checked:
                                            item = self.get_workset_name(room)
                                            if v != 0:
                                                value += ".{}".format(item)
                                            else:
                                                value += item
                                        else: value += item
                                    elif self.formula_parts[v].startswith("[Пом] "):
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                try:
                                                    item = room.LookupParameter(self.formula_parts[v][6:]).AsString()
                                                except: item == "??"
                                                if item == "": 
                                                    item = "??"
                                                    warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??»'.format(self.out.linkify(room.Id), self.formula_parts[v][6:])
                                                    if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                                value += ".{}".format(item)
                                            else:
                                                try:
                                                    item = room.LookupParameter(self.formula_parts[v][6:]).AsString()
                                                except: item == "??"
                                                if item == "":
                                                    item = "??"
                                                    warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??»'.format(self.out.linkify(room.Id), self.formula_parts[v][6:])
                                                    if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                                value += item
                                        else:
                                            try:
                                                item = room.LookupParameter(self.formula_parts[v][6:]).AsString()
                                            except: item == "??"
                                            if item == "":
                                                item = "??"
                                                warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??»'.format(self.out.linkify(room.Id), self.formula_parts[v][6:])
                                                if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                            value += item
                                    elif self.formula_parts[v].startswith("[Ур] "):
                                        level = doc.GetElement(room.LevelId)
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                try:
                                                    item = level.LookupParameter(self.formula_parts[v][5:]).AsString()
                                                except: item == "??"
                                                if item == "":
                                                    item = "??"
                                                    warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??» (ROOM{}id)'.format(self.out.linkify(level.Id), self.formula_parts[v][5:], str(room.Id))
                                                    if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                                value += ".{}".format(item)
                                            else:
                                                try:
                                                    item = level.LookupParameter(self.formula_parts[v][5:]).AsString()
                                                except: item == "??"
                                                if item == "":
                                                    item = "??"
                                                    warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??» (ROOM{}id)'.format(self.out.linkify(level.Id), self.formula_parts[v][5:], str(room.Id))
                                                    if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                                value += item
                                        else:
                                            try:
                                                item = level.LookupParameter(self.formula_parts[v][5:]).AsString()
                                            except: item == "??"
                                            if item == "":
                                                item = "??"
                                                warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??» (ROOM{}id)'.format(self.out.linkify(level.Id), self.formula_parts[v][5:], str(room.Id))
                                                if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                            value += item
                                par = room.LookupParameter(par_rooms_roomnum)
                                par.Set(value)
                        if self.chbx_settings[4].Checked:
                            for room in other_rooms:
                                value = ""
                                for v in range(0, len(self.formula_parts_o)):
                                    if self.formula_parts_o[v] == "Нумерация: номер по назначению":
                                        item = self.get_department_number(room)
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                value += ".{}".format(item)
                                            else:
                                                value += item
                                        else:
                                            value += item
                                    elif self.formula_parts_o[v].startswith("«"):
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                value += ".{}".format(self.formula_parts_o[v][2])
                                            else:
                                                value += self.formula_parts_o[v][2]
                                        else: value += self.formula_parts_o[v][2]
                                    elif self.formula_parts_o[v] == "Нумерация: номер внутри назначения":
                                        for department_rooms in self.depflatrooms_sorted:
                                            for j in range(0, len(department_rooms)):
                                                if str(department_rooms[j].Id) == str(room.Id):
                                                    k = str(j)
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                value += ".{}".format(k)
                                            else:
                                                value += k
                                        else: value += k
                                    elif self.formula_parts_o[v] == "Системный: id помещения":
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                value += ".{}".format(str(room.Id))
                                            else:
                                                value += str(room.Id)
                                        else:
                                            value += str(room.Id)
                                    elif self.formula_parts_o[v] == "Системный: рабочий набор":
                                        item = self.get_workset_name(room)
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                value += ".{}".format(item)
                                            else:
                                                value += item
                                        else:
                                            value += item
                                    elif self.formula_parts_o[v].startswith("[Пом] "):
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                try:
                                                    item = room.LookupParameter(self.formula_parts_o[v][6:]).AsString()
                                                except: item == "??"
                                                if item == "": 
                                                    item = "??"
                                                    warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??»'.format(self.out.linkify(room.Id), self.formula_parts_o[v][6:])
                                                    if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                                value += ".{}".format(item)
                                            else:
                                                try:
                                                    item = room.LookupParameter(self.formula_parts_o[v][6:]).AsString()
                                                except: item == "??"
                                                if item == "":
                                                    item = "??"
                                                    warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??»'.format(self.out.linkify(room.Id), self.formula_parts_o[v][6:])
                                                    if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                                value += item
                                        else:
                                            try:
                                                item = room.LookupParameter(self.formula_parts_o[v][6:]).AsString()
                                            except: item == "??"
                                            if item == "":
                                                item = "??"
                                                warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??»'.format(self.out.linkify(room.Id), self.formula_parts_o[v][6:])
                                                if self.chbx_settings[2].Checked: print(warning)
                                            value += item
                                    elif self.formula_parts_o[v].startswith("[Ур] "):
                                        level = doc.GetElement(room.LevelId)
                                        if self.chbx_settings[5].Checked:
                                            if v != 0:
                                                try:
                                                    item = level.LookupParameter(self.formula_parts_o[v][5:]).AsString()
                                                except: item == "??"
                                                if item == "":
                                                    item = "??"
                                                    warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??» (ROOM{}id)'.format(self.out.linkify(level.Id), self.formula_parts_o[v][5:], str(room.Id))
                                                    if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                                value += ".{}".format(item)
                                            else:
                                                try:
                                                    item = level.LookupParameter(self.formula_parts_o[v][5:]).AsString()
                                                except: item == "??"
                                                if item == "":
                                                    item = "??"
                                                    warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??» (ROOM{}id)'.format(self.out.linkify(level.Id), self.formula_parts_o[v][5:], str(room.Id))
                                                    if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                                value += item
                                        else:
                                            try:
                                                item = level.LookupParameter(self.formula_parts_o[v][5:]).AsString()
                                            except: item == "??"
                                            if item == "":
                                                item = "??"
                                                warning = '<font color=#ff6666><b>Ошибка записи номера помещения:</b></font> {} отсутствует значение параметра «{}». Установлено по умолчанию: «??» (ROOM{}id)'.format(self.out.linkify(level.Id), self.formula_parts_o[v][5:], str(room.Id))
                                                if self.chbx_settings[2].Checked: self.out.print_html(warning)
                                            value += item
                                par = room.LookupParameter(par_rooms_roomnum)
                                par.Set(value)
                    self.errors_amount = len(missing_values_rooms) + len(unidentified_rooms) + len(error_param_rooms) + len(unplaced_rooms)
                    if self.errors_amount == 0:
                        result = "\n\t\t<b>Готово</b>"
                    else:
                        result = "\n\t\t<b>Ошибок</b> - {}".format(str(self.errors_amount))
                    if self.chbx_settings[2].Checked: self.out.print_html(result)
                    now = datetime.datetime.now()
                    self.status = ["{}-{}-{} {}:{}:{}".format(now.year, now.month, now.day, now.hour, now.minute, now.second), revit.username, result]
                    if self.chbx_settings[1].Checked:
                        self.generate_tep(other_rooms, flat_rooms)
                    self.save_settings_manual()
                    self.load_settings()
                    self.Close()
            except Exception as e: 
                self.save_settings_manual(override = False)
                self.load_settings()
                self.show_alert("Окно будет закрыто", str(e))
                self.Close()

    def get_workset_name(self, room):
        workset_collector = DB.FilteredWorksetCollector(doc).ToWorksets()
        for ws in workset_collector:
            if ws.Id == room.WorksetId:
                ws_name = ws.Name
                return ws_name
        return "??"

    def get_department_number(self, room):
        department = room.get_Parameter(self.builtin_room_department_par).AsString()
        for i in range(0, len(self.department_dict)):
            if department == self.department_dict[i]:
                return str(i)
        return "??"



if revit.doc.IsWorkshared:
    form = CreateWindow()
    Application.Run(form)
else:
    Alert("Файл не настроен для совместной работы!", title="KPLN Квартирография", header = "Ошибка")

