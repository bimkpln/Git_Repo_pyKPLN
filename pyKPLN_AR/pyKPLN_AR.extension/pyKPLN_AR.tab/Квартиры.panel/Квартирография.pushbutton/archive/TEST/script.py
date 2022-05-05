# -*- coding: utf-8 -*-
"""
Квартирография

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Квартиры"
__doc__ = 'KPLN Квартирография разработана для автоматической нумерации помещений и определения показетелей площади. См. инструкцию по использованию скрипта внутри (значок «?»)\n' \
          'ВАЖНО: Для корректной работы файл должен быть настроен для совместной работы\n' \
          'Квартиры определяются по параметру «Назначение» (см. вкладка «Параметры» скрипта) и значение у всек квартир должно быть равно «Кв»;\n\n' \
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

class Flat():
	def __init__(self):
		#Elements
		self.LivingRooms = []
		self.UnlivingRooms = []
		self.HeatingRooms = []
		self.SummerRooms = []
		self.WetRooms = []
		self.Rooms = []
		self.Unplaced = []
		#Identification
		self.Id = 0
		self.Count = 0
		self.Key = ""
		self.Code = ""
		self.Name = ""
		self.Type = ""
		self.Levels = []
		self.Department = ""
		self.Difference = 0.0
		#Area
		self.Area = 0.0
		self.AreaK = 0.0
		self.AreaLiving = 0.0
		self.AreaUnliving = 0.0
		self.AreaHeating = 0.0
		self.AreaSummer = 0.0
		self.AreaWet = 0.0
		self.AreaLoggia = 0.0
		self.AreaBalcony = 0.0
		self.AreaTerrace = 0.0