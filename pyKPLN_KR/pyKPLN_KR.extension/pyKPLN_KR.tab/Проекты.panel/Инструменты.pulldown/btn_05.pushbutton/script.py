# -*- coding: utf-8 -*-
"""
DIV_Level

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "DIV Этаж"
__doc__ = 'Заполнение параметра ДИВ_Этаж_Текст' \
"""
"""
from pyrevit.framework import clr
import re
import os
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction
from System.Collections.Generic import *
from System.Windows.Forms import *
from System.Drawing import *
import System
from System import Guid
import datetime
from System.Drawing import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert

with db.Transaction(name = "DIV"):
	for element in DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
		if element != None:
			try:
				#element.LookupParameter("ДИВ_Этаж").Set(doc.GetElement(element.Host.LevelId).LookupParameter("ДИВ_Этаж").AsString())
				element.LookupParameter("ДИВ_Этаж_Текст").Set(doc.GetElement(element.Host.LevelId).LookupParameter("ДИВ_Этаж_Текст").AsString())
			except:
				try:
					#element.LookupParameter("ДИВ_Этаж").Set(doc.GetElement(element.LevelId).LookupParameter("ДИВ_Этаж").AsString())
					element.LookupParameter("ДИВ_Этаж_Текст").Set(doc.GetElement(element.LevelId).LookupParameter("ДИВ_Этаж_Текст").AsString())
				except:
					try:
						#element.LookupParameter("ДИВ_Этаж").Set(element.Host.LookupParameter("ДИВ_Этаж").AsString())
						element.LookupParameter("ДИВ_Этаж_Текст").Set(element.Host.LookupParameter("ДИВ_Этаж_Текст").AsString())
					except: pass