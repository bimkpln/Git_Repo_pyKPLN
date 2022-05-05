# -*- coding: utf-8 -*-
"""
INGRD_Level

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "INGRD Этаж"
__doc__ = 'Заполнение системного параметра INGRD_Этаж' \
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

INGD_elevation_guid = Guid("b3c207a9-2a51-4a18-a7fd-13012a9643c4")

with db.Transaction(name = "INGRD"):
	for element in DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
		if element != None:
			try:
				element.get_Parameter(INGD_elevation_guid).Set(doc.GetElement(element.Host.LevelId).get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString())
			except:
				try:
					element.get_Parameter(INGD_elevation_guid).Set(doc.GetElement(element.LevelId).get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString())
				except:
					try:
						element.get_Parameter(INGD_elevation_guid).Set(element.Host.get_Parameter(DB.BuiltInParameter.DATUM_TEXT).AsString())
					except: pass