# -*- coding: utf-8 -*-
"""
Project_Mirrored_Windows

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Зеркальные"
__doc__ = 'Выбрать все зеркальнные окна на виде\nФильтр : если имя семейства начинается с «120_»' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr

# clr.AddReference("RevitNodes")

import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import Alert

import datetime
def logger(result):
	try:
		now = datetime.datetime.now()
		filename = "{}-{}-{}_{}-{}-{}_{}_Зеркальные окна.txt".format(str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second), revit.username)
		file = open("Z:\\Отдел BIM\\PyRevitReports\\{}".format(filename), "w+")
		text = "kpln report\nfile:{}\nversion:{}\nuser:{}\nresult: {}".format(doc.PathName, revit.version, revit.username, result)
		file.write(text.encode('utf-8'))
		file.close()
	except: pass

out = script.get_output()
view = doc.ActiveView
application = revit.uiapp.Application
documents = application.Documents
collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(DB.BuiltInCategory.OST_Windows)
wnd_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Windows)
windows = collector.WhereElementIsNotElementType().ToElements()
doc_windows = wnd_collector.WhereElementIsNotElementType().ToElements()
mir_windows = []
doc_mir_windows = []
counter = 0
counter_max = 0
Linked_Docs = ""

for window in windows:
    try:
        if window.Mirrored and window.Symbol.Family.Name.startswith("120_"):
            mir_windows.append(window)
    except AttributeError:
        pass  # for Symbols that don't have Mirrored attribute.

for window in doc_windows:
    try:
        if window.Mirrored and window.Symbol.Family.Name.startswith("120_"):
            doc_mir_windows.append(window)
    except AttributeError:
        pass  # for Symbols that don't have Mirrored attribute.

out.print_md('###RESULT')	
if len(mir_windows) == 0:
	print("Не найдено ни одного зеркального окна на виде!\n\nВсего в активном проекте зеркальных окон - {}".format(str(len(doc_mir_windows))))
else:
	print("На виде найдено {} окон;\n\nВсего в активном проекте зеркальных окон - {}".format(str(len(mir_windows)), str(len(doc_mir_windows))))
out.print_md('###DETAILS')	
for window in doc_mir_windows:
	if revit.doc.IsWorkshared:
		wti = DB.WorksharingUtils.GetWorksharingTooltipInfo(revit.doc, window.Id)
		print("\n{} : id{}\n      автор - {}\n      изменил - {};\n".format(window.Symbol.Family.Name, str(window.Id), wti.Creator, wti.LastChangedBy))
	else:
		print("\n - {} : {};".format(window.Symbol.Family.Name, str(window.Id)))
log_text = "active view_{}\ndocument_{}".format(str(len(mir_windows)), str(len(doc_mir_windows)))
logger(log_text)

if len(mir_windows) != 0:
	selection = uidoc.Selection
	collection = List[DB.ElementId]([door.Id for door in mir_windows])
	selection.SetElementIds(collection)