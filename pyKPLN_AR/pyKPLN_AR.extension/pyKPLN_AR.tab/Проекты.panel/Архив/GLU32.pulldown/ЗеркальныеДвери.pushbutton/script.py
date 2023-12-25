# -*- coding: utf-8 -*-
"""
Project_Mirrored_Doors

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Зеркальные"
__doc__ = 'Выбрать все зеркальнные двери на виде\nФильтр : если имя семейства начинается с «100_»' \

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
		filename = "{}-{}-{}_{}-{}-{}_{}_Зеркальные двери.txt".format(str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second), revit.username)
		file = open("Z:\\Отдел BIM\\PyRevitReports\\{}".format(filename), "w+")
		text = "kpln report\nfile:{}\nversion:{}\nuser:{}\nresult: {}".format(doc.PathName, revit.version, revit.username, result)
		file.write(text.encode('utf-8'))
		file.close()
	except: pass

out = script.get_output()
view = doc.ActiveView
application = revit.uiapp.Application
documents = application.Documents
collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(DB.BuiltInCategory.OST_Doors)
doc_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Doors)
doors = collector.WhereElementIsNotElementType().ToElements()
doc_doors = doc_collector.WhereElementIsNotElementType().ToElements()
mir_doors = []
doc_mir_doors = []
counter = 0
counter_max = 0
Linked_Docs = ""

for door in doors:
    try:
        if door.Mirrored and door.Symbol.Family.Name.startswith("100_"):
            mir_doors.append(door)
    except AttributeError:
        pass  # for Symbols that don't have Mirrored attribute.

for door in doc_doors:
    try:
        if door.Mirrored and door.Symbol.Family.Name.startswith("100_"):
            doc_mir_doors.append(door)
    except AttributeError:
        pass  # for Symbols that don't have Mirrored attribute.

out.print_md('###RESULT')	
if len(mir_doors) == 0:
	print("Не найдено ни одной зеркальной двери на виде!\n\nВсего в активном проекте зеркальных дверей - {}".format(str(len(doc_mir_doors))))
else:
	print("На виде найдено {} дверей;\n\nВсего в активном проекте зеркальных дверей - {}".format(str(len(mir_doors)), str(len(doc_mir_doors))))
out.print_md('###DETAILS')	
for door in doc_mir_doors:
	if revit.doc.IsWorkshared:
		wti = DB.WorksharingUtils.GetWorksharingTooltipInfo(revit.doc, door.Id)
		print("\n{} : id{}\n      автор - {}\n      изменил - {};\n".format(door.Symbol.Family.Name, str(door.Id), wti.Creator, wti.LastChangedBy))
	else:
		print("\n - {} : {};".format(door.Symbol.Family.Name, str(door.Id)))
log_text = "active view_{}\ndocument_{}".format(str(len(mir_doors)), str(len(doc_mir_doors)))
logger(log_text)

if len(mir_doors) != 0:
	selection = uidoc.Selection
	collection = List[DB.ElementId]([door.Id for door in mir_doors])
	selection.SetElementIds(collection)