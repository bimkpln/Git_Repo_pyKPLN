# -*- coding: utf-8 -*-
"""
ROOMS_Identical_Numbers

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Дубл. номера"
__doc__ = 'Проверить проект на повторяющиеся номера помещений' \
"""
Архитекурное бюро KPLN

"""
import collections
import System
import datetime
import os
from pyrevit import revit, DB

def logger(result = "Успешно!"):
	try:
		now = datetime.datetime.now()
		filename = "{}-{}-{}_{}-{}-{}_ПРОВЕРКА ДУБЛИКАТОВ_{}.txt".format(str(now.year), str(now.month), str(now.day), str(now.hour), str(now.minute), str(now.second), revit.username)
		file = open("Z:\\Отдел BIM\\PyRevitReports\\{}".format(filename), "w+")
		text = "kpln log\nfile:{}\nversion:{}\nuser:{}\n\nlog:\n{}".format(doc.PathName, revit.version, revit.username, result)
		file.write(text.encode('utf-8'))
		file.close()
	except: pass

rooms = DB.FilteredElementCollector(revit.doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElementIds()
roomnums = [revit.doc.GetElement(rmid).Number for rmid in rooms]
duplicates = [item
              for item, count in collections.Counter(roomnums).items()
              if count > 1]

if len(duplicates) > 0:
	for rn in duplicates:
		print('IDENTICAL ROOM NUMBER:  {}'.format(rn))
		for rmid in rooms:
			rm = revit.doc.GetElement(rmid)
			if rm.Number == rn:
				print('\tROOM{}ID:  №{} ({}) : Уровень - {}'.format(str(rm.Id),rm.Parameter[DB.BuiltInParameter.ROOM_NUMBER].AsString(),rm.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString(),rm.Level.Name))
		print('\n')
		logger('\tROOM{}ID:  №{} ({}) : Уровень - {}'.format(str(rm.Id),rm.Parameter[DB.BuiltInParameter.ROOM_NUMBER].AsString(),rm.Parameter[DB.BuiltInParameter.ROOM_NAME].AsString(),rm.Level.Name))
else:
	print('Не найдено повторяющихся номеров!')
	logger('Не найдено повторяющихся номеров!')
