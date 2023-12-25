# -*- coding: utf-8 -*-
"""
DIV_SharedParameters

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Марка отверстий"
__doc__ = '' \
"""
Архитекурное бюро KPLN

"""
from pyrevit.framework import clr
import re
import os
import math
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction
from System.Collections.Generic import *
from System.Windows.Forms import *
from System.Drawing import *
import System
import datetime
from System.Drawing import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert

class LevelForms(Form):
	def __init__(self, keys):
		global level_keys
		global level_start
		self.level_boxes = []
		self.Name = "KPLN DIV"
		self.Text = "KPLN DIV"
		self.Size = Size(400, 500)
		self.MinimumSize = Size(400, 300)
		self.MaximumSize = Size(400, 900)
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.ControlBox = False
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.TopMost = True
		self.AutoScroll = True

		for i in range(0, len(level_keys)):
			lbl = Label()
			lbl.Parent = self
			lbl.Text = "ДИВ_Этаж : «{}» | Нумарация с:".format(level_keys[i])
			lbl.Size = Size(350, 20)
			lbl.Location = Point(10, 50 * i + 20)
			self.level_boxes.append(TextBox())
			self.level_boxes[-1].Parent = self
			self.level_boxes[-1].Size = Size(350, 20)
			self.level_boxes[-1].Location = Point(10, 50 * i + 40)
			self.level_boxes[-1].Text = "1"

		self.btn = Button(Text="OK")
		self.btn.Parent = self
		self.btn.Click += self.OnOk
		self.btn.Location = Point(10, 50 * len(level_keys) + 25)

		self.MaximumSize = Size(400, 50 * len(level_keys) + 100)
		self.Size = Size(400, 50 * len(level_keys) + 100)

	def OnOk(self, sender, args):
		global level_start
		global level_keys
		for i in range(0, len(level_keys)):
			try:
				level_start[i] = int(self.level_boxes[i].Text)
			except:
				self.Hide()
				forms.alert("Недопустимое значение!")
				self.level_boxes[i].Text = "1"
				self.Show()
				return
		self.Close()

def IsConcrete(element):
	host = element.Host.Name
	if host.startswith("00"):
		return True
	return False

def Normalize(num):
	v = str(math.fabs(num))
	if len(v) < 10:
		for i in range(0, 10 - len(v)):
			v = "0" + v
	if num < 0:
		v = "0" + v
	else:
		v = "1" + v
	return v

def get_elementsByKey(key):
	list = []
	for i in stack_elements:
		if key == i[1]:
			list.append(i[0])
	return list

def get_Parameter(name, element):
	try:
		value = element.LookupParameter(name).AsString()
		return value
	except:
		try:
			element.Symbol.LookupParameter(name).AsString()
			return value
		except:
			try:
				type = doc.GetElement(element.GetTypeId())
				value = type.LookupParameter(name).AsString()
				return value
			except:
				try:
					type = doc.GetElement(element.Symbol.GetTypeId())
					value = type.LookupParameter(name).AsString()
					return value
				except:
					return False

uniq_keys = []
level_keys = []
level_start = []
stack_elements = []
for element in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_MechanicalEquipment).WhereElementIsNotElementType().ToElements():
	try:	
		fam_name = element.Symbol.FamilyName
		if fam_name.lower().startswith("199_отверстие"):		
			par_section = get_Parameter("ДИВ_Секция", element)
			par_elevation = Normalize(int(doc.GetElement(element.LevelId).Elevation * 304.8))
			par_lid = get_Parameter("ДИВ_Этаж", element)
			par_size = get_Parameter("КП_Размер_Текст", element)
			par_height = get_Parameter("КП_Высота", element)
			par_comment = get_Parameter("КП_В_Примечание", element)
			if not par_lid in level_keys:
				level_keys.append(par_lid)
				level_start.append(1)
			key = "{}|{}|{}|{}|{}".format(par_section, par_lid, par_size, par_height, par_comment, str(IsConcrete(element)))
			if not key in uniq_keys:
				uniq_keys.append(key)
			stack_elements.append([element, key])
			

	except Exception as e:
		print(str(e))
		pass

level_keys.sort()
uniq_keys.sort()
num = 1

try:
	form = LevelForms(level_keys)
	Application.Run(form)
	print("НОМЕРАЦИЯ ПО УРОВНЯМ:")
	with db.Transaction(name = "DIV"):
		for i in range(0, len(level_keys)):
			print("\nЭтаж : «{}»".format(level_keys[i]))
			num = level_start[i]
			for z in range(0, len(uniq_keys)):
				if level_keys[i] == uniq_keys[z].split('|')[1]:
					els = get_elementsByKey(uniq_keys[z])
					if len(els) != 0:
						for e in els:
							e.LookupParameter("КП_Марка_Номер").Set(str(num))
						num += 1
			print("\tМакс номер марки - {}".format(str(num)))
except:
	forms.alert("Произошла ошибка!")

