# -*- coding: utf-8 -*-

__title__ = "Маркировать"
__author__ = 'Tima Kutsko'
__doc__ = "Маркировка пространств или помещений"

import sys
import Autodesk
from Autodesk.Revit.DB import *
from rpw import revit, db, ui
from rpw.ui.forms import Console
from pyrevit import script


spaceList, roomList, tagList = [], [], []
doc = revit.doc


# Functions
def GetUVPoint(pt):
    if pt.GetType().ToString() == 'Autodesk.Revit.DB.XYZ':
        return Autodesk.Revit.DB.UV(pt.X, pt.Y)
    elif pt.GetType().ToString() == 'Autodesk.Revit.DB.UV':
        return Autodesk.Revit.DB.UV(pt.U, pt.V)


def CreateSpaceTag(space, uv, view):
    return doc.Create.NewSpaceTag(space, uv, view)


def CreateRoomTag(space, uv, view):
    return doc.Create.NewRoomTag(roomId, uv, view)


# Tag types
types = FilteredElementCollector(doc).OfClass(FamilySymbol)
typeTags = [t for t in types if t.ToString() == 'Autodesk.Revit.DB.Mechanical.SpaceTagType' or t.ToString() == 'Autodesk.Revit.DB.Architecture.RoomTagType']
tagTypeDict = {t.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString(): t for t in typeTags}
# Linked model
linkModels = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.RevitLinkInstance)
linkDict = {l.Name: l for l in linkModels}
if len(typeTags) == 0 or linkModels.FirstElement() is None:
    print("В проекте не хватает данных для обработки:\n"
        "1. Нет пространств/помещений;\n"
        "2. Нет подгруженных марок для пространств/помещений (хотя бы одного типа);\n"
        "3. Нет подгруженных моделей (связей)\n"
        "Все вышеперечисленные пункты содержат типовые проекты. Вы либо работаете с маленьким объектом, либо слишком рано пытаетесь запустить программу\n"
        "Для небольших объектов достаточно стандартных инструментов Revit"
    )
    script.exit()

# Selected planes
selection = ui.Selection()
views = [v.unwrap() for v in selection]
for view in views:
    try:
        view_type = view.ViewType
    except:
        view_type = None
    if view_type != ViewType.FloorPlan:
        ui.forms.Alert('Нужно выбирать только планы!', title='pyKPLN_Маркировка пространств/помещений', exit=True)	
# Form
ComboBox = ui.forms.flexform.ComboBox
Label = ui.forms.flexform.Label
Button = ui.forms.flexform.Button
CheckBox = ui.forms.flexform.CheckBox
TextBox = ui.forms.flexform.TextBox
Separator = ui.forms.flexform.Separator
if len(selection) == 0:
    ui.forms.Alert('Выбери (выдели через shift/ctrl) в модели план(-ы) для маркировки', title='pyKPLN_Маркировка пространств/помещений', exit=True)		
else:
    components = [Label("Узел ввода данных"),
            Label("1. Пространства или помещения:"),
            CheckBox('isRoom', 'Помещения', default = True),
            CheckBox('isSpace', 'Пространства'),
            Label("2. Тип марки для выбранной категории:"),
            ComboBox("tagType", tagTypeDict),
            Label("3. Связанная модель (ТОЛЬКО для помещений)?"),
            CheckBox('isLink', 'Да, модель из связанного файла', default = True),
            Separator(),
            Button("Запуск")]
    form = ui.forms.FlexForm("Маркировка помещений или пространств", components)
    form.ShowDialog()
    isSpace = form.values["isSpace"]
    isRoom = form.values["isRoom"]
    tagType = form.values["tagType"]
    isLink = form.values["isLink"]
    if isLink:
        components = [Label("Узел ввода данных"),
            Label("Выбери модель из списка (если файл связанный):"),
            ComboBox("link", linkDict),
            Separator(),
            Button("Запуск")]
        form = ui.forms.FlexForm("Маркировка помещений или пространств", components)
        form.ShowDialog()
        link = form.values["link"]



#Main code
try:
	#main part of code
	with db.Transaction('pyKPLN_Маркировка пространств/помещений'):
		#Spaces
		if isSpace and isRoom:
			print('Выбери только одну(!) категорию')
		elif isSpace and tagType.ToString() == 'Autodesk.Revit.DB.Mechanical.SpaceTagType':	
			for view in views:
				spaces = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()
				for space in spaces:
					try:				
						point = space.Location.Point
						if point:
							uv = GetUVPoint(point)														
							spaceTag = CreateSpaceTag(space, uv, view)
							spaceTag.SpaceTagType = tagType	
					except:
						print("Элемент с Id: {} не размещен".format(space.Id))						
		#Rooms
		elif isRoom and tagType.ToString() == 'Autodesk.Revit.DB.Architecture.RoomTagType':	
			if isLink:		
				linkDoc = link.GetLinkDocument()
				linkId = link.Id
				transformCoord = link.GetTransform()
			else:
				linkDoc = doc				
			rooms = FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()			
			for room in rooms:
				try:			
					loc = room.Location
					if loc:
						point = loc.Point
						if isLink:
							transformPoint = Autodesk.Revit.DB.Transform.OfPoint(transformCoord, point)
						else:
							transformPoint = point				
						uv = GetUVPoint(transformPoint)
						for view in views:					
							if isLink:	
								roomId = LinkElementId(linkId, room.Id)	
							else:
								roomId = LinkElementId(room.Id)				
							roomTag = CreateRoomTag(roomId, uv, view.Id)
							roomTag.RoomTagType = tagType
				except:
					print("Элемент с Id: {} не размещен".format(room.Id))						
		else:
			print('Не совпадает тип марки и выбранная категория')
except:
	pass