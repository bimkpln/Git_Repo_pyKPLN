# -*- coding: utf-8 -*-

__title__ = "Уровнь для спецификации элементов 'На грани'"
__author__ = 'Tima Kutsko'
__doc__ = "Заполнение прарматра 'КП_О_Этаж' для соед. деталей и арматуры ИОС"



from Autodesk.Revit.DB import *
import re
import Autodesk
from rpw import revit, db, ui
from pyrevit import script
from rpw.ui.forms import Console
from System import Guid, Enum



#Variables
doc = revit.doc
listOfBuiltInCategories = Enum.GetValues(BuiltInCategory)
guidLevel = Guid("9eabf56c-a6cd-4b5c-a9d0-e9223e19ea3f")	#КП_О_Этаж
message = 'Готово!'
output = script.get_output()
listOfElements = []
linkModels = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.RevitLinkInstance)
linkDict = {l.Name.split(':')[0]: l for l in linkModels}
floorTypeDict = {"Перекрытия": BuiltInCategory.OST_Floors, "Фундамент несущей конструкции": BuiltInCategory.OST_StructuralFoundation}



#Forms
ComboBox = ui.forms.flexform.ComboBox
Label = ui.forms.flexform.Label
Button = ui.forms.flexform.Button
TextBox = ui.forms.flexform.TextBox
CheckBox = ui.forms.flexform.CheckBox
Separator = ui.forms.flexform.Separator
try:
	components = [Label("Узел ввода данных"),
				CheckBox('reset', 'Сбросить предыдущую работу?', default=False),
				Label("Выбери модель из списка (АР/КР):"),
				ComboBox("link", linkDict),
				Label("Выбери тип перекрытия из списка:"),			  
				ComboBox("floorType", floorTypeDict),
				Label("Введи часть имени перекрытия для проверок"),
				Label("(для быстрого анализа - используй"),
				Label("плиты перекрытия вместо полов):"),
				TextBox("nameARCH", Text="Жб"),				  
				Separator(),
				Button("Запуск")]
	form = ui.forms.FlexForm("Определение уровня спецификации", components)
	form.ShowDialog()
	link = form.values["link"]
	builtInCat = form.values["floorType"]
	nameARCH = form.values["nameARCH"]
	reset = form.values["reset"]
except:
	script.exit()



#Functions
def elementsList():	
	categoryList = ["OST_DuctCurves", "OST_PipeCurves", "OST_CableTray", "OST_Conduit", "OST_MechanicalEquipment", "OST_DuctAccessory", "OST_DuctFitting",
					"OST_DuctTerminal", "OST_PipeAccessory", "OST_PipeFitting", "OST_ElectricalEquipment", "OST_CableTrayFitting", "OST_ConduitFitting",
					"OST_LightingFixtures", "OST_ElectricalFixtures", "OST_LightingDevices"]
	for i in categoryList:
			for j in listOfBuiltInCategories:
				if str(j) == i:
					collector = FilteredElementCollector(doc).OfCategory(j).WhereElementIsNotElementType().ToElements()
					for element in collector:
						listOfElements.append(element)
	return listOfElements


def resetFunc():	
	for element in elementsList():
		try:
			element.get_Parameter(guidLevel).Set("Нет!")
		except:
			pass
	print("Сброшено!")		



#Main code
transform = link.GetTotalTransform()
floorCollector = FilteredElementCollector(link.GetLinkDocument()).OfCategory(builtInCat).WhereElementIsNotElementType().ToElements()
distanceDict = {}


with db.Transaction("pyKPLN_Уровнь для спецификации"):
	if reset:
			resetFunc()					
	else:
		try:					
			for element in elementsList():
				distanceDict.clear()
				keyList = []				
				try:				
					centerPoint = Autodesk.Revit.DB.XYZ((element.Location.Curve.Origin.X + element.Location.Curve.GetEndPoint(1).X) / 2, (element.Location.Curve.Origin.Y
														+ element.Location.Curve.GetEndPoint(1).Y) / 2, (element.Location.Curve.Origin.Z + element.Location.Curve.GetEndPoint(1).Z) / 2)							
				except:				
					centerPoint = element.Location.Point								
				for floor in floorCollector:													
					try:
						centerPointPrj = floor.GetVerticalProjectionPoint(centerPoint, Autodesk.Revit.DB.FloorFace.Top)												
						if centerPointPrj.Z < centerPoint.Z:
							boundingBoxFloor = floor.get_Geometry(Options()).GetBoundingBox()					
							transformFloorMaxPoint = transform.OfPoint(boundingBoxFloor.Max)
							transformFloorMinPoint = transform.OfPoint(boundingBoxFloor.Min)
							checkMax = transformFloorMaxPoint.X>centerPointPrj.X and transformFloorMaxPoint.Y>centerPointPrj.Y
							checkMin = transformFloorMinPoint.X<centerPointPrj.X and transformFloorMinPoint.Y<centerPointPrj.Y							
							boolean = checkMax and checkMin	
						else:
							boolean = False	
					except:
						boolean = False							
					if boolean:	
						try:
							distance = abs(centerPoint.Z - centerPointPrj.Z)
							distanceDict = {distance: floor}							
							for key in distanceDict.keys():
								keyList.append(key)							
							nearestFloor = distanceDict[min(keyList)]		
							try:
								floorLevel = nearestFloor.LookupParameter("Уровень").AsValueString()
							except AttributeError:
								floorLevel = nearestFloor.LookupParameter("Level").AsValueString()
							except:
								print("Languedge error: only English or Russian")
							element.get_Parameter(guidLevel).Set(floorLevel)
						except:
							message = 'Не все элементы имеют параметр "КП_О_Этаж". Готово!'
			print(message)
		except:
			print("Обратись к разработчику!")		