# -*- coding: utf-8 -*-

__title__ = "Проверка высоты размещения элементов ИОС"
__author__ = 'Tima Kutsko'
__doc__ = "Определение регламентного расстояния между перекрытиями и инженерными сетями"



import Autodesk
from Autodesk.Revit.DB import *
import System
from System import Enum
import re
from rpw import revit, db, ui
from pyrevit import script



#definitions
#getting builincategory from element's category
def get_BuiltInCategory(category):
	if Enum.IsDefined(BuiltInCategory, category.Id.IntegerValue):
		BuiltInCat_id = category.Id.IntegerValue
		BuiltInCat = Enum.GetName(BuiltInCategory, BuiltInCat_id)
		return BuiltInCat
	else:
		return BuiltInCategory.INVALID



# main code
doc = revit.doc
output = script.get_output()
list_of_BuiltInCategories = System.Enum.GetValues(BuiltInCategory)
linkModels = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.RevitLinkInstance)
linkDict = {l.Name: l for l in linkModels if l.GetLinkDocument()}
floorTypeDict = {"Перекрытия": BuiltInCategory.OST_Floors, "Фундамент несущей конструкции": BuiltInCategory.OST_StructuralFoundation}


## get all elements of current project
parameter_set = set()
category_set = set()
elements_list = []
categoryList = ["OST_DuctCurves", "OST_PipeCurves", 
				"OST_MechanicalEquipment", "OST_DuctAccessory", 
				"OST_DuctFitting", "OST_DuctTerminal", 
				"OST_PipeAccessory", "OST_PipeFitting", 
				"OST_CableTray", "OST_Conduit",
				"OST_CableTrayFitting", "OST_ConduitFitting",
				"OST_LightingFixtures", "OST_ElectricalFixtures", 
				"OST_ElectricalEquipment"]

all_elements_at_active_view = FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
for current_element in all_elements_at_active_view:
	if current_element.Category:
		if get_BuiltInCategory(current_element.Category) in categoryList:
			elements_list.append(current_element)
			## parameters of elements			
			for current_param in current_element.Parameters:
				if current_param.Definition.ParameterType.ToString() == 'Text' and not current_param.IsShared:
					parameter_set.add(current_param)
parameterDict_height =  {p.Definition.Name: p.Definition.Name for p in parameter_set}


## input form 2
try:
	components = [ui.forms.flexform.Label("Узел ввода данных"),				
				ui.forms.flexform.CheckBox('selection', 'Выбранный фрагмент', default=True),
				ui.forms.flexform.CheckBox('all_elements', 'Все элементы на плане', default=False),
				ui.forms.flexform.Label("Выбери пар-р для присв. высоты элемента:"),
				ui.forms.flexform.ComboBox("paramName_height", parameterDict_height),				
				ui.forms.flexform.Label("Выбери модель из списка (нужен файл АР):"),
				ui.forms.flexform.ComboBox("link", linkDict),
				ui.forms.flexform.Label("Тип перекрытия для проверки:"),
				ui.forms.flexform.ComboBox("floorType", floorTypeDict),
				ui.forms.flexform.Label("Введи часть имени перекрытия для проверок"),			  
				ui.forms.flexform.Label("(для быстрого анализа - используй"),
				ui.forms.flexform.Label("плиты перекрытия вместо полов):"),
				ui.forms.flexform.TextBox("nameARCH", Text="Жб"),					
				ui.forms.flexform.Separator(),
				ui.forms.flexform.Button("Запуск")]
	form = ui.forms.FlexForm("Заполнение параметров для расчёта креплений ЭОМ", components)
	form.ShowDialog()
		
	
	selection = form.values["selection"]
	all_elements = form.values
	paramName_height = form.values["paramName_height"]	
	link = form.values["link"]
	builtInCat = form.values["floorType"]
	nameARCH = form.values["nameARCH"]	
except:
	ui.forms.Alert("Ошибка ввода данных. Обратись к разработчику!", title="Внимание", exit=True)	


## selecting and analazing mep and arch elements 
transform = link.GetTotalTransform()
elementsARCH = FilteredElementCollector(link.GetLinkDocument()).OfCategory(builtInCat).WhereElementIsNotElementType().ToElements()
if selection:
	selected_elements = ui.Pick.pick_element(msg='Выбери фрагмент для анализа', multiple=True)
elif all_elements:
	selected_elements = elements_list
else:
	ui.form.Alert("Выбери только один тип выобра элементов!", title="Внимание", exit=True)


#main code
with db.Transaction("pyKPLN_Проверка высоты элементов"):
	for current_item in selected_elements:
		try:
			element = doc.GetElement(current_item.id)			
		except:
			element = current_item		
		if get_BuiltInCategory(element.Category) in categoryList:			
			try:	
				centerPoint = Autodesk.Revit.DB.XYZ((element.Location.Curve.GetEndPoint(0).X + element.Location.Curve.GetEndPoint(1).X) / 2, (element.Location.Curve.GetEndPoint(0).Y
													+ element.Location.Curve.GetEndPoint(1).Y) / 2, (element.Location.Curve.GetEndPoint(0).Z + element.Location.Curve.GetEndPoint(1).Z) / 2)			
				if get_BuiltInCategory(element.Category) == "OST_DuctCurves":
					try:
						element_height = element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsValueString()
					except:
						element_height = element.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM).AsValueString()
					element_insulation_height_param = element.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsValueString()
				elif get_BuiltInCategory(element.Category) == "OST_PipeCurves":
					element_height = element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsValueString()
					element_insulation_height_param = element.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsValueString()
				elif get_BuiltInCategory(element.Category) == "OST_CableTray":
					element_height = element.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM).AsValueString()
					element_insulation_height_param = None
				element_height = int(re.match(r'\d*', element_height).group(0))
				if element_insulation_height_param:
					element_insulation_height = int(re.match(r'\d*', element_insulation_height_param).group(0))
				else:
					element_insulation_height = 0				
			except:				
				centerPoint = element.Location.Point
				element_height = 100000 #points dosen't have a slope!			
			boundingBoxElement = element.get_Geometry(Options()).GetBoundingBox()
			center_boundBox_Point = XYZ(centerPoint.X, centerPoint.Y, boundingBoxElement.Min.Z)
			##find elements with slope with 10% error				
			error_distance_check = centerPoint.DistanceTo(center_boundBox_Point) * 304.8
			if error_distance_check > element_height/1.9: 	
				error_distance = error_distance_check - element_height/2
			else:
				error_distance = 0
			###arch_elements analising
			min_distance = None
			for floor in elementsARCH:
				try:					
					if nameARCH.upper() in floor.Name.upper():
						centerPointPrj = floor.GetVerticalProjectionPoint(centerPoint, Autodesk.Revit.DB.FloorFace.Top)	
						boundingBoxFloor = floor.get_Geometry(Options()).GetBoundingBox()						
						transformFloorMaxPoint = transform.OfPoint(boundingBoxFloor.Max)
						transformFloorMinPoint = transform.OfPoint(boundingBoxFloor.Min)
						checkMax = transformFloorMaxPoint.X>centerPointPrj.X and transformFloorMaxPoint.Y>centerPointPrj.Y
						checkMin = transformFloorMinPoint.X<centerPointPrj.X and transformFloorMinPoint.Y<centerPointPrj.Y	
						boolean = checkMax and checkMin																	
						if boolean:							
							if centerPointPrj.Z < boundingBoxElement.Min.Z:
								###finded the distance between elements (mep vs arch)							 
								distance = (centerPointPrj.DistanceTo(center_boundBox_Point)) * 304.8 + error_distance - element_insulation_height						
								if min_distance is None or distance  < min_distance:
									min_distance = distance																	
									min_distance_project_point_floor = floor
							###set data to current parameter		
							element.LookupParameter(paramName_height).Set(str(round(min_distance, 1)))							
				except:
					pass
					#print(min_distance_project_point_floor.Id)		