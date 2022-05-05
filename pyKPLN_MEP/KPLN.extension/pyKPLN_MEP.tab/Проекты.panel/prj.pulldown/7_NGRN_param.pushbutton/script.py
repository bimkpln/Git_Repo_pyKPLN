# -*- coding: utf-8 -*-

__title__ = "НГРН_Заполнение параметров 'Этаж'"
__author__ = 'Tima Kutsko'
__doc__ = "Автоматическое заполнение параметров этажа для ИОС."


import Autodesk
from Autodesk.Revit.DB import *
import System
from System import Enum, Guid
import re
from rpw import revit, db, ui
from pyrevit import script, forms



#classes
class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value  
    def __nonzero__(self):
        return self.state    
    def __bool__(self):
        return self.state



#definitions
def get_true_level_name(element):
	try:	
		centerPoint = Autodesk.Revit.DB.XYZ((element.Location.Curve.GetEndPoint(0).X + element.Location.Curve.GetEndPoint(1).X) / 2, (element.Location.Curve.GetEndPoint(0).Y
											+ element.Location.Curve.GetEndPoint(1).Y) / 2, (element.Location.Curve.GetEndPoint(0).Z + element.Location.Curve.GetEndPoint(1).Z) / 2)
	except:				
		centerPoint = element.Location.Point
	boundingBoxElement = element.get_Geometry(Options()).GetBoundingBox()
	center_boundBox_Point = XYZ(centerPoint.X, centerPoint.Y, boundingBoxElement.Min.Z)
	##arch_elements analising
	min_distance = None	
	min_floor_level_name = None
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
						##finded the distance between elements (mep vs arch)							 
						distance = (centerPointPrj.DistanceTo(center_boundBox_Point)) * 304.8						
						if min_distance is None or distance < min_distance:
							min_distance = distance	
							min_floor_level_name = floor.get_Parameter(BuiltInParameter.LEVEL_PARAM).AsValueString()
		except:					
			pass		
	return min_floor_level_name


def create_check_boxes_by_name(elements):
    # Create check boxes for elements if they have Name property.   
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements, key=lambda x: x.Name) if not e.GetLinkDocument() is None]
    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect=True,
                                                    title='Выбери подгруженные модели, из которой брать помещения:',
                                                    width=600,
                                                    button_name='Запуск!')
    return elements_checkboxes



# input form 0
""" commands= [ui.forms.CommandLink('Заполнение пар-ра этажа', return_value='Заполнение пар-ра этажа'),
			ui.forms.CommandLink('Заполнение пар-ра секции', return_value='Заполнение пар-ра секции')]
dialog = ui.forms.TaskDialog('Выбери формат теста',																
							commands=commands,
							buttons=['Cancel'],
							footer='ДИВ_Заполнение параметров',
							show_close=True)
strat_flag = dialog.show() """


#main code
doc = revit.doc
output = script.get_output()			
## get all elements of current project
category_list = [BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves, 
				BuiltInCategory.OST_FlexPipeCurves, BuiltInCategory.OST_FlexDuctCurves,
				BuiltInCategory.OST_MechanicalEquipment, BuiltInCategory.OST_DuctAccessory, 
				BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_DuctTerminal, 
				BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_PipeFitting, 
				BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit, 
				BuiltInCategory.OST_ElectricalEquipment, BuiltInCategory.OST_CableTrayFitting,
				BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_LightingFixtures, 
				BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_DataDevices,
				BuiltInCategory.OST_LightingDevices, BuiltInCategory.OST_NurseCallDevices,
				BuiltInCategory.OST_SecurityDevices, BuiltInCategory.OST_FireAlarmDevices,
				BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_CommunicationDevices]				


# input form 1
guid_level_param = Guid("9eabf56c-a6cd-4b5c-a9d0-e9223e19ea3f")	    #КП_О_Этаж


# input form 1.1
try:
    components = [ui.forms.flexform.Label("Узел ввода данных"),				
                ui.forms.flexform.CheckBox('all_elements', 'Все элементы в проекте', default=False),
                ui.forms.flexform.CheckBox('all_elements_on_view', 'Все элементы на активном виде', default=True),					
                ui.forms.flexform.Label("Введи часть имени перекрытия для проверок"),			  
                ui.forms.flexform.Label("(для быстрого анализа - используй"),
                ui.forms.flexform.Label("плиты перекрытия вместо полов):"),
                ui.forms.flexform.TextBox("nameARCH", Text="Жб"),					
                ui.forms.flexform.Separator(),
                ui.forms.flexform.Button("Запуск")]
    form = ui.forms.FlexForm("Автоматическое заполнение параметров для классификатора", components)
    form.ShowDialog()
    all_elements = form.values["all_elements"]
    all_elements_on_view = form.values["all_elements_on_view"]		
    nameARCH = form.values["nameARCH"]	
except:
    ui.forms.Alert("Ошибка ввода данных. Обратись к разработчику!", title="Внимание", exit=True)	


# selecting and analazing mep and arch elements 
elements_list = list()
if all_elements:			
    for current_category in category_list:
        elements_in_project = FilteredElementCollector(doc).OfCategory(current_category).WhereElementIsNotElementType().ToElements()
        elements_list.extend(elements_in_project)	
elif all_elements_on_view:	
    for current_category in category_list:
        elements_at_active_view = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(current_category).WhereElementIsNotElementType().ToElements()
        elements_list.extend(elements_at_active_view)	
else:
    ui.form.Alert("Выбери только один тип выобра элементов!", title="Внимание", exit=True)

# analazing links
linkModelInstances = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.RevitLinkInstance)	
linkModels_checkboxes = create_check_boxes_by_name(linkModelInstances)	
if linkModels_checkboxes:
    elementsARCH = list()
    linkModels = [c.value for c in linkModels_checkboxes if c.state == True] 
    with db.Transaction("ДИВ_Заполнение пар-ра этажа"):
        for link in linkModels:
            transform = link.GetTotalTransform()
            elementsARCH = FilteredElementCollector(link.GetLinkDocument()).OfClass(Autodesk.Revit.DB.Floor).ToElements()
            
            #main code 1		
            for current_item in elements_list:
                try:
                    element = doc.GetElement(current_item.id)			
                except:
                    element = current_item
                try:
                    super_component = element.SuperComponent														
                except:
                    super_component = None								
                if super_component is None:												
                    true_floor_level_name = get_true_level_name(element)							
                    if not true_floor_level_name is None:
                        ## get&set true level param						
                        """ lvl_pat = re.compile("(\-?\d\d)\_{1}")
                        lvl_index = re.findall(lvl_pat, str(true_floor_level_name))[0] """
                        lvl_index = true_floor_level_name.split(" ")[0]
                        if '-' in lvl_index:
                            index = "П"
                        else:
                            index = str(lvl_index)
                        try:
                            element.get_Parameter(guid_level_param).Set(index)	
                        except:
                            pass                                                #В исключения попадут все элементы, у которых нет параметра!        
    
    # finish flag
    ui.forms.Alert("Завершено!", title="Завершено") 

""" ## input form 1
elif strat_flag == 'Заполнение пар-ра секции':
    guid_system_param = Guid("1b2a7570-5cf6-4db9-9490-d8a147466c53")	#ДИВ_Система_Текст
    cnt = 0			
	flag = True
	guid_section_param = Guid("a3513e5d-b1f8-4f60-ae09-5c0fb579458a") #ДИВ_Секция_Текст			
	try:
		components = [ui.forms.flexform.Label("Узел ввода данных"),				
					ui.forms.flexform.Label("Стартовый номер для наземной части:"),			  
					ui.forms.flexform.TextBox("startNumber", Text="1"),	
					ui.forms.flexform.Label("Номер для подземной части:"),			  
					ui.forms.flexform.TextBox("startNumber_downFloor", Text="1"),		
					ui.forms.flexform.Separator(),
					ui.forms.flexform.Button("Запуск")]
		form = ui.forms.FlexForm("Заполнение пар-ра секции", components)
		form.ShowDialog()
		startNumber = int(form.values["startNumber"]) - 1	
		startNumber_downFloor = int(form.values["startNumber_downFloor"])
	except:
		ui.forms.Alert("Ошибка ввода данных. Обратись к разработчику!", title="Внимание", exit=True)
	ui.forms.Alert('Выбери оси - границы секций')	
	grids_list = list()	
	while flag:																	
		try:			
			selectedReference = ui.Pick.pick_element(msg='Выбери оси - границы секций. По окончанию - нажми "Esc"!')
			element = doc.GetElement(selectedReference.id)					
			if element.Category.Name == "Оси":
				grid = element									
				if grid.Id not in [g.Id for g in grids_list]:
					grids_list.append(grid)
					cnt += 1						
					if cnt%4 == 0:
						startNumber += 1
						count = 0							
						elements_in_outline_list = list()
						min_intersection_point = None
						max_intersection_point = None
						while count < 4:																			
							current_grid1 = grids_list[count]
							curve_grid1 = current_grid1.Curve	
							count += 1							
							for current_grid2 in grids_list:									
								if current_grid1.Id != current_grid2.Id:
									curve_grid2 = current_grid2.Curve
									intersect = curve_grid1.Intersect(curve_grid2)																			
									if intersect == SetComparisonResult.Overlap:											
										start_point1 = curve_grid1.GetEndPoint(1)
										end_point1 = curve_grid1.GetEndPoint(0)
										start_point2 = curve_grid2.GetEndPoint(1)
										end_point2 = curve_grid2.GetEndPoint(0)										
										x1 = start_point1.X
										y1 = start_point1.Y
										x2 = end_point1.X
										y2 = end_point1.Y
										x3 = start_point2.X
										y3 = start_point2.Y
										x4 = end_point2.X
										y4 = end_point2.Y										
										current_intersect_point = XYZ(
																((x1*y2-y1*x2)*(x3-x4) - (x1-x2)*(x3*y4-y3*x4)) /
																((x1-x2)*(y3-y4) - (y1-y2)*(x3-x4)),
																((x1*y2-y1*x2)*(y3-y4) - (y1-y2)*(x3*y4-y3*x4)) /
																((x1-x2)*(y3-y4) - (y1-y2)*(x3-x4)),
																0
																)		
										if min_intersection_point is None or current_intersect_point.X < min_intersection_point.X + 0.001 \
																			and current_intersect_point.Y < min_intersection_point.Y + 0.001:											
											min_intersection_point = current_intersect_point											
										if max_intersection_point is None or current_intersect_point.X + 0.001 > max_intersection_point.X \
																			and current_intersect_point.Y + 0.001 > max_intersection_point.Y:											
											max_intersection_point = current_intersect_point											
						true_outline = Outline(XYZ(min_intersection_point.X, min_intersection_point.Y, -15), XYZ(max_intersection_point.X, max_intersection_point.Y, 100))
						bounding_box_filter = BoundingBoxIntersectsFilter(true_outline)								
						for current_category in category_list:
							elements_in_outline = FilteredElementCollector(doc).OfCategory(current_category).WhereElementIsNotElementType().WherePasses(bounding_box_filter).ToElements()							
							if elements_in_outline:
								elements_in_outline_list.extend(elements_in_outline)
						with db.Transaction("ДИВ_Заполнение пар-ра секции № {}".format(startNumber)):
							for current_element in elements_in_outline_list:								
								try:
									super_component = current_element.SuperComponent									
								except:
									super_component = None								
								if super_component is None:
									try:	
										centerPoint = Autodesk.Revit.DB.XYZ((current_element.Location.Curve.GetEndPoint(0).X + current_element.Location.Curve.GetEndPoint(1).X) / 2, (current_element.Location.Curve.GetEndPoint(0).Y
																			+ current_element.Location.Curve.GetEndPoint(1).Y) / 2, (current_element.Location.Curve.GetEndPoint(0).Z + current_element.Location.Curve.GetEndPoint(1).Z) / 2)
									except:				
										centerPoint = current_element.Location.Point															
									#ВНИМАНИЕ!!! Задан отступ в -200 мм от уровня +0.000 из-за труб ОВ
									if centerPoint.Z > -0.66:				
										try:
											current_element.get_Parameter(guid_section_param).Set(str(startNumber)	)									
										except:	
											output.print_md('Элемент **{} с id {}** не имеет нужных параметров!'.format(current_element.Name, 
																												output.linkify(current_element.Id)))
									else:
										try:
											current_element.get_Parameter(guid_section_param).Set(str(startNumber_downFloor))										
										except:	
											output.print_md('Элемент **{} с id {}** не имеет нужных параметров!'.format(current_element.Name, 
																												output.linkify(current_element.Id)))
						ui.forms.Alert("Выбери снова 4 оси, либо нажми 'Esc'", title="Внимание")
						grids_list = list()
				else:
					ui.forms.Alert("Оси не должны повторяться!!! Можно продолжить выбор осей", title="Внимание")
			else:
				ui.forms.Alert("Только оси!!!", title="Внимание")									
		except: 
			ui.forms.Alert("Завершено!", title="Завершено")
			flag = False """				