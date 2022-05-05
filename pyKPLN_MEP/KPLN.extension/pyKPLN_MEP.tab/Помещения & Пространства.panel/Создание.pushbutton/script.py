# -*- coding: utf-8 -*-

__title__ = "Создать"
__author__ = 'Tima Kutsko'
__doc__ = "Создание пространств на основе помещений\nиз связанных файлов"

import Autodesk
from Autodesk.Revit.DB import *
from rpw import revit, db
from pyrevit import forms, script
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger


# Classes
class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value  
    def __nonzero__(self):
        return self.state    
    def __bool__(self):
        return self.state



# Functions
def create_check_boxes_by_name(elements):
    # Create check boxes for elements if they have Name property.   
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements, key=lambda x: x.Name) if not e.GetLinkDocument() is None]
    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect=True,
                                                    title='Выбери подгруженные модели, из которой брать помещения:',
                                                    width=600,
                                                    button_name='Запуск!')
    return elements_checkboxes


def GetUVPoint(pt):
	if pt.GetType().ToString() == 'Autodesk.Revit.DB.XYZ':
		return Autodesk.Revit.DB.UV(pt.X, pt.Y)
	elif pt.GetType().ToString() == 'Autodesk.Revit.DB.UV':
		return Autodesk.Revit.DB.UV(pt.U, pt.V)


def round_def(number):
    number = number * 1000 // 1 / 1000
    return number


def GetNearestModelLevel(level):
	room_level_elevation = round_def(level.Elevation)
	for lvl in levelCollector:
		space_level_elevation = round_def(lvl.Elevation)
		if space_level_elevation == room_level_elevation:
			nearest_level = lvl						
	return nearest_level


# Main code
count = 0
error_set = set()
doc = revit.doc
output = script.get_output()
levelCollector = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.Level)
linkModelInstances = FilteredElementCollector(doc).OfClass(Autodesk.Revit.DB.RevitLinkInstance)
linkModels_checkboxes = create_check_boxes_by_name(linkModelInstances)


if linkModels_checkboxes:
    #getting info logger about user
    log_name = "Пространства и помещения_" + str(__title__)
    InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
    #main code
    with db.Transaction('pyKPLN_Создание пространств/помещений'):
        linkModels = [c.value for c in linkModels_checkboxes if c.state == True] 
        for link in linkModels:
            linkDoc = link.GetLinkDocument()		
            rooms = FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()	
            transformCoord = link.GetTransform()
            for room in rooms:	
                room_upper_offset = room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsDouble()
                loc = room.Location
                if loc:
                    point = loc.Point					
                    transformPoint = Autodesk.Revit.DB.Transform.OfPoint(transformCoord, point)										
                    uv = GetUVPoint(transformPoint)						
                    try:
                        level = GetNearestModelLevel(room.Level)								
                        newSpace = doc.Create.NewSpace(level, uv)						
                        newSpace.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(room_upper_offset)
                        count += 1	
                    except:
                        error_set.add(room.get_Parameter(BuiltInParameter.LEVEL_NAME).AsString())						
                else:
                    output.print_md("Помещение '{0}' (id:{1}) в файле {2} - не размещено! **Можно продолжить работу**".format(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(), 
                                                                                                                            room.Id, 
                                                                                                                            link.Name.split(':')[0]))
else:
	forms.alert('Нужно выбрать хотя бы одну связанную модель!')

if len(error_set) > 0:
	output.print_md("**Следующих уровней нет в твоем проекте, либо они имеют не корректную отметку:**")
	for current_error in error_set:
		print(current_error)

if count > 0:
	print("____________________________________________________________")
	output.print_md("Было создано **{}** пространств(-a)!".format(count))
