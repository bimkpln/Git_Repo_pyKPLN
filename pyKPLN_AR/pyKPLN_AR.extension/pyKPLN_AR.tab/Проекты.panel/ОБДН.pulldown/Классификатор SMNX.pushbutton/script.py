# -*- coding: utf-8 -*-

import io
import clr
clr.AddReference('RevitAPI')
from rpw import doc, DB, db
from Autodesk.Revit.DB import *

codes = []
elements = []
none_class_elements = []
rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

with io.open('Y:\\Жилые здания\\Обыденский\\6.BIM\\6.Классификатор\\KPLN_Классификатор\\АР\\SMNX_Классификатор.txt', "r", encoding='utf-8') as f:
    lines = f.readlines()
    for i in lines:
        parts = i.split(';')
        codes.append([parts[0], parts[1]])

# классифицируемые категории:
categories = [
    DB.BuiltInCategory.OST_Walls, 
    DB.BuiltInCategory.OST_Floors, 
    DB.BuiltInCategory.OST_Doors, 
    DB.BuiltInCategory.OST_Windows, 
    DB.BuiltInCategory.OST_StairsRailing, 
    DB.BuiltInCategory.OST_CurtainWallPanels, 
    DB.BuiltInCategory.OST_MechanicalEquipment, 
    DB.BuiltInCategory.OST_GenericModel, 
    DB.BuiltInCategory.OST_Parking, 
    DB.BuiltInCategory.OST_Doors
    ]
for category in categories:
    elements += DB.FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements()
# неклассифицируемые категории:
none_class_categories = [
    DB.BuiltInCategory.OST_PlumbingFixtures, 
    DB.BuiltInCategory.OST_Furniture
    ]
for category in none_class_categories:
    none_class_elements += DB.FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().ToElements()

def GetDescription(value, list):
    for i in list:
        if value == i[0]:
            return i[1]

def SetValue(element, value):
    description = GetDescription(value, codes)
    element.LookupParameter("SMNX_Код по классификатору").Set(value)
    element.LookupParameter("SMNX_Имя_классификатора").Set(description)
    element.LookupParameter("BIM_АР_Индикатор").Set("BIM_Классификатор")

with db.Transaction(name = "Запись кодов"):
    try:
        for element in none_class_elements:
            SetValue(element, '9999')
            if "710_" in element.Name or "710_" in element.Symbol.Family.Name:
                SetValue(element, '8888')
        for element in elements:
            if "160_Кол" in element.Name:
                SetValue(element, '3.4.1.1.5.1')
    except:
        print(element.Id)
    







# with db.Transaction(name = "ta"):
#     for room in rooms:
#         r_type_com = room.LookupParameter("SMNX_Тип_ком.помещений").AsString()
#         r_level = room.LookupParameter("SMNX_Этаж").AsString()

#         s_room_koef = room.LookupParameter("ПОМ_Площадь_К").AsValueString()
#         room.LookupParameter("SMNX_Площадь с коэффициентом").Set(s_room_koef)














