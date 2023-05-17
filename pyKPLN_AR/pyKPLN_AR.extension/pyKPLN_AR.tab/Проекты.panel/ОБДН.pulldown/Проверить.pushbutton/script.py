# -*- coding: utf-8 -*-

import clr
clr.AddReference('RevitAPI')
from rpw import doc, DB
from pyrevit import script

out = script.get_output()

rooms = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()
parkings = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Parking).WhereElementIsNotElementType().ToElements()

rooms_errors = []
parkings_errors = []
rooms_params = ["SMNX_Код по классификатору", 
                "SMNX_Имя_классификатора", 
                "SMNX_Этаж", 
                "SMNX_Секция", 
                "SMNX_Коэффициент", 
                "SMNX_Высота_потолка", 
                "SMNX_Типоразмер", 
                "SMNX_Количество спален", 
                "SMNX_Вид балкона", 
                "SMNX_Тип планировки", 
                "SMNX_Сторона света", 
                "SMNX_Мокрая зона", 
                "SMNX_Этаж_Имя", 
                "SMNX_Тип_ком.помещений", 
                "SMNX_Высота_этажа", 
                "SMNX_Пожарный отсек"]
parkings_params = ["SMNX_Код по классификатору", "SMNX_Имя_классификатора", "SMNX_Этаж", "SMNX_Этаж_Имя", "SMNX_Секция", "SMNX_Типоразмер", "SMNX_Номер_парк.места", "SMNX_Тип_ком.помещений"]

def get_unique(errors):
    unique = []
    for item in errors:
        if item not in unique:
            unique.append(item)
    return unique

def get_value(parameter):
    if str(parameter.StorageType) == 'String':
        value = parameter.AsString()
    if str(parameter.StorageType) == 'Double':
        value = parameter.AsDouble()
    if str(parameter.StorageType) == 'Integer':
        value = parameter.AsValueString()
    return value


def check_category(category, parameters, errors_list):
    check = True
    for i in category:
        for p in parameters:
            param = i.LookupParameter(p)
            if param != None:
                if get_value(param) == None:
                    check = False
                    errors_list.append(param.Definition.Name)
    return check




# for room in rooms:
#     for p in rooms_params:
#         param = room.LookupParameter(p)
#         if get_value(param) == None:
#             rooms_check = False
#             rooms_errors.append(param.Definition.Name)
    # out.print_html("<font color=#ff6666>\t\t<b>Неразмещенных / избыточных / неокруженных помещений</b> - {}</font>".format('kjbkjb'))

if check_category(rooms, rooms_params, rooms_errors) == False:
    out.print_html("\t<b>Для категории {} не заполнены параметры:</b></font>".format('<font color=#B22222>ПОМЕЩЕНИЯ</font>'))
    for i in get_unique(rooms_errors):
        out.print_html("\t{}</font>".format(i))
if check_category(parkings, parkings_params, parkings_errors) == False:
    out.print_html("\n\t<b>Для категории {} не заполнены параметры:</b></font>".format('<font color=#B22222>ПАРКОВКИ</font>'))
    for i in get_unique(parkings_errors):
        out.print_html("\t{}</font>".format(i))
