# -*- coding: utf-8 -*-

__title__ = "СИТИ_Заполнение параметра 'Система'для вентиляции"
__author__ = 'Tima Kutsko'
__doc__ = "Автоматическое заполнение параметров для классификатора.\n"\
		"Заполянет параметры 'Система', 'Секция' для элементов ОВ2 \n"\
		"Актуально только для ОВ2!"



from Autodesk.Revit.DB import *
from System import Guid, Enum
import re
from rpw import revit, db
from pyrevit import script, forms



#definitions
##getting builincategory from element's category
def get_BuiltInCategory(category):
	if Enum.IsDefined(BuiltInCategory, category.Id.IntegerValue):
		BuiltInCat_id = category.Id.IntegerValue
		BuiltInCat = Enum.GetName(BuiltInCategory, BuiltInCat_id)
		return BuiltInCat
	else:
		return BuiltInCategory.INVALID



# main code

## main alert
forms.alert('Формат имён для систем ОВ2 следующий: \n'
			'ОБОБЩЕННОЕ ИМЯ СИСТЕМ_ИМЯ ТИПА СИСТЕМЫ_ 666', title="Внимание")


# variables
doc = revit.doc
output = script.get_output()
# СИТИ_Система
guid_system_param = Guid("98b30303-a930-449c-b6c2-11604a0479cb")
# СИТИ_Этаж
guid_level_param = Guid("79c6dee3-bd92-4fdf-ae12-c06683c61946")


## set system data
mep_systems = FilteredElementCollector(doc).OfClass(MEPSystem).ToElements()
for system in mep_systems:	
	try:
		duct_network_elements = system.DuctNetwork		
	except:
		duct_network_elements = None			
	if not duct_network_elements is None:			
		with db.Transaction("ДИВ_Заполнение пар-ра системы: {}".format(system.Name)):
			system_name = system.Name
			try:
				system_type_data = system_name.split("_")			
				system_type = system_type_data[-2]
				for element in duct_network_elements:
					builtInCat = get_BuiltInCategory(element.Category)
					if builtInCat != "OST_DuctTerminal":
						element.get_Parameter(guid_system_param).Set(system_type)
			except:
				output.print_md("У системы **{}** - сокращение задано не по шаблону (**ОБОБЩЕННОЕ ИМЯ СИСТЕМ_ИМЯ ТИПА СИСТЕМЫ_ 666**)".format(system.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()))