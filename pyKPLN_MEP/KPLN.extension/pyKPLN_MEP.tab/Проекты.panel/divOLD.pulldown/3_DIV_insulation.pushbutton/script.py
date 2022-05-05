# -*- coding: utf-8 -*-

__title__ = "ДИВ_Анализ изоляции ВК/ОВ на активном виде"
__author__ = 'Tima Kutsko'
__doc__ = "Заполнение параметра ДИВ_Секция_Текст, ДИВ_Система_Текст/ДИВ_Этаж_Текст и ДИВ_RBS_Код по классификатору в изоляцию воздуховодов трубопроводов\n"\
		"Актуально только для ВК, ОВ!"



import Autodesk
from Autodesk.Revit.DB import *
from rpw import revit, db, ui
from pyrevit import script
from System import Guid



# definitions
def insulation(elements_insulation):
	for current_insulation in elements_insulation:
		for current_guid in guids_list:
			current_element = doc.GetElement(current_insulation.HostElementId)
			element_current_guid_param = current_element.get_Parameter(current_guid).AsString()
			if element_current_guid_param is None:
				element_current_guid_param = doc.GetElement(current_element.GetTypeId()).get_Parameter(current_guid).AsString()
			try:
				if not element_current_guid_param is None:				
					try:			
						current_insulation.get_Parameter(current_guid).Set(str(element_current_guid_param))
					except:
						doc.GetElement(current_insulation.GetTypeId()).get_Parameter(current_guid).Set(str(element_current_guid_param))
			except:
				output.print_md("**Элемент, у которого нет нужных праметров:** {} - {} \
									**Внимание!!! Работа скрипта остановлена. Устрани ошибку!**".format(current_insulation.Name, output.linkify(current_insulation.Id)))
				script.exit()


""" def insulation(elements_insulation):   #for lookup 
	for current_insulation in elements_insulation:
		for current_guid in guids_list:
			current_element = doc.GetElement(current_insulation.HostElementId)
			element_current_guid_param = current_element.LookupParameter(current_guid).AsString()
			if element_current_guid_param is None:
				try:	
					element_current_guid_param = doc.GetElement(current_element.GetTypeId()).LookupParameter(current_guid).AsString() 
				except:
					pass   #you can get elements with out data in try parameter
			try:
				if not element_current_guid_param is None:				
					try:
						current_insulation.LookupParameter(current_guid).Set(str(element_current_guid_param)) 
					except:
						doc.GetElement(current_insulation.GetTypeId()).LookupParameter(current_guid).Set(str(element_current_guid_param))
			except:
				output.print_md("**Элемент, у которого нет нужных праметров:** {} - {} \
									**Внимание!!! Работа скрипта остановлена. Устрани ошибку!**".format(current_insulation.Name, output.linkify(current_insulation.Id)))
				script.exit() """


# main code
doc = revit.doc
output = script.get_output()
guid_section_param = Guid("a3513e5d-b1f8-4f60-ae09-5c0fb579458a") 	#ДИВ_Секция_Текст	
guid_system_param = Guid("1b2a7570-5cf6-4db9-9490-d8a147466c53")	#ДИВ_Система_Текст
guid_level_param = Guid("830e785e-ece4-499d-943f-ffb314fe5031")		#ДИВ_Этаж_Текст	
guids_list = [guid_section_param, guid_system_param,
				guid_level_param]


""" guids_list = ["КП_ПП_Принадлежность к корпусу"]  #for lookup """


## insulations
mep_elements_insulation = list()
pipe_insulation = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().ToElements()
duct_insulation = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_DuctInsulations).WhereElementIsNotElementType().ToElements()
mep_elements_insulation.extend(pipe_insulation)
mep_elements_insulation.extend(duct_insulation)

with db.Transaction("ДИВ_Анализ изоляции ОВ/ВК"):	
	if len(mep_elements_insulation) > 0:
		insulation(mep_elements_insulation)	
		ui.forms.Alert('Выполнено успешно!', title='Анализ изоляции ОВ/ВК')		
	else:
		print("В проекте нет изоляции труб/воздуховодов. Только для ОВ/ВК!")


