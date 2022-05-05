# -*- coding: utf-8 -*-

__title__ = "Замена номеров эл-в с определенным индексом"
__author__ = 'Tima Kutsko'
__doc__ = "Изменения номера выборки элементов"

from Autodesk.Revit.DB import *
import re
from rpw import revit, db, ui
from pyrevit import script
from System import Guid, Enum
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger


# variables
doc = revit.doc
output = script.get_output()
true_BuilInCategories = ["OST_DetailComponents", "OST_FireAlarmDevices", 
						"OST_NurseCallDevices", "OST_SecurityDevices"]
selectedDetailItems = list()


# functions
#getting builincategory from element's category
def get_BuiltInCategory(category):
	if Enum.IsDefined(BuiltInCategory, category.Id.IntegerValue):
		BuiltInCat_id = category.Id.IntegerValue
		BuiltInCat = Enum.GetName(BuiltInCategory, BuiltInCat_id)
		return BuiltInCat
	else:
		return BuiltInCategory.INVALID


def elements_filter(BuiltInCategory_text):	
	for element in elementsSelectReference:
		docElement = doc.GetElement(element.id)		
		if get_BuiltInCategory(docElement.Category) == BuiltInCategory_text:			
			selectedDetailItems.append(docElement)			
	return selectedDetailItems



# form
try:
	components = [ui.forms.flexform.Label("Узел ввода данных"),				
				ui.forms.flexform.Label("Выбери индекс числа для замены:"),
				ui.forms.flexform.TextBox("change_number_index", Text="1"),
				ui.forms.flexform.Label("Заменить на:"),			  
				ui.forms.flexform.TextBox("change_number", Text="2"),
				ui.forms.Separator(),
				ui.forms.Button('Выбрать')]	
	form = ui.forms.FlexForm("Замена номера эл-в", components)
	form.ShowDialog()	
	
	change_number_index = int(form.values["change_number_index"])
	change_number = [form.values["change_number"]]
except:
	script.exit()

	
ui.forms.Alert('Выбери фрагмент для анализа')
elementsSelectReference = ui.Pick.pick_element(msg='Выбери фрагмент для анализа', multiple=True)


#main code

#getting info logger about user
log_name = "ЭОМ_СС_" + str(__title__)
InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)

#main part of code
guidSystemName = Guid("ae8ff999-1f22-4ed7-ad33-61503d85f0f4")	#"КП_О_Позиция"
guidTag = Guid("2204049c-d557-4dfc-8d70-13f19715e46d")			#"КП_О_Марка"
for current_category in true_BuilInCategories:	
	elements_filter(current_category)


with db.Transaction("pyKPLN_Нумерация группы объектов"):
	patt_numbers = re.compile('\d+')
	patt_separator = re.compile('\W')
	for element in selectedDetailItems:
		try:				
			elementType = doc.GetElement(element.GetTypeId())			
			elementSystemName = elementType.get_Parameter(guidTag).AsString()  
			if elementSystemName == None:
					elementSystemName = element.get_Parameter(guidTag).AsString()	
			name_numbers = patt_numbers.findall(elementSystemName)						
			new_name_numbers = name_numbers[:(change_number_index -1)] + change_number + name_numbers[change_number_index:]						
			name_separator = patt_separator.findall(elementSystemName)
			name_text = elementType.get_Parameter(guidSystemName).AsString()		
			if name_text == None:
				name_text = element.get_Parameter(guidSystemName).AsString()		
			new_name = name_text
			for i in zip(name_separator, new_name_numbers):
				for j in i:				
					new_name += j			
			element.get_Parameter(guidTag).Set(new_name)			
		except: 
			output.print_md("**Некорректный элемент с id:** {} \n\n"
							"Либо не назначен параметр экземпляра 'КП_О_Марка' и параметр типа/экземпляра 'КП_О_Позиция', либо был выбран ложный элемент. \n\n"
							"Можно продолжить выборку!".format(output.linkify(element.Id)))
