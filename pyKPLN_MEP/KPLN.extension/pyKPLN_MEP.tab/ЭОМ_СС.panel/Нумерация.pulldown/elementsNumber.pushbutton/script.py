# -*- coding: utf-8 -*-

__title__ = "Нумерация эл-в"
__author__ = 'Tima Kutsko'
__doc__ = "Задание порядкового номера для элементов схем СС"

from Autodesk.Revit.DB import *
import re
from rpw import revit, db, ui
from pyrevit import script
from System import Guid
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger


# variables
doc = revit.doc
output = script.get_output()
sortedList, sortedList_x, sortedList_y = [], [], []
step = 1


# functions
def getElements(textCategory):
	elementSelect = []
	for element in elementsSelectReference:
		docElement = doc.GetElement(element.id)
		if docElement.Category.Name == textCategory:
			elementSelect.append(docElement)
	return elementSelect


# form
try:
	components = [ui.forms.flexform.Label("Узел ввода данных"),
				ui.forms.flexform.Label("Стартовый номер:"),			  
				ui.forms.flexform.TextBox("startNumber", Text="1"),			  
				ui.forms.flexform.Label("Приставка (если нет - оставь пустым):"),
				ui.forms.flexform.TextBox("preff"),			 
				ui.forms.flexform.CheckBox('systemName', 'Имя датчика?', default=True),
				ui.forms.flexform.Label("Суффикс:"),
				ui.forms.flexform.TextBox("suff", Text=".1.1."),				
				ui.forms.Label("Выбери направление анализа:"),
				ui.forms.CheckBox('boolean', 'Слева направо/снизу вверх', default=True),
				ui.forms.CheckBox('boolean', 'Справа налево/сверху вниз'),
				ui.forms.Separator(),
				ui.forms.Button('Выбрать')]
	form = ui.forms.FlexForm("Задание порядкового номера для элемента схем СС", components)
	form.ShowDialog()
	boolean = form.values['boolean']
	startNumber = int(form.values["startNumber"])
	preff = form.values["preff"]
	systemName = form.values["systemName"]
	suff = form.values["suff"]	
except:
	script.exit()

	
ui.forms.Alert('Выбери фрагмент для анализа')
elementsSelectReference = ui.Pick.pick_element(msg='Выбери фрагмент для анализа', multiple=True)


#main code
#getting info logger about user
log_name = "ЭОМ_СС_" + str(__title__)
InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
#main part of code
startNumber -= 1
guidSystemName = Guid("ae8ff999-1f22-4ed7-ad33-61503d85f0f4")	#"КП_О_Позиция"
guidTag = Guid("2204049c-d557-4dfc-8d70-13f19715e46d")			#"КП_О_Марка

selectedDetailItems = getElements("Элементы узлов")
sortedList_x = sorted(selectedDetailItems, key=lambda s: round(s.Location.Point.X, 3), reverse=boolean)


for item in sortedList_x:	
	if round(item.Location.Point.X, 3) not in [round(i.Location.Point.X, 3) for i in sortedList]:			
		sortedList.extend(sorted(sortedList_y, key=lambda s: round(s.Location.Point.Y, 3), reverse=boolean))
		sortedList_y = []
		sortedList.append(item)
	else:
		sortedList_y.append(item)

for element in sortedList:		
	try:				
		elementType = doc.GetElement(element.GetTypeId())			
		elementSystemName = elementType.get_Parameter(guidSystemName).AsString()  
		if elementSystemName == None:
				elementSystemName = element.get_Parameter(guidSystemName).AsString() 	
		startNumber += step
		if systemName:
				newName = preff + elementSystemName + suff + str(startNumber)
		else:
			newName = preff + suff + str(startNumber)					
		with db.Transaction("pyKPLN_Нумерация группы объектов"):		
			element.get_Parameter(guidTag).Set(newName)
		
	except: 
		output.print_md("**Некорректный элемент с id:** {} \n\n"
						"Либо не назначен параметр экземпляра 'КП_О_Марка' и параметр типа/экземпляра 'КП_О_Позиция', либо был выбран ложный элемент. \n\n"
						"Можно продолжить выборку!".format(output.linkify(element.Id)))
