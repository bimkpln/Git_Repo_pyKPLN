# -*- coding: utf-8 -*-

__title__ = "Нумерация эл-та"
__author__ = 'Tima Kutsko'
__doc__ = "Задание порядкового номера для элемента схем СС"


from Autodesk.Revit.DB import *

from rpw import revit, db, ui
from pyrevit import script
from System import Guid
from rpw.ui.forms import*
from pyrevit.forms import WPFWindow
from System.Windows import MessageBox
from pyrevitmep.event import CustomizableEvent
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger



# variables
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
cnt = 0
guidSystemName = Guid("ae8ff999-1f22-4ed7-ad33-61503d85f0f4")	#"КП_О_Позиция"
guidTag = Guid("2204049c-d557-4dfc-8d70-13f19715e46d")			#"КП_О_Марка"
Params = FilteredElementCollector(doc).OfClass(SharedParameterElement)

paramsList = dict()
for i in Params:
	paramsList[i.Name] = i
Param = SelectFromList('Выберите параметр', paramsList)

def main_func(step, flag,Param):
	#main code
	global startNumber		
	while flag:
		try:
			selectedReference = ui.Pick.pick_element(msg='Выбери следующий элемент. По окончанию - нажми "Esc"!')		
			try:
				element = doc.GetElement(selectedReference.id)	
				if systemName:
					elementType = doc.GetElement(element.GetTypeId())			
					elementSystemName = elementType.get_Parameter(guidSystemName).AsString()
					if elementSystemName == None:
						elementSystemName = element.get_Parameter(guidSystemName).AsString()
				startNumber += step	
				if systemName:
					newName = preff + elementSystemName + suff + str(startNumber)
				else:
					newName = preff + suff + str(startNumber)
				if str(Param.GetDefinition().ParameterType) == 'Integer' or str(Param.GetDefinition().ParameterType) == 'Double':
					with db.Transaction("pyKPLN_Нумерация объектов: "):
						try:
							element.get_Parameter(Param.GuidValue).Set(int(newName))
						except:
							ui.forms.Alert('Параметр числовой!')
				elif str(Param.GetDefinition().ParameterType) == 'Text':
					with db.Transaction("pyKPLN_Нумерация объектов: "):
						element.get_Parameter(Param.GuidValue).Set(newName)
				else:
					flag = False
			except:	
				output.print_md("**Некорректный элемент с id:** {} \n\n"
								"Либо не назначен параметр экземпляра 'КП_О_Марка' и параметр типа/экземпляра 'КП_О_Позиция', либо был выбран ложный элемент.".format(output.linkify(element.Id)))	
		except: flag = False



#calsses
class XamlApp(WPFWindow): 
	def __init__(self, xaml_file_name):
		WPFWindow.__init__(self, xaml_file_name)
		
		
	def select_step(self, sender, e):
		step = int(self.step_xaml.Text)
		flag = True
		customizable_event.raise_event(main_func, step, flag, Param)

if str(Param.GetDefinition().ParameterType) == 'Integer' or str(Param.GetDefinition().ParameterType) == 'Double':
	text1 = ''
	syf = False
else:
	text1 = ".1.1."
	syf = True

# form
try:
	components = [ui.forms.flexform.Label("Узел ввода данных"),              
				ui.forms.flexform.Label("Стартовый номер:"),			  
				ui.forms.flexform.TextBox("startNumber", Text="1"),			  
				ui.forms.flexform.Label("Приставка (если нет - оставь пустым):"),
				ui.forms.flexform.TextBox("preff"),			 
				ui.forms.flexform.CheckBox('systemName', 'Имя датчика?', default=syf),
				ui.forms.flexform.Label("Суффикс:"),
				ui.forms.flexform.TextBox("suff", Text=text1),				
				ui.forms.flexform.Separator(),
				ui.forms.flexform.Button("Запуск")]
	form = ui.forms.FlexForm("Задание порядкового номера для элемента схем СС", components)
	form.ShowDialog()
	startNumber = int(form.values["startNumber"]) - 1
	preff = form.values["preff"]
	systemName = form.values["systemName"]
	suff = form.values["suff"]
except:
	script.exit()



#main code
#getting info logger about user
log_name = "ЭОМ_СС_" + str(__title__)
InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)	
#main part of code	
customizable_event = CustomizableEvent()

XamlApp("change_number.xaml").Show()

