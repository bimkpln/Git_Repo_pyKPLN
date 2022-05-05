#coding: utf-8

__title__ = "Заполнение расхода решётки"
__author__ = 'Tima Kutsko'
__doc__ = "Заполнение расходов в решетках из заданных расходов приточного/удаляемого воздуха в пространствах"



import Autodesk
from Autodesk.Revit.DB import *
from rpw import revit, db, ui
from pyrevit import script, forms
from System import Guid
from re import split
from math import ceil
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger


#definitions
def systems_counter(list_of_duct_terminal_ids):
    supply_airflow_count = 0
    exhaust_airflow_count = 0    
    for duct_terminal_id in list_of_duct_terminal_ids:        
        duct_terminal = doc.GetElement(duct_terminal_id)
        duct_terminal_system_classification = duct_terminal.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString() 
        duct_terminal_system_name = duct_terminal.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString()
        if duct_terminal_system_classification == 'Приточный воздух' and exception not in duct_terminal_system_name:
            supply_airflow_count += 1  
        elif duct_terminal_system_classification == 'Отработанный воздух' and exception not in duct_terminal_system_name:
            exhaust_airflow_count += 1    
    return supply_airflow_count, exhaust_airflow_count


def big_data(selected_elements):    
    phase = list(doc.Phases)[-1]
    for current_element in selected_elements:  
        located_space = current_element.Space[phase]        
        if located_space:
            if duct_terminal_dict.get(located_space.Id) is None:
                duct_terminal_dict[located_space.Id] = []
            duct_terminal_dict[located_space.Id] += [current_element.Id]
        else:            
            error_duct_terminals_spaces.append(current_element)
    for current_space_id, current_duct_terminal_id in duct_terminal_dict.items():            
        current_space = doc.GetElement(current_space_id)            
        supply_airflow = split(r'\ ', current_space.get_Parameter(BuiltInParameter.ROOM_DESIGN_SUPPLY_AIRFLOW_PARAM).AsValueString())[0]               
        exhaust_airflow = split(r'\ ', current_space.get_Parameter(BuiltInParameter.ROOM_DESIGN_EXHAUST_AIRFLOW_PARAM).AsValueString())[0]
        counter = systems_counter(current_duct_terminal_id)
        for current_element_id in current_duct_terminal_id:
            current_duct_terminal = doc.GetElement(current_element_id)
            duct_terminal_system_classification = current_duct_terminal.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString()
            duct_terminal_system_name = current_duct_terminal.get_Parameter(BuiltInParameter.RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM).AsString()
            if duct_terminal_system_classification == 'Приточный воздух' and exception not in duct_terminal_system_name:                        
                duct_supply_airflow = current_duct_terminal.get_Parameter(Guid("9b7d541b-cc04-4e41-949d-a0d6ed778a25")).SetValueString(str(ceil(int(supply_airflow) / int(counter[0]))))
            elif duct_terminal_system_classification == 'Отработанный воздух' and exception not in duct_terminal_system_name:                                             
                duct_exhaust_airflow = current_duct_terminal.get_Parameter(Guid("9b7d541b-cc04-4e41-949d-a0d6ed778a25")).SetValueString(str(ceil(int(exhaust_airflow) / int(counter[1]))))                              
            if duct_terminal_system_name == '':
                error_duct_terminals_systems.append(current_duct_terminal)             
            true_duct_airflow = split(r'\ ', current_duct_terminal.get_Parameter(Guid("9b7d541b-cc04-4e41-949d-a0d6ed778a25")).AsValueString())[0] 
            if true_duct_airflow == str(0):                   
                error_airflow_duct_terminals.append(current_duct_terminal)                
    return error_duct_terminals_spaces, error_airflow_duct_terminals, error_duct_terminals_systems


def error_check(index):
    for res in result[index]:
        output.print_md('Решетка с id: {}'.format(output.linkify(res.Id)))



#main data
doc = revit.doc
output = script.get_output()
error_duct_terminals_spaces = list()
error_duct_terminals_systems = list()
error_airflow_duct_terminals = list()
duct_terminal_dict = dict()


# input form
selected_option = forms.CommandSwitchWindow.show(['Весь проект', 'Активный вид'],
                                                    message='Что делать?')
if selected_option:
    components = [ui.forms.flexform.Label("Часть 'Сокращение для системы' для исключения"),
				ui.forms.flexform.TextBox("exception", Text="Противодым."),
				ui.forms.flexform.Separator(),
				ui.forms.flexform.Button("Запуск")]
    form = ui.forms.FlexForm("Задание порядкового номера для элемента схем СС", components)
    form.ShowDialog()
    exception = form.values["exception"]


# main script
#getting info logger about user
log_name = "Воздухообмен_" + str(__title__)
InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
#main part of code
with db.Transaction('pyKPLN: Заполнение расхода решётки'):
    if selected_option== 'Весь проект':        
        selected_duct_terminal = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()        
        result = big_data(selected_duct_terminal)

    elif selected_option== 'Активный вид':
        active_view = doc.ActiveView.Id        
        selected_duct_terminal = FilteredElementCollector(doc, active_view).OfCategory(BuiltInCategory.OST_DuctTerminal).WhereElementIsNotElementType().ToElements()
        result = big_data(selected_duct_terminal)


# output results
try:
    if len(result[0]) > 0:
        output.print_md('**Нижеперечисленные элементы не попадают в пространтсва:**')
        error_check(0)
    if len(result[1]) > 0:
        output.print_md('**Нижеперечисленные элементы имеют нулевой расход:**')
        error_check(1)
    if len(result[2]) > 0:
        output.print_md('**Нижеперечисленные элементы не имеют системы:**')
        error_check(2) 
except:
    script.exit()