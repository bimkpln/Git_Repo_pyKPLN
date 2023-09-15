#coding: utf-8

__title__ = "Добавление ОП в семейство с преднастройкой"
__author__ = 'Kapustin Roman'
__doc__ = "Создание набора ОП из ФОП в семействе"


# from typing import DefaultDict
from Autodesk.Revit.DB import *
from Autodesk.Revit.ApplicationServices import Application
from rpw import revit, ui, db
from pyrevit import script, forms
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger
from rpw.ui.forms import*
from System import Enum, Guid
import os
FilePathPred = 'Z:\\pyRevit\\pyRevit_MEP_ОбщПараметры_Конфигурации\\'

# variables
output = script.get_output()
doc = revit.doc
if doc.IsFamilyDocument:
    family_parameter_names = [p.Definition.Name for p in doc.FamilyManager.GetParameters()]
    family_parameter_GUIDs = [p.GUID for p in doc.FamilyManager.GetParameters() if p.IsShared]
else:
    ui.forms.Alert("", header="Ты находишься в окне проекта! Необходимо открыть семейство для редактирования", title="ОШИБКА!", exit=True)
app = doc.Application
shared_file = app.OpenSharedParameterFile()	

commands= [
    CommandLink('Создать новый файл настроек', return_value='CommandLink_new'),
    CommandLink('Использовать предыдущие настройки', return_value='CommandLink_set')
]
dialog = TaskDialog(
    'Выберите вариант работы',
    commands=commands,
    show_close=True)
dialog_out = dialog.show()
class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.value = value
        self.state = default_state
def create_check_boxes_by_name(elements):
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements, key=lambda x: x.Name)]
    elements_checkboxes = forms.SelectFromList.show(
        elements_options,
        multiselect = True,
        title='Выбери уровни',
        width=300,
        button_name='Выбрать')
    return elements_checkboxes

files = os.listdir(FilePathPred)
curFiles = []
for file in files:
    if '.txt' in file:
        curFiles.append(file)
if dialog_out == 'CommandLink_new':
    NewFileName = TextInput('Введите название файла', default="New_options_op")
    shared_params_groups = []
    shared_params_definitions = []
    for i in shared_file.Groups:
        shared_params_groups.append(i)
    parameters_checkbox_group = create_check_boxes_by_name(i for i in shared_params_groups)
    for i in parameters_checkbox_group:
        for j in i.value.Definitions:
            shared_params_definitions.append(j)
    if len(parameters_checkbox_group)>0:
        parameters_checkbox_definitions = create_check_boxes_by_name(i for i in shared_params_definitions)
        ui.forms.Alert('Выберите параметры, которые будут параметрами типа')
        if len(parameters_checkbox_definitions) > 0:
            parameters_checkbox_definitions_type = create_check_boxes_by_name(i.value for i in parameters_checkbox_definitions) 
        else:
            parameters_checkbox_definitions_type = []
    parameters_checkbox_definitions_List = []
    parameters_checkbox_definitions_type_List = []
    for i in parameters_checkbox_definitions:
        flagFind = False
        parameters_checkbox_definitions_List.append(i.value.GUID)
        for j in parameters_checkbox_definitions_type:
            if i.value.Name == j.value.Name:
                flagFind = True
        if flagFind:
            parameters_checkbox_definitions_type_List.append(True)
        else:
            parameters_checkbox_definitions_type_List.append(False)
    FilePath = FilePathPred + NewFileName + '.txt'
    with open(FilePath, 'w') as out:
        out.write(str(parameters_checkbox_definitions_List)+'/'+str(parameters_checkbox_definitions_type_List))
    TypeParam = []
    if len(parameters_checkbox_definitions_type)>0:
        for i in parameters_checkbox_definitions_type:
            TypeParam.append(i.value.Name)
    for current_parameter_definition in parameters_checkbox_definitions:
        TypeFlag = True
        log_name = "Координация_" + str(__title__)
        InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
        #main code
        parameter_name = current_parameter_definition.value.Name
        parameter_GUID = current_parameter_definition.value.GUID
        if parameter_name in TypeParam:
            TypeFlag = False
        builtInParameter_current_parameter = current_parameter_definition.value.ParameterGroup
        
        if parameter_GUID in family_parameter_GUIDs and parameter_name not in family_parameter_names:
            family_types = doc.FamilyManager.Types
            family_parameters = doc.FamilyManager.Parameters
            for current_parameter in family_parameters:
                if current_parameter.IsShared:
                    if current_parameter.GUID == parameter_GUID:
                        false_family_parameter = current_parameter
            types_data_list = []
            builtInParameter_current_parameter = false_family_parameter.Definition.ParameterGroup
            for current_type in family_types: 
                if false_family_parameter.StorageType == StorageType.String:
                    false_parameter_data = current_type.AsString(false_family_parameter)
                elif false_family_parameter.StorageType == StorageType.Double:
                    false_parameter_data = current_type.AsDouble(false_family_parameter)
                elif false_family_parameter.StorageType == StorageType.Integer:
                    false_parameter_data = current_type.AsInteger(false_family_parameter)
                elif false_family_parameter.StorageType == StorageType.ElementId:
                    false_parameter_data = current_type.AsElementId(false_family_parameter)
                
                types_data_list.append(false_parameter_data)
            with db.Transaction("pyKPLN_Замена ОП: {}".format(false_family_parameter.Definition.Name)): 
                #output message
                output.print_md("Параметр **{}** в семействе заменен на **{}**. Данные перенесены в новый параметр".format(false_family_parameter.Definition.Name, parameter_name))
                #deleting false parameter
                doc.Delete(false_family_parameter.Id)
                #adding true parameter with data
                added_parameter = doc.FamilyManager.AddParameter(current_parameter_definition.value, builtInParameter_current_parameter, TypeFlag)
                n = 0
                for current_type in family_types:
                    doc.FamilyManager.CurrentType = current_type
                    doc.FamilyManager.Set(added_parameter, types_data_list[n])
                    n += 1	
        elif parameter_name not in family_parameter_names:
            with db.Transaction("pyKPLN_Добавление ОП: {}".format(current_parameter_definition.value.Name)):
                added_parameter = doc.FamilyManager.AddParameter(current_parameter_definition.value, builtInParameter_current_parameter, TypeFlag)
                if TypeFlag:
                    output.print_md("Параметр **{}** добавлен как параметр экземпляра!".format(parameter_name))
                else:
                    output.print_md("Параметр **{}** добавлен как параметр типа!".format(parameter_name))
        else:
            output.print_md("Параметр **{}** был добвлен в семейство ранее!".format(parameter_name))

elif dialog_out=='CommandLink_set':
    FileName = SelectFromList('Выберите имя файла настроек', curFiles)
    exFilePath = FilePathPred + FileName
    with open(exFilePath,'r') as out:
        settings = out.read().decode('utf-8').split('/')
    try:
        settingsParam = settings[0][1:-1].split(',')
        settingsType = settings[1][1:-1].split(',')
    except:
        ui.forms.Alert('Некорректный файл настроек!')
        script.exit()
    guidList = []
    for i in settingsParam:
        guidList.append(i[i.index('[')+1:i.index(']')])
    shared_params_groups = []
    shared_params_definitions = []
    parameters_checkbox_definitions = []
    for i in shared_file.Groups:
        shared_params_groups.append(i)
    for i in shared_params_groups: 
        shared_params_definitions.append(i.Definitions)
    for i in shared_params_definitions:
        for j in i:
            if str(j.GUID) in guidList:
                parameters_checkbox_definitions.append(j)
    iter2 = 0

    for current_parameter_definition in parameters_checkbox_definitions:
        log_name = "Координация_" + str(__title__)
        InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
        #main code
        parameter_name = current_parameter_definition.Name
        parameter_GUID = current_parameter_definition.GUID
        TypeFladInd = guidList.index(str(parameter_GUID))
        TypeFlag = str(settingsType[TypeFladInd])
        if 'true' in TypeFlag.lower():
            TypeFlag = False
        else:
            TypeFlag = True
        builtInParameter_current_parameter = current_parameter_definition.ParameterGroup
        iter2+=1
        if parameter_GUID in family_parameter_GUIDs and parameter_name not in family_parameter_names:
            family_types = doc.FamilyManager.Types
            family_parameters = doc.FamilyManager.Parameters
            for current_parameter in family_parameters:
                if current_parameter.IsShared:
                    if current_parameter.GUID == parameter_GUID:
                        false_family_parameter = current_parameter
            types_data_list = []
            builtInParameter_current_parameter = false_family_parameter.Definition.ParameterGroup
            for current_type in family_types: 
                if false_family_parameter.StorageType == StorageType.String:
                    false_parameter_data = current_type.AsString(false_family_parameter)
                elif false_family_parameter.StorageType == StorageType.Double:
                    false_parameter_data = current_type.AsDouble(false_family_parameter)
                elif false_family_parameter.StorageType == StorageType.Integer:
                    false_parameter_data = current_type.AsInteger(false_family_parameter)
                elif false_family_parameter.StorageType == StorageType.ElementId:
                    false_parameter_data = current_type.AsElementId(false_family_parameter)
                
                types_data_list.append(false_parameter_data)
            with db.Transaction("pyKPLN_Замена ОП: {}".format(false_family_parameter.Definition.Name)): 
                #output message
                output.print_md("Параметр **{}** в семействе заменен на **{}**. Данные перенесены в новый параметр".format(false_family_parameter.Definition.Name, parameter_name))
                #deleting false parameter
                doc.Delete(false_family_parameter.Id)
                #adding true parameter with data						
                added_parameter = doc.FamilyManager.AddParameter(current_parameter_definition, builtInParameter_current_parameter, TypeFlag)
                n = 0
                for current_type in family_types:					
                    doc.FamilyManager.CurrentType = current_type
                    doc.FamilyManager.Set(added_parameter, types_data_list[n])
                    n += 1	
        elif parameter_name not in family_parameter_names:
            with db.Transaction("pyKPLN_Добавление ОП: {}".format(current_parameter_definition.Name)):
                added_parameter = doc.FamilyManager.AddParameter(current_parameter_definition, builtInParameter_current_parameter, TypeFlag)
                if TypeFlag:
                    output.print_md("Параметр **{}** добавлен как параметр экземпляра!".format(parameter_name))
                else:
                    output.print_md("Параметр **{}** добавлен как параметр типа!".format(parameter_name))
        else:
            output.print_md("Параметр **{}** был добвлен в семейство ранее!".format(parameter_name))
        