#coding: utf-8

__title__ = "Добавить ОП в семействе"
__author__ = 'Tima Kutsko'
__doc__ = "Добавление/замена общих параметров (ОП) из ФОП в семействе"


from Autodesk.Revit.DB import *
from Autodesk.Revit.ApplicationServices import Application
from rpw import revit, ui, db
from pyrevit import script, forms
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger


# variables
output = script.get_output()


# classes
class CheckBoxOption:
	def __init__(self, name, value, default_state=True):
		self.name = name
		self.value = value
		self.state = default_state


def create_check_boxes_by_name(elements):
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements, key=lambda x: x.Name)]
    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect = True,
                                                    title='Выбери уровни',
                                                    width=300,
                                                    button_name='Выбрать')
    return elements_checkboxes



#main code
doc = revit.doc
if doc.IsFamilyDocument:
	family_parameter_names = [p.Definition.Name for p in doc.FamilyManager.GetParameters()]
	family_parameter_GUIDs = [p.GUID for p in doc.FamilyManager.GetParameters() if p.IsShared]
else:
	ui.forms.Alert("", header="Ты находишься в окне проекта! Необходимо открыть семейство для редактирования", title="ОШИБКА!", exit=True)


app = doc.Application
shared_file = app.OpenSharedParameterFile()	
shared_file_groups_dict = {g.Name: g for g in shared_file.Groups}
components = [ui.forms.flexform.Label("Выбери группу ОП из ФОПа"),
			ui.forms.flexform.ComboBox("shared_file_groups", shared_file_groups_dict),                        
			ui.forms.flexform.Button("Выбрать")]
form = ui.forms.FlexForm("shared_file", components)
form.ShowDialog()
try:
	shared_file_current_group = form.values["shared_file_groups"]
except:
	script.exit()


parameters_checkbox = create_check_boxes_by_name(shared_file_current_group.Definitions)
if parameters_checkbox:
	selected_parameter_definitions = [p.value for p in parameters_checkbox if p.state == True]
	for current_parameter_definition in selected_parameter_definitions:
		#getting info logger about user
		log_name = "Координация_" + str(__title__)
		InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
		#main code
		parameter_name = current_parameter_definition.Name	
		parameter_GUID = current_parameter_definition.GUID
		builtInParameter_current_parameter = current_parameter_definition.ParameterGroup
		
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
				added_parameter = doc.FamilyManager.AddParameter(current_parameter_definition, builtInParameter_current_parameter, False)
				n = 0
				for current_type in family_types:					
					doc.FamilyManager.CurrentType = current_type
					doc.FamilyManager.Set(added_parameter, types_data_list[n])
					n += 1	

		elif parameter_name not in family_parameter_names:
			with db.Transaction("pyKPLN_Добавление ОП: {}".format(current_parameter_definition.Name)):			
				added_parameter = doc.FamilyManager.AddParameter(current_parameter_definition, builtInParameter_current_parameter, False)
				output.print_md("Параметр **{}** в семейство добавлен!".format(parameter_name))

		else:
			output.print_md("Параметр **{}** был добавлен в семейство ранее!".format(parameter_name))
else:
	script.exit()


