# -*- coding: utf-8 -*-

__title__ = "Проверка мониторинга осей и уровней"
__author__ = 'Tima Kutsko'
__doc__ = "Проверка наличия мониторинга у осей и уровней"


from Autodesk.Revit.DB import *
from rpw import revit, db
from pyrevit import forms, script
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger


# Classes
class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value  
    def __nonzero__(self):
        return self.state    
    def __bool__(self):
        return self.state



# Functions
def create_check_boxes_by_name(elements):
    # Create check boxes for elements if they have Name property.   
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements, key=lambda x: x.Name) if not e.GetLinkDocument() is None]
    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect=True,
                                                    title='Выбери файлы, из которых были скопированы оси и уровни:',
                                                    width=600,
                                                    button_name='Запуск!')
    return elements_checkboxes



# Main code
output = script.get_output()
error_notMonitored_set = set()
error_falseElement_dict = dict()


doc = revit.doc
doc_gridsAndLevels = list()
doc_grids = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
doc_levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
doc_gridsAndLevels.extend([g for g in doc_grids])
doc_gridsAndLevels.extend([l for l in doc_levels])


forms.alert('ВАЖНО! Выбери файлы, из которых должен был быть осуществлен процесс копирования/мониторинга')    
links_name_set = set()
linkModelInstances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
linkModels_checkboxes = create_check_boxes_by_name(linkModelInstances)
if linkModels_checkboxes:
    #getting info logger about user
    log_name = "Элементы ИОС_" + str(__title__)
    InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
    #main part
    with db.Transaction('pyKPLN_Контроль мониторинга'):
        linkModels = [c.value for c in linkModels_checkboxes if c.state == True] 
        for link in linkModels:
            links_name_set.add(link.Name)
        for current_element in doc_gridsAndLevels:
            isMonitored = current_element.IsMonitoringLinkElement()            
            if isMonitored:
                monitoredLinks = current_element.GetMonitoredLinkElementIds()
                for monitored_element in monitoredLinks:
                    link_name = doc.GetElement(monitored_element).Name
                    if not link_name in links_name_set:
                        error_falseElement_dict[current_element] = link_name.split(":")[0]
            else:
                error_notMonitored_set.add(current_element)


    # Output part
    if len(error_falseElement_dict) > 0:
        output.print_md("**Следующие элементы замониторены из других файлов:**")
        for current_element_id, current_link_name in error_falseElement_dict.items():
            output.print_md("Имя элемента в проекте: **{0}**. Его id: {1}. Файл, с которым **ошибочно** осуществлен мониторинг: **{2}**".format(current_element_id.Name,
                                                                                                                                                output.linkify(current_element_id.Id),
                                                                                                                                                current_link_name))
        print("_" * 70)
    if len(error_notMonitored_set) > 0:
        output.print_md("**Следующие элементы НЕ замониторены:**")
        for current_error in error_notMonitored_set:
            output.print_md("**Имя элемента в проекте: {0}. Его id: {1}**".format(current_error.Name,
                                                                                output.linkify(current_error.Id)))
    if len(error_falseElement_dict) == 0 and len(error_notMonitored_set) == 0:
        forms.alert('Ошибок в проекте нет')
else:
	forms.alert('Нужно выбрать хотя бы одну связанную модель!')        




