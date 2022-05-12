# -*- coding: utf-8 -*-

__title__ = "Создание\nрабочих наборов"
__author__ = 'Tima Kutsko'
__doc__ = "Создание рабочих наборов для подгруженных моделей"


from pyrevit import revit
from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkInstance,\
                              BuiltInParameter, ImportInstance, Workset
from pyrevit import script
from pyrevit import forms
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger


doc = revit.doc
output = script.get_output()


# Classes
class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value


# Functions
def create_check_boxes_by_name(elements):
    # Create check boxes for elements if they have Name property.
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements,
                                                                  key=lambda x:
                                                                  x.Name)]
    elementsCheckboxes = forms.\
        SelectFromList.\
        show(
            elements_options,
            multiselect=True,
            title='Выбери подгруженные модели',
            width=500,
            button_name='Выбрать'
        )
    return elementsCheckboxes


# Main code
linkModelInstances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
linkModels_checkboxes = create_check_boxes_by_name(linkModelInstances)


if linkModels_checkboxes:
    # getting info logger about user
    log_name = "Связи_Создание рабочих наборов"
    InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
    # main part of code
    linkModels = [c.value for c in linkModels_checkboxes if c.state is True]
    for element in linkModels:
        if isinstance(element, RevitLinkInstance):
            linkedModelName = element.Name.split(':')[0][0:-5]
        elif isinstance(element, ImportInstance):
            linkedModelName = element.\
                                Parameter[BuiltInParameter.
                                          IMPORT_SYMBOL_NAME].\
                                AsString()
        linkedName = "00_" + linkedModelName
        if linkedName:
            if not revit.doc.IsWorkshared\
                    and revit.doc.CanEnableWorksharing:
                revit.doc.EnableWorksharing('О_Оси и уровни',
                                            '_Модель')
                output.print_md(
                    "**Рабочие наборы для элементов проекта - созданы!**"
                )
            worksetParam = \
                element.Parameter[BuiltInParameter.ELEM_PARTITION_PARAM]
            if not worksetParam.IsReadOnly:
                with revit.Transaction('pyKPLN_Создать рабочие наборы'):
                    newWs = Workset.Create(revit.doc, linkedName)
                    worksetParam.Set(newWs.Id.IntegerValue)
                    output.print_md(
                        "Рабочий набор для связи **{0} (id: {1})** - создан!".
                        format(linkedModelName, output.linkify(element.Id))
                        )
            else:
                output.print_md(
                    "Связь **{}** - вложенная. Отдельный раб. набор-не нужен!".
                    format(element.Name.split(':')[1][0:-5])
                )
else:
    forms.alert('Нужно выбрать хотя бы одну связанную модель!')
