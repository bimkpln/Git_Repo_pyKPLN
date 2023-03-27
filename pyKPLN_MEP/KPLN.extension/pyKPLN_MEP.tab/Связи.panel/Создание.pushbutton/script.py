# -*- coding: utf-8 -*-

__title__ = 'Рабочие\n наборы'
__doc__ = '''Поместить связи в рабочие наборы'''
__version__ = 1.1
__highlight__ = 'updated'

from Autodesk.Revit import DB
from rpw.ui import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document

if not doc.IsWorkshared and doc.CanEnableWorksharing:
    commands= [forms.CommandLink('Активировать функцию совместной работы и создать рабочие наборы', return_value = True), 
    forms.CommandLink('Отмена', return_value = False)]

    dialog = forms.TaskDialog('В документе не включена функция совместной работы',
    title_prefix = False,
    content = 'В данном документе не включена функция совместной работы. Хотите активировать функцию совместной работы и создать рабочие наборы ? \
Это действие нельзя будет отменить.',
    commands = commands)
    value = dialog.show()

    if value:
        try:
            doc.EnableWorksharing('!!!_Оси и уровни', '!!!_Модель')
        except:
            forms.Alert('Документ не предназначен для совместной работы', title = __title__)
            script.exit()
    else:
        script.exit()
worksets = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset).ToWorksets()
worksets_dict = {workset.Name : workset.Id for workset in worksets}

rvtlink_collector = DB.FilteredElementCollector(doc).\
    OfClass(DB.RevitLinkInstance).\
    WhereElementIsNotElementType().\
    ToElements()

rvtlink_collector = [rvtlink for rvtlink in rvtlink_collector \
    if not ('.rvt' in rvtlink.Name.split(':')[0].lower() and '.rvt' in rvtlink.Name.split(':')[1].lower())]

prefix = forms.TextInput('Плагин: Создание рабочих наборов', '00_', 'Введите префикс для рабочих наборов: ', True)

transacion = DB.Transaction(doc, 'Плагин: Создание рабочих наборов')
transacion.Start()
for rvtlink in rvtlink_collector:
    rvt_workset = rvtlink.Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
    rvt_name = rvtlink.Name.split(':')[0][:-5]
    rvt_workset_symbol = doc.GetElement(rvtlink.GetTypeId()).Parameter[DB.BuiltInParameter.ELEM_PARTITION_PARAM]
    if prefix + rvt_name != rvt_workset.AsValueString() or prefix + rvt_name != rvt_workset_symbol.AsValueString():
        if not prefix + rvt_name in worksets_dict.keys():
            new_workset = DB.Workset.Create(doc, prefix + rvt_name)
            rvt_workset.Set(new_workset.Id.IntegerValue)
            rvt_workset_symbol.Set(new_workset.Id.IntegerValue)
            print("Создан рабочий набор: " + prefix + rvt_name)
        else:
            rvt_workset.Set(worksets_dict[prefix + rvt_name].IntegerValue)
            rvt_workset_symbol.Set(worksets_dict[prefix + rvt_name].IntegerValue)
            print("Связь " + rvt_name + " -->> " + prefix + rvt_name)
transacion.Commit()