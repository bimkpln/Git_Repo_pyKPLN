# -*- coding: utf-8 -*-
__title__ = "Мощность отоп.\nприборов"
__author__ = 'Tima Kutsko'
__doc__ = "Перенос данных из пространства в отопительные приборы разделов ОВ1"


import os
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,\
    ElementId, FilterStringContains, ParameterValueProvider, FilterStringRule,\
    BuiltInParameter, ElementParameterFilter, Transaction, BuiltInParameterGroup
from rpw import revit, db
from rpw.ui.forms.flexform import Label, TextBox, Separator, FlexForm, Button
from rpw.ui.forms import Alert
from pyrevit import script
from System import Guid

# input form
components = [Label("Узел ввода"),
              Label("Введи индивидуальную часть имени системы:"),
              TextBox("sys_name", Text="От"),
              Label("ВНИМАНИЕ:"),
              Label("КП_И_Отопительная мощность прибора -"),
              Label("должен быть параметром экземпляра"),
              Separator(),
              Button("Запуск")]
form = FlexForm("Мощность отоп. приборов", components)
form.ShowDialog()
try:
    sys_name = form.values["sys_name"]
except Exception:
    script.exit()


# main code
doc = revit.doc
app = doc.Application
output = script.get_output()
# for DIV
# heating_data_param = Guid("05eb9c5e-c5a5-4c2c-b3fa-ca673358703c")
# КП_О_Отопительная мощность приборов
heatingDataParam_old = Guid("feb9ff82-b0cb-42ea-9bb5-455719071482")
# КП_И_Тепловая мощность
heatingDataParam_new = Guid("be7d2b1b-1916-428f-87f0-d9ee8d4f1efe")
# ОМП_Перезаписать
revalueParam = Guid("831d3a27-e65b-48be-9d04-f69ebc496e35")

# Подгружаю параметры
comParamsFilePath = "X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Плагины_Общий.txt"
# Список параметров, которые нужно обработать (заполняется из ФОП)
specialParamsGUIDList = []
if os.path.exists(comParamsFilePath):
    try:
        # Создаю спец класс CategorySet и добавляю в него зависимости
        # (категории)
        catSetElements = app.Create.NewCategorySet()
        catSetElements.Insert(
            doc
            .Settings
            .Categories
            .get_Item(BuiltInCategory.OST_MechanicalEquipment))

        # Забираю все парамтеры проекта в список
        prjParamsNamesList = []
        paramBind = doc.ParameterBindings
        fIterator = paramBind.ForwardIterator()
        fIterator.Reset()
        while fIterator.MoveNext():
            prjParamsNamesList.append(fIterator.Key.Name)

        # Забираю все параметры из спец ФОПа
        app.SharedParametersFilename = comParamsFilePath
        sharedParamsFile = app.OpenSharedParameterFile()

        # Добавляю недостающие парамтеры в проект
        with Transaction(doc, 'pyKPLN: Добавить параметры') as t:
            t.Start()
            for defGroups in sharedParamsFile.Groups:
                if defGroups.Name == "ОВВК_Общие":
                    for extDef in defGroups.Definitions:
                        # Список параметров (GUID) из ФОП, которые используются
                        # для плагина
                        specialParamsGUIDList.append(extDef.GUID)

                        # Добавляю параметры (если они не были ранее загружены)
                        if extDef.Name not in prjParamsNamesList and extDef.Name == "ОМП_Перезаписать":
                            paramBind = doc.ParameterBindings
                            newIB = app.\
                                Create.\
                                NewInstanceBinding(catSetElements)
                            paramBind.Insert(
                                extDef,
                                newIB,
                                BuiltInParameterGroup.PG_DATA
                            )

                            # Разворачиваю проход по параметрам проекта
                            revFIterator = doc.\
                                ParameterBindings.\
                                ReverseIterator()
                            while revFIterator.MoveNext():
                                if extDef.Name == revFIterator.Key.Name:
                                    break
            t.Commit()
    except Exception as e:
        Alert(
            "Ошибка при загрузке параметров:\n[{}]".format(str(e)),
            title="Загрузчик параметров", header="Ошибка"
        )
else:
    Alert(
        "Файл общих параметров не найден:{}".format(comParamsFilePath),
        title="Загрузчик параметров",
        header="Ошибка"
    )


# get elements with liquid systems
all_mech_equipment = FilteredElementCollector(doc).\
    OfCategory(BuiltInCategory.OST_MechanicalEquipment)
provider = ParameterValueProvider(
    ElementId(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)
    )
evaluator = FilterStringContains()
filter_rule = FilterStringRule(provider, evaluator, sys_name, False)
param_filter = ElementParameterFilter(filter_rule)
all_liquid_mech_equipment = all_mech_equipment.\
    WherePasses(param_filter).\
    WhereElementIsNotElementType().\
    ToElements()

# working with data
phase = list(doc.Phases)[-1]
space_dict = dict()
error_dict = dict()
none_space_error_set = set()
for current_element in all_liquid_mech_equipment:
    loc_space = current_element.Space[phase]
    if loc_space is None:
        none_space_error_set.add(current_element.Id)
        continue
    try:
        space_dict[loc_space.Id] += [current_element.Id]
    except Exception:
        space_dict[loc_space.Id] = [current_element.Id]

# setting data to equipment
with db.Transaction('pyKPLN: Мощность отоп. приборов'):
    for space_id, element_ids in space_dict.items():
        # working with spaces
        space = doc.GetElement(space_id)
        space_heating_power_revit = space.\
            LookupParameter("Проектная отопительная нагрузка").\
            AsValueString()
        # for DIV
        # space_heating_power_revit = space.get_Parameter(Guid("05eb9c5e-c5a5-4c2c-b3fa-ca673358703c")).AsValueString()
        space_heating_power = int(space_heating_power_revit.split(" ")[0])
        # working with elements
        elements_in_space = len(space_dict[space_id])
        for element_id in element_ids:
            current_element = doc.GetElement(element_id)
            try:
                heatingParam = current_element.\
                    get_Parameter(heatingDataParam_new)
                if heatingParam is None:
                    heatingParam = current_element.\
                        get_Parameter(heatingDataParam_old)
                if heatingParam is None\
                        or heatingParam.IsReadOnly:
                    raise Exception("Parameter is not family instance parameter")
                
                isRevalueParam = current_element.get_Parameter(revalueParam)
                if isRevalueParam is not None and isRevalueParam.HasValue:
                    if isRevalueParam.AsInteger() != 1:
                        continue
                else:
                    set_heating_data = heatingParam.\
                        Set(space_heating_power/elements_in_space * 10.7639)
            except Exception as exc:
                try:
                    error_dict[element_id] += [exc]
                except Exception:
                    error_dict[element_id] = [exc]

# error elemtns output
for element_id, true_exc in error_dict.items():
    family_symbol_id = doc.GetElement(element_id).GetTypeId()
    output.print_md("Элемент **{}:** {}, Id: {}. **Ошибка**: {}".format(
        doc.GetElement(family_symbol_id).FamilyName,
        doc.GetElement(element_id).Name,
        output.linkify(element_id),
        true_exc)
        )
print("_" * 100)

# error none space elemtns output
output.print_md("**Элементы, для которых было не найдено пространство:**")
for element_id in none_space_error_set:
    family_symbol_id = doc.GetElement(element_id).GetTypeId()
    output.print_md("Элемент **{}:** {}, Id: {}.".format(
        doc.GetElement(family_symbol_id).FamilyName,
        doc.GetElement(element_id).Name,
        output.linkify(element_id))
        )