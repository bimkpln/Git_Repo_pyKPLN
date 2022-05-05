#coding: utf-8

__title__ = "Перенос данных спецификации"
__author__ = 'Tima Kutsko'
__doc__ = "Перенос данных расчета воздухообмена в пространства"



import Autodesk
from Autodesk.Revit.DB import *
from rpw import revit, db, ui
from pyrevit import script
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger



# main script
doc = revit.doc
selectSchedule = doc.ActiveView
tableRevitElems = FilteredElementCollector(doc, selectSchedule.Id).ToElements()
table_data = selectSchedule.GetTableData()
table_columns = table_data.GetSectionData(SectionType.Body).NumberOfColumns
table_rows = table_data.GetSectionData(SectionType.Body).NumberOfRows

# input form
columsHeaderDict = dict()
for i in range(0, table_columns):
    columsHeaderDict[selectSchedule.GetCellText(SectionType.Body, 0, i)] = i

components = [ui.forms.flexform.Label("Сопоставь столбцы переноса данных"),
            ui.forms.flexform.Label("1. Приток расчетный:"),
            ui.forms.flexform.ComboBox("calc_supply_air_index", columsHeaderDict),
            ui.forms.flexform.Label("2. Вытяжка расчетная:"),
            ui.forms.flexform.ComboBox("calc_exhaust_air_index", columsHeaderDict),
            ui.forms.flexform.Label("3. ID пространства (может быть пустой):"),
            ui.forms.flexform.ComboBox("space_id_index", columsHeaderDict),
            ui.forms.flexform.Separator(),
            ui.forms.flexform.Button("Запуск")]
form = ui.forms.FlexForm("Сопоставь столбцы переноса данных", components)
form.ShowDialog()
calc_exhaust_air_index = form.values["calc_exhaust_air_index"]
calc_supply_air_index = form.values["calc_supply_air_index"]
space_id_index = form.values["space_id_index"]

#working with data
#getting info logger about user
log_name = "Пространства и помещения_" + str(__title__)
InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
#set space ID to special parameter
with db.Transaction("pyKPLN: Перенос ID прастранств в параметр"):
    space_id_param = selectSchedule.GetCellText(SectionType.Body, 0, space_id_index)
    for current_element in tableRevitElems:
        current_element.LookupParameter(space_id_param).Set(str(current_element.Id))

# set calculate exhaust/supply airflow data to the space
with db.Transaction("pyKPLN: Перенос данных в спецификации"):
    for i in range(1,table_rows):
        calc_supply_air_data = selectSchedule.GetCellText(SectionType.Body, i, calc_supply_air_index)
        calc_exhaust_air_data = selectSchedule.GetCellText(SectionType.Body, i, calc_exhaust_air_index)
        space_id = selectSchedule.GetCellText(SectionType.Body, i, space_id_index)
        try:
            doc.GetElement(ElementId(int(space_id))).get_Parameter(BuiltInParameter.ROOM_DESIGN_SUPPLY_AIRFLOW_PARAM).SetValueString(str(calc_supply_air_data))
            doc.GetElement(ElementId(int(space_id))).get_Parameter(BuiltInParameter.ROOM_DESIGN_EXHAUST_AIRFLOW_PARAM).SetValueString(str(calc_exhaust_air_data))
        except:
            pass
