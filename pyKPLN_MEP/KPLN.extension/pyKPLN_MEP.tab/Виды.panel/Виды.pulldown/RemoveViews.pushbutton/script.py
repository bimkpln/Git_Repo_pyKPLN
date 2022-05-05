# -*- coding: utf-8 -*-

__title__ = "Удаление лишних видов"
__author__ = 'Tima Kutsko'
__doc__ = "Проверка модели на наличее видов (план этажа/потолка, 3d-вид, чертежный вид), не размещенных на листах, и удаление их."


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,\
                              BuiltInParameter, ScheduleSheetInstance, \
                              ViewSchedule
from rpw import revit, db
from pyrevit import script, forms


# classes
class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value
        if self.name == "Спецификации":
            self.value = BuiltInCategory.OST_Schedules


# definitions
# creating checkbox
def create_check_boxes_by_name(elements, all_schedules):
    elements_options = [CheckBoxOption(e.Name, e) for e in sorted(elements,
                                                                  key=lambda x:
                                                                  x.Name)]
    elements_options.extend([CheckBoxOption(schdl, schdl) for schdl
                            in all_schedules])

    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect=True,
                                                    title='Выбери вид, чтобы удалить.',
                                                    width=500,
                                                    button_name='УДАЛИТЬ!')
    return elements_checkboxes


# main code
doc = revit.doc
output = script.get_output()
views_list = list()
false_views_set = set()
all_views_in_doc = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()
for current_view in all_views_in_doc:
    flag = True
    if not current_view.IsTemplate:
        sheet_viewport = current_view.get_Parameter(BuiltInParameter.VIEW_SHEET_VIEWPORT_INFO).AsString()
        sheet_dependecy = current_view.get_Parameter(BuiltInParameter.VIEW_DEPENDENCY).AsString()
        if sheet_dependecy == "Основной":
            depended_view_ids = current_view.GetDependentViewIds()
            # cleaning from false elements
            for current_id in depended_view_ids:
                sheet_viewport_depended = doc.GetElement(current_id).get_Parameter(BuiltInParameter.VIEW_SHEET_VIEWPORT_INFO).AsString()
                if sheet_viewport_depended != "Не на листе":
                    flag = False
        if flag:
            if sheet_viewport == "Не на листе" \
                    and current_view.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString() != "Легенда":
                views_list.append(current_view)


if views_list:
    schedules = ["Спецификации (ВНИМАНИЕ: удаляются все не размещенные!)"]
    views_checkbox = create_check_boxes_by_name(views_list, schedules)
    if views_checkbox:
        # main part of code
        selected_views = [v.value for v in views_checkbox if v.state is True]
        with db.Transaction('KPLN_Удалить лишние виды'):
            for view in selected_views:
                count_items = 0

                # deleted schedules
                if view in schedules:
                    all_view_shdl_ID_set = set()
                    all_schdl_sht_inst = FilteredElementCollector(doc).OfClass(ScheduleSheetInstance).ToElements()
                    for current_schdl_sht_inst in all_schdl_sht_inst:
                        all_view_shdl_ID_set.add(current_schdl_sht_inst.ScheduleId)
                    all_schedules_in_doc = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
                    for current_schdl in all_schedules_in_doc:
                        schdl_type = current_schdl.get_Parameter(BuiltInParameter.SCHEDULE_TYPE_FOR_BROWSER).AsValueString()
                        if current_schdl.Id not in all_view_shdl_ID_set and\
                           "Ключ" not in schdl_type:
                            count_items += 1
                            doc.Delete(current_schdl.Id)

                # deleted views
                try:
                    count_items += 1
                    doc.Delete(view.Id)
                except Exception as exc:
                    count_items -= 1
                    try:
                        output.print_md('Не удаляется вид **{}**. Причина: {}'.format(view.Name, exc))
                    except:
                        pass
            output.print_md('Удалено **{}** видов'.format(count_items))
else:
    output.print_md("**Выбери хотябы один вид!!!**")