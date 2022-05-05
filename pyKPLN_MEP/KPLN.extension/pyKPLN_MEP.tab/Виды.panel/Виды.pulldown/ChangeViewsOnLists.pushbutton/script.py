# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import ViewSheet, Viewport, ViewSchedule,\
                              FilteredElementCollector, BuiltInCategory,\
                              ScheduleSheetInstance, BuiltInParameter, XYZ
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons
from pyrevit import revit, forms, DB, output, script
import clr
clr.AddReference('RevitAPI')


# Classes
class CheckBoxOption:
    def __init__(self, name, value):
        self.name = name
        self.value = value


# Functions
def createCBO(elements, viewTypeStr):
    # Create check boxes for elements if they have Name property.
    elements_options = list()
    for e in sorted(elements, key=lambda e: e.Name):
        try:
            catName = e.\
                get_Parameter(BuiltInParameter.VIEW_FAMILY).\
                AsString()
        except Exception:
            catName = e.\
                get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).\
                AsValueString()
        elName = e.Name
        fullName = catName + ": " + elName
        elements_options.append(CheckBoxOption(fullName, e))

    elements_checkboxes = forms.\
        SelectFromList.\
        show(elements_options,
             multiselect=False,
             title='Выбери вид для замены вида - {}'.format(viewTypeStr),
             width=750,
             button_name='Выбрать',
             )
    return elements_checkboxes


doc = revit.doc
currentView = doc.ActiveView
output = output.get_output()
if not isinstance(currentView, ViewSheet):
    forms.alert('Активным окном - должен быть лист,\n'
                'на котором нужно произвести замену',
                exitscript=True
                )
else:
    currentSheet = currentView


selEl = revit.pick_element("KPLN: Выбери вид для копирования")
with DB.Transaction(doc, 'KPLN_Замена видов') as t:
    try:
        t.Start()
        if isinstance(selEl, Viewport):
            # Работа с видами
            vpId = selEl.ViewId
            vpCat = doc.GetElement(vpId)
            vpCenter = selEl.GetBoxCenter()
            # Чистка от шаблонов видов
            viewsInst = filter(lambda x: not(x.IsTemplate),
                                FilteredElementCollector(doc).
                                OfCategory(BuiltInCategory.OST_Views)
                                )
            vpName = selEl.\
                get_Parameter(BuiltInParameter.VIEW_FAMILY).\
                AsString() + ": " +\
                vpCat.Name
            viewsCBO = createCBO(viewsInst, vpName)
            if viewsCBO:
                curView = viewsCBO.value
                nvp = Viewport.Create(doc,
                                        currentSheet.Id,
                                        curView.Id,
                                        vpCenter
                                        )
                currentSheet.DeleteViewport(selEl)
            else:
                script.exit()

        elif isinstance(selEl, ScheduleSheetInstance):
            # Работа со спецификациями
            # В основной надписи есть спецификация.
            # Этот шаг её отсеивает
            schedInst = filter(lambda x: "<" not in x.Name,
                                FilteredElementCollector(doc).
                                OfClass(ViewSchedule)
                                )
            schedName = selEl.\
                get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).\
                AsValueString() + ": " + \
                selEl.Name
            schedsCBO = createCBO(schedInst, schedName)
            if schedsCBO:
                curSched = schedsCBO.value
                selElBB_Max = selEl.get_BoundingBox(doc.ActiveView).Max
                nssi = ScheduleSheetInstance.Create(
                                                doc,
                                                currentSheet.Id,
                                                curSched.Id,
                                                selEl.Point
                                                )
                doc.Delete(selEl.Id)
                # Есть проблемы с положением спеки на листе.
                # Ниже - корректировка
                falseNssiPnt = nssi.Point
                nssiBB_Max = nssi.get_BoundingBox(doc.ActiveView).Max
                dist_X = selElBB_Max.X - nssiBB_Max.X
                dist_Y = selElBB_Max.Y - nssiBB_Max.Y
                trueNssiPnt = XYZ((falseNssiPnt.X + dist_X), (falseNssiPnt.Y + dist_Y), 0)
                nssi.Point = trueNssiPnt
            else:
                script.exit()

        else:
            TaskDialog.Show("Ой-ёй....",
                            "Можно выбирать только видовые окна",
                            TaskDialogCommonButtons.Close)
        t.Commit()

    except Exception as exc:
        t.RollBack()
        if "cannot be added to the ViewSheet" in str(exc):
            alreadyAddedList = curView.\
                get_Parameter(BuiltInParameter.VIEW_REFERENCING_SHEET).\
                AsString()
            vpViewName = curView.\
                get_Parameter(BuiltInParameter.VIEW_NAME).\
                AsString()
            alert = "Данный вид Имя: {} - уже размещен на листе Имя: {}".\
                    format(vpViewName, alreadyAddedList)
            TaskDialog.Show("Ой-ёй....", alert, TaskDialogCommonButtons.Close)
        else:
            print(exc)
