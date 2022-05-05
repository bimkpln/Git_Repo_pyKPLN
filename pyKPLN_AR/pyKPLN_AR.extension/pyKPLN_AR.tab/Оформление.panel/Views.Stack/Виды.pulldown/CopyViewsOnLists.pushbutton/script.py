# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import ViewSheet, ElementId, Viewport,\
                              FilteredElementCollector, BuiltInCategory,\
                              CopyPasteOptions, ScheduleSheetInstance, View,\
                              FamilyInstance, BuiltInParameter, TextNote,\
                              ElementTransformUtils
from pyrevit import revit, forms, DB, output
import viewedit as ve
from System.Collections.Generic import List


doc = revit.doc
currentView = doc.ActiveView
output = output.get_output()
if not isinstance(currentView, ViewSheet):
    forms.alert('Активным окном - должен быть лист, который нужно скопировать',
                exitscript=True
                )
else:
    currentSheet = currentView


items = ['Моделируемые виды',
         'Экземпляры семейств на листе',
         'Спецификации',
         'Легенды',
         'Текст на листах']
ops = forms.SelectFromList.show(items,
                                multiselect=True,
                                title='Узел ввода',
                                button_name='Запуск')


currentSheetNumber = currentSheet.SheetNumber
currentSheetName = currentSheet.Name
TBlockInst = FilteredElementCollector(doc, currentView.Id).\
             OfCategory(BuiltInCategory.OST_TitleBlocks).\
             WhereElementIsNotElementType().\
             FirstElement()
TBlockFam = TBlockInst.Symbol
TBlockInstLoc = TBlockInst.Location.Point
sheetElements = FilteredElementCollector(doc, currentSheet.Id).ToElements()


with DB.Transaction(doc, 'KPLN_Копировать листы') as t:
    prefix = "автКопия_"
    suffix = "_автКопия"
    try:
        t.Start()
        newSheet = ViewSheet.Create(doc, ElementId.InvalidElementId)
        newSheet.SheetNumber = prefix + currentSheetNumber
        newSheet.Name = prefix + currentSheetName

        for el in sheetElements:
            if isinstance(el, FamilyInstance):
                if 'Экземпляры семейств на листе' in ops:
                    cpOpt = CopyPasteOptions()
                    IColEl = List[ElementId]()
                    IColEl.Add(el.Id)
                    ElementTransformUtils.CopyElements(currentView,
                                                       IColEl,
                                                       newSheet,
                                                       None,
                                                       cpOpt)
            if isinstance(el, TextNote):
                if 'Текст на листах' in ops:
                    cpOpt = CopyPasteOptions()
                    IColEl = List[ElementId]()
                    IColEl.Add(el.Id)
                    ElementTransformUtils.CopyElements(currentView,
                                                       IColEl,
                                                       newSheet,
                                                       None,
                                                       cpOpt)
            elif isinstance(el, Viewport):
                vpId = el.ViewId
                vpCat = doc.GetElement(vpId)
                vpCenter = el.GetBoxCenter()
                vpTypeId = el.GetTypeId()
                if type(vpCat) is View:
                    if 'Легенды' in ops:
                        # Работа с легендами.
                        nvp = Viewport.Create(doc,
                                              newSheet.Id,
                                              vpId,
                                              vpCenter)
                else:
                    if 'Моделируемые виды' in ops:
                        # Работа с моделируемыми видами
                        view = doc.GetElement(vpId)
                        name = view.\
                            get_Parameter(BuiltInParameter.VIEW_NAME).\
                            AsString()
                        templateID = view.ViewTemplateId
                        newView = ve.\
                            duplicate_plan_view_with_detailing(
                                                               view,
                                                               templateID
                                                               )
                        newViewParam = newView.\
                            get_Parameter(BuiltInParameter.VIEW_NAME).\
                            Set(name + suffix)
                        try:
                            nvp = Viewport.Create(doc,
                                                  newSheet.Id,
                                                  newView.Id,
                                                  vpCenter
                                                  )
                            nvp.ChangeTypeId(vpTypeId)
                        except Exception as exc:
                            output.print_md("Ошибка **{}** у ID:**{}**".
                                            format(exc, name, vpId)
                                            )
            elif isinstance(el, ScheduleSheetInstance):
                if 'Спецификации' in ops:
                    # Работа со спецификациями
                    try:
                        nvp = ScheduleSheetInstance.Create(
                                        doc,
                                        newSheet.Id,
                                        el.ScheduleId,
                                        el.Point
                                        )
                    except Exception as exc:
                        # В основной надписи есть спецификация.
                        # Этот шаг её отсеивает
                        if "ViewSchedule that can be added" not in str(exc):
                            output.\
                                print_md("Ошибка **{}** у ID: **{}**".
                                         format(exc, el.Id))

        t.Commit()

    except Exception as exc:
        t.RollBack()
        output.print_md("Ошибка высшего уровня **{}** у ID: **{}**".
                        format(exc, el.Id))
