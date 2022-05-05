# -*- coding: utf-8 -*-

__title__ = "Проверка размеров"
__author__ = 'Tima Kutsko'
__doc__ = "Проверка наличия ошибок при нанесении размеров"

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
from rpw import revit, db
from pyrevit import forms, script


output = script.get_output()
doc = revit.doc
docDems = FilteredElementCollector(doc).\
    OfCategory(BuiltInCategory.OST_Dimensions).\
    WhereElementIsNotElementType().\
    ToElements()

for currentDem in docDems:
    overrideValue = currentDem.ValueOverride
    if overrideValue:
        currentViewId = currentDem.OwnerViewId
        output.print_md("{}".format(output.linkify(currentViewId)))


