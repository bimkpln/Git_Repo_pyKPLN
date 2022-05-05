# coding: utf-8

__title__ = "СИТИ_Размер_Диаметр"
__author__ = 'Tima Kutsko'
__doc__ = "Заполняет параметр размера для элементов ИОС на активном виде"


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,\
                              BuiltInParameter
from rpw import revit, db
from pyrevit import script
import codecs
from System import Guid


# main code: input part
doc = revit.doc
output = script.get_output()
# СИТИ_Диаметр
guid_diameter_param = Guid("d03cddce-6d51-442c-80bd-5ee41ac3e55c")
# СИТИ_Размер
guid_size_param = Guid("a1e634ab-3018-4e25-8ebf-0c65af534fbe")
# СИТИ_Классификатор
guid_RBS_param = Guid("32bf9389-a9b8-4db4-8d67-2fce3844b607")
catList = [BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves,
           BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_DuctFitting,
           BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
           BuiltInCategory.OST_CableTrayFitting,
           BuiltInCategory.OST_ConduitFitting]
catPipeList = [BuiltInCategory.OST_PipeInsulations]
catDuctList = [BuiltInCategory.OST_DuctInsulations]
with db.Transaction('СИТИ_Размер_Диаметр'):
    for cat in catList:
        catColl = FilteredElementCollector(doc, doc.ActiveView.Id).\
            OfCategory(cat).WhereElementIsNotElementType().ToElements()
        for curEl in catColl:
            # Начало блока исключений
            try:
                if "OSTEC_UM" in curEl.Symbol.Family.Name:
                    continue
                elif curEl.\
                        Symbol.\
                        get_Parameter(guid_RBS_param).AsString() == "9999":
                    continue
            except:
                pass
            if curEl.get_Parameter(guid_RBS_param).AsString() == "9999":
                continue
            # Конец блока исключений
            try:
                superComponent = curEl.SuperComponent
            except:
                superComponent = None
            if superComponent is None:
                sizeAll = curEl.\
                    get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE).\
                    AsString()
                if "x" in sizeAll:
                    curEl.get_Parameter(guid_size_param).Set(sizeAll)
                elif "ш" in sizeAll:
                    sizeSplited = sizeAll.split("ш")
                    clearSize = ""
                    for sizePart in sizeSplited:
                        clearSize += sizePart
                    clearSize += " мм"
                    curEl.get_Parameter(guid_diameter_param).Set(clearSize)
                elif "ø" in sizeAll:
                    sizeSplited = sizeAll.split("ø")
                    clearSize = ""
                    for sizePart in sizeSplited[1:]:
                        clearSize += sizePart
                        clearSize += "-"
                    clearSize = clearSize[:-1]
                    curEl.get_Parameter(guid_diameter_param).Set(clearSize)
                else:
                    output.print_md('''Элемент {} - либо новый суффикс размера, 
                                    либо нет параметра'''.
                                    format(output.linkify(curEl.Id)))
    for cat in catPipeList:
        catColl = FilteredElementCollector(doc, doc.ActiveView.Id).\
            OfCategory(cat).WhereElementIsNotElementType().ToElements()
        for curEl in catColl:
            sizePipe = curEl.\
                get_Parameter(BuiltInParameter.RBS_PIPE_CALCULATED_SIZE).\
                AsString()
            if not sizePipe:
                output.print_md('''Элемент {} - не имеет размера'''.
                                format(output.linkify(curEl.Id)))
                continue
            if "ш" in sizePipe:
                sizeSplited = sizePipe.split("ш")
                clearSize = ""
                for sizePart in sizeSplited:
                    clearSize += sizePart
                curEl.get_Parameter(guid_diameter_param).Set(clearSize)
            elif "ø" in sizePipe:
                sizeSplited = sizePipe.split("ø")
                clearSize = ""
                for sizePart in sizeSplited[1:]:
                    clearSize += sizePart
                    clearSize += "-"
                clearSize = clearSize[:-1]
                curEl.get_Parameter(guid_diameter_param).Set(clearSize)
            else:
                output.print_md('''Элемент {} - либо новый осуффикс размера, 
                    либо нет параметра'''.
                    format(output.linkify(curEl.Id)))
    for cat in catDuctList:
        catColl = FilteredElementCollector(doc, doc.ActiveView.Id).\
            OfCategory(cat).WhereElementIsNotElementType().ToElements()
        for curEl in catColl:
            sizeDuct = curEl.\
                get_Parameter(BuiltInParameter.RBS_DUCT_CALCULATED_SIZE).\
                AsString()
            if "x" in sizeDuct:
                curEl.get_Parameter(guid_size_param).Set(sizeDuct)
            elif "ш" in sizeDuct:
                sizeSplited = sizeDuct.split("ш")
                clearSize = ""
                for sizePart in sizeSplited:
                    clearSize += sizePart
                curEl.get_Parameter(guid_diameter_param).Set(clearSize)
            elif "ø" in sizeDuct:
                sizeSplited = sizeDuct.split("ø")
                clearSize = ""
                for sizePart in sizeSplited[1:]:
                    clearSize += sizePart
                    clearSize += "-"
                clearSize = clearSize[:-1]
                curEl.get_Parameter(guid_diameter_param).Set(clearSize)
            else:
                output.print_md('''Элемент {} - либо новый осуффикс размера, 
                                либо нет параметра'''.
                                format(output.linkify(curEl.Id)))
