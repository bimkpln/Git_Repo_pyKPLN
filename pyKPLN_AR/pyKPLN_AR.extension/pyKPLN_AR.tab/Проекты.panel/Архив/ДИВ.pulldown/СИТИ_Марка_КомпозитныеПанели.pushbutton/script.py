# coding: utf-8

__title__ = "СИТИ_Марка - Композитные панели"
__author__ = 'Tima Kutsko'
__doc__ = '''Запполняет номер панелей (сквозной на весь проект).\n
             Формат [префикс].[порядковый номер].\n
             Условия префикса:\n
             8 - 1018(Желтая),\n
             9 - 7026(Серая),\n
             10 - 3001(Красная),\n
             11 - 4008(Фиолетовая)'''


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,\
                              BuiltInParameter, ParameterValueProvider,\
                              FilterStringBeginsWith, FilterStringRule,\
                              ElementId, ElementParameterFilter, Options,\
                              RevitLinkInstance, ViewDetailLevel
from rpw import revit, db
from pyrevit import script
from System import Guid
import clr
from rpw import revit, db, ui
from rpw.ui.forms import*
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPI')


class MainDict:
    def __init__(self):
        self.elemDict = dict()

    def accumElems(self, elem, area, apprLenght):
        self.elem = elem
        self.area = area
        self.apprLenght = apprLenght
        try:
            self.color = self.elem.Symbol.get_Parameter(colParam).AsString()
        except AttributeError:
            self.color = 'RAL 7026'
        try:
            self.elemDict[self.color,
                          self.area,
                          self.apprLenght]\
                          +=\
                          [self.elem.Id]
        except KeyError:
            self.elemDict[self.color,
                          self.area,
                          self.apprLenght]\
                          =\
                          [self.elem.Id]


def getGreaterFace(facesArray):
    greaterArea = 0.5
    greaterApproxLenght = 0.5
    greateFace = None
    for face in facesArray:
        faceArea = face.Area
        if greateFace is None or\
                faceArea > greaterArea:
            greaterArea = faceArea
            greateFace = face
    if greateFace:
        edgesArrayArray = greateFace.EdgeLoops
        for edgeArray in edgesArrayArray:
            for edge in edgeArray:
                apprLenght = edge.ApproximateLength
                if apprLenght > greaterApproxLenght:
                    greaterApproxLenght = apprLenght
    # else:
    #     commands= [CommandLink('Да', return_value='true'),
    #         CommandLink('Нет', return_value='false')]
    #     dialog = TaskDialog('Элемент с 0 площадью. Удалить элемент из модели?',
    #                 commands=commands,
    #                     show_close=True)
    #     dialog_out = dialog.show()
    #     if dialog_out == 'true':
    #         with db.Transaction('Удаление элемента с 0 площадью'):
    #             doc.Delete(curItem.Id)

    return greaterArea, greaterApproxLenght


def numbSetter(ralTXT, indexName, numb=0):
    for keys, values in sortedDict.items():
        if not keys[0]:
            ui.forms.Alert("Параметр КП_О_Цвет не заполнен!", title="Ошибка")
            script.exit()
        if ralTXT in keys[0]:
            numb += 1
            for idL in values:
                elem = doc.GetElement(idL)
                if elem:
                    trueValue = indexName + str(numb)
                    try:
                        elem.\
                            get_Parameter(setParam).\
                            Set(trueValue)
                    except AttributeError:
                        pass


def elemGetter(origDoc, builtInCat, filterRule):
    try:
        for link in origDoc:
            linkDoc = link.GetLinkDocument()
            if linkDoc:
                modColl.extend(FilteredElementCollector(linkDoc).
                               OfCategory(builtInCat).
                               WherePasses(filterRule).
                               WhereElementIsNotElementType().
                               ToElements())
    except TypeError:
        modColl.extend(FilteredElementCollector(origDoc).
                       OfCategory(builtInCat).
                       WherePasses(filterRule).
                       WhereElementIsNotElementType().
                       ToElements())


# main code
output = script.get_output()
mainDict = MainDict()
colParam = Guid('0d78bf19-19ff-471a-91b7-87e33879f0a7')  # КП_О_Цвет
setParam = Guid('97b1e635-50c2-43fb-9d04-4b2e587b06ed')  # СИТИ_Марка
doc = revit.doc
modColl = list()

# ВЫБОР ФАЙЛОВ АР
bip = BuiltInParameter.ELEM_TYPE_PARAM
provider = ParameterValueProvider(ElementId(bip))
evaluater = FilterStringBeginsWith()
fStrRule = FilterStringRule(provider, evaluater, "АР_ДИВ", False)
fRule = ElementParameterFilter(fStrRule)
linkModels = FilteredElementCollector(doc).\
             OfClass(RevitLinkInstance).\
             WherePasses(fRule).\
             WhereElementIsNotElementType().\
             ToElements()

# ФИЛЬТР ДЛЯ ЭЛЕМЕНТОВ ОБОЩЕННЫХ МОДЕЛЕЙ
bip = BuiltInParameter.ELEM_FAMILY_PARAM
provider = ParameterValueProvider(ElementId(bip))
evaluater = FilterStringBeginsWith()
fStrRule_125 = FilterStringRule(provider, evaluater, "125_", False)
fRule_125 = ElementParameterFilter(fStrRule_125)

# ФИЛЬТР ДЛЯ ЭЛЕМЕНТОВ ПАНЕЛЕЙ ВИТРАЖА
bip = BuiltInParameter.ELEM_TYPE_PARAM
provider = ParameterValueProvider(ElementId(bip))
fStrRule_13 = FilterStringRule(provider, evaluater, "13_НА(Ком-25)", False)
fRule_13 = ElementParameterFilter(fStrRule_13)

# ПОЛУЧЕНИЕ СПИСКА (LIST) ЭЛЕМЕНТОВ МОДЕЛИ
elemGetter(doc, BuiltInCategory.OST_GenericModel, fRule_125)
elemGetter(doc, BuiltInCategory.OST_CurtainWallPanels, fRule_13)
elemGetter(linkModels, BuiltInCategory.OST_GenericModel, fRule_125)
elemGetter(linkModels, BuiltInCategory.OST_CurtainWallPanels, fRule_13)
opts = Options()
opts.DetailLevel = ViewDetailLevel.Fine
opts.ComputeReferences = True

for curItem in modColl:
    # area = curItem.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble()
    # if area == 0 :
    #     commands= [CommandLink('Да', return_value='true'),
    #         CommandLink('Нет', return_value='false')]
    #     dialog = TaskDialog('Элемент с 0 площадью. Продолжить?',
    #                 commands=commands,
    #                     show_close=True)
    #     dialog_out = dialog.show()
    #     if dialog_out == 'true':
    #         continue
    #     if dialog_out == 'false':
    #         print(curItem.Id)
    #         ui.forms.Alert("Площадь = 0!!!", title="Ошибка")
    #         script.exit()

    geom = curItem.get_Geometry(opts)
    IEnum = geom.GetEnumerator()
    for enum in IEnum:
        for solid in enum.GetSymbolGeometry():
            faces = solid.Faces
            greaterFaceItems = getGreaterFace(faces)
            mainDict.accumElems(curItem,
                                greaterFaceItems[0],
                                greaterFaceItems[1],)
            break

sortedDict = dict()
sortedKeys = sorted(mainDict.elemDict, key=mainDict.elemDict.get)
for i in sortedKeys:
    sortedDict[i] = mainDict.elemDict[i]
# for i in sortedDict:
#     print(i[0])

with db.Transaction('pyKPLN: div'):
    numbSetter('1018', '8.')
    numbSetter('7026', '9.')
    numbSetter('3001', '10.')
    numbSetter('4008', '11.')
