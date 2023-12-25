#coding: utf-8


__title__ = "Декларации для машиномест"
__author__ = 'Tima Kutsko'
__doc__ = 'Автоматическая генерация видов, листов '\
    'для для создания деклараций.\n'\
    'Перед началом - выбери планы для копирования.\n'\
    'Скрипт снимет с них копии (вместе с марками), назначит спец. шаблон,'\
    ' создаст листы и разместит на них созданные виды.'


import math
import re
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import BuiltInParameter, ParameterFilterRuleFactory,\
    ParameterFilterElement, ViewDuplicateOption, ElementId,\
    ElementParameterFilter, FilteredElementCollector,\
    BuiltInCategory, FilterStringRule, FilterStringEquals,\
    ParameterValueProvider,ElementId, Transaction,\
    ViewPlan, BoundingBoxXYZ, XYZ, ViewSheet, FamilySymbol,\
    Viewport, OverrideGraphicSettings, Line, Curve, CurveLoop, Color,\
    FillPatternElement
from datetime import date
from System import Guid
from System.Collections.Generic import List
from rpw import ui
from pyrevit import script


class DecParkSpaces:
    """ Класс для создания коллекции парковочных мест, объедененных
        под одним планом
    """

    def __init__(self, parkSpace, view, separNumb, separNumbParam):
        self.ParkSpace = parkSpace
        self.View = view
        self.SeparNumb = separNumb
        self.SeparNumbParam = separNumbParam


class ViewsCreator:
    """ Класс для создание видов для декларации:
    """

    def __init__(self, doc, decParkSpaces):
        self.Doc = doc
        self.DecParkSpaces = decParkSpaces


    def __planCopier(self, parkSpace, sepNumbParam, viewTemplate, bugName):
        # Создание вида
        if "мото" in parkSpace.ParkSpace.Name.lower():
            famDefType = "М/Мт"
        else:
            famDefType = "М/Ав"
        try:
            newName = "KPLN_Декларация_{}_{}_{}_{}".format(
                famDefType,
                bugName,
                sepNumbParam.Definition.Name,
                sepNumbParam.AsString()
            )
        except:
            newName = "KPLN_Декларация_{}_{}_{}_{}_{}".format(
                famDefType,
                bugName,
                sepNumbParam.Name,
                sepNumbParam.AsString(),
                date.today().strftime("%d/%m/%y")
            )
        newViewId = parkSpace.View.Duplicate(ViewDuplicateOption.WithDetailing)
        newView = doc.GetElement(newViewId)
        newView.ViewTemplateId = viewTemplate.Id
        newView.get_Parameter(BuiltInParameter.VIEW_NAME).Set(newName)
        # Предварительно нужно включить границы подрезки (для активации ViewCropRegionShapeManager)
        parkSpace.View.CropBoxActive = True
        parkSpace.View.CropBoxVisible = True
        newView.CropBoxActive = True
        newView.CropBoxVisible = True
        self.Doc.Regenerate()
        # Примененеие подрезки
        oldViewCropManager = parkSpace.View.GetCropRegionShapeManager()
        newViewCropManager = newView.GetCropRegionShapeManager()
        # Добавляю подрезку
        newViewCropManager.SetCropShape(oldViewCropManager.GetCropShape()[0])
        # Применение подрезки аннотаций (минимальная подрезка, равная 3 мм)
        newView.get_Parameter(
            BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE
        ).Set(1)
        newViewCropManager.LeftAnnotationCropOffset = 1 / 304.1
        newViewCropManager.RightAnnotationCropOffset = 1 / 304.1
        newViewCropManager.TopAnnotationCropOffset = 1 / 304.1
        newViewCropManager.BottomAnnotationCropOffset = 1 / 304.1
        # Создание переопределителя видимости фильтров (интересует прозрачность)
        return newView


    def createViews(self, viewTemplateMain):
        with Transaction(self.Doc, 'KPLN_Декларация. Копировать виды') as t:

            t.Start()
            decViewsDataList = list()
            fillPatternColl = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
            solidFillPattern = filter(lambda fp: fp.GetFillPattern().IsSolidFill, fillPatternColl)[0]
            for currentPrakSpace in self.DecParkSpaces:
                separNumberParam = currentPrakSpace.SeparNumbParam
                # Копирование планов
                newMainView = self.__planCopier(
                    currentPrakSpace,
                    separNumberParam,
                    viewTemplateMain,
                    "План"
                )
                # Добавляю переоперделение элементов на виде
                userColorNumbList = userColorNumb.split(',')
                ogs = OverrideGraphicSettings()
                ogs.SetProjectionLineColor(
                    Color(
                        int(userColorNumbList[0]),
                        int(userColorNumbList[1]),
                        int(userColorNumbList[2])
                    )
                )
                ogs.SetSurfaceBackgroundPatternId(solidFillPattern.Id)
                newMainView.SetElementOverrides(currentPrakSpace.ParkSpace.Id, ogs)
                # Создание коллекции видов
                if "мото" in currentPrakSpace.ParkSpace.Name.lower():
                    famDefType = "М/Мт"
                else:
                    famDefType = "М/Ав"

                try:
                    newListName = "KPLN_Декларация_{}_{}_{}".format(
                        famDefType,
                        separNumberParam.Definition.Name,
                        separNumberParam.AsString()
                    )
                except:
                    newListName = "KPLN_Декларация_{}_{}_{}".format(
                        famDefType,
                        separNumberParam.Name,
                        separNumberParam.AsString()
                    )
                # Создание надписи на листе
                newViewsCollName = "{} - {}".format(
                    userPartName,
                    separNumberParam.AsString()
                )
                decViewData = DecViewsData(
                    newViewsCollName,
                    newListName,
                    newMainView,
                )
                decViewsDataList.append(decViewData)
            t.Commit()
            return decViewsDataList


class DecViewsData:
    """ Класс для создании коллекций видов для размещения на листе
    """

    def __init__(self, viewsCollName, listName, mainView):
        self.ViewsCollName = viewsCollName
        self.ListName = listName
        self.MainView = mainView


class ListCreator:
    """ Класс для создание листов для декларации
    """

    def __init__(self, doc, decViews):
        self.Doc = doc
        self.DecViews = decViews


    def __dropViews(self, newSheet, currentView, tBlockBBox):
        # Определяю центр основной надписи
        vpCenter = XYZ(
            tBlockBBox.Min.X / 2,
            tBlockBBox.Max.Y / 2,
            0
        )
        vp = Viewport.Create(self.Doc,
            newSheet.Id,
            currentView.Id,
            vpCenter
        )
        # Рунчая замена типа на "00_Без названия" из шаблона АР
        vp.ChangeTypeId(ElementId(2049736))

    def createLists(self, titleFamilySymbols):
        with Transaction(self.Doc, 'KPLN_Декларация. Создать листы') as t:
            t.Start()
            currentListNumber = userStartListNumber - 1
            for decView in self.DecViews:
                currentListNumber += 1
                # Определеяю положение основной надиси: Книжное или Альбомное
                dVBBox = decView.MainView.get_BoundingBox(self.Doc.GetElement(decView.MainView.Id))
                dVBBoxMaxZZero = XYZ(dVBBox.Max.X, dVBBox.Max.Y, 0)
                dVBBoxMinZZero = XYZ(dVBBox.Min.X, dVBBox.Min.Y, 0)
                dVMaxMinDiagDist = dVBBoxMaxZZero.DistanceTo(dVBBoxMinZZero)
                dVMinYPoint = XYZ(dVBBoxMaxZZero.X, dVBBoxMinZZero.Y, dVBBoxMaxZZero.Z)
                maxHeightDist = dVBBoxMaxZZero.DistanceTo(dVMinYPoint)
                # 0.785398 радиан - это 45 градусов
                if (math.asin(maxHeightDist / dVMaxMinDiagDist) * 180 / 3.14 >= 45):
                    tFS = filter(
                        lambda t: "А1К" in t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), titleFamilySymbols
                    )[0]
                else:
                    tFS = filter(
                        lambda t: "А1А" in t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), titleFamilySymbols
                    )[0]
                newSheet = ViewSheet.Create(self.Doc, tFS.Id)
                newSheet.Name = decView.ListName
                newSheet.SheetNumber = "DEC_{}".format(currentListNumber)
                # Заполняю параметр для маркировки листа (спец. метка)
                newSheet.get_Parameter(userListMarkParam).Set(decView.ViewsCollName)
                self.Doc.Regenerate()

                # Размещаю виды на листы
                sheetTBlock = FilteredElementCollector(self.Doc, newSheet.Id).\
                    OfCategory(BuiltInCategory.OST_TitleBlocks).\
                    WhereElementIsNotElementType().\
                    FirstElement()
                tBlockBBox = sheetTBlock.get_BoundingBox(
                    self.Doc.GetElement(newSheet.Id)
                )
                self.__dropViews(
                    newSheet,
                    decView.MainView,
                    tBlockBBox,
                )
            t.Commit()


def checkElement(element, *params):
    result = True
    for param in params:
        try:
            if not element.get_Parameter(param).HasValue:
                result = False
        except TypeError:
            if not element.LookupParameter(param.Definition.Name).HasValue:
                result = False
    return result


def passesColl(doc, builInParam, fStringRule, builInCat, strName, plan=False):
    vpr = ParameterValueProvider(ElementId(builInParam))
    eRule = FilterStringRule(vpr, fStringRule, strName, True)
    eFilter = ElementParameterFilter(eRule)
    if plan:
        elColl = FilteredElementCollector(doc, plan.Id).\
            OfCategory(builInCat).\
            WherePasses(eFilter)
    else:
        elColl = FilteredElementCollector(doc).\
            OfCategory(builInCat).\
            WherePasses(eFilter)
    return elColl

# Основная часть
# Получаю документ и выбранные планы для создания деклараций
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
selRefs = __revit__.ActiveUIDocument.Selection
output = script.get_output()
selPlans = list()
decParkSpacesColl = list()
for refId in selRefs.GetElementIds():
    element = doc.GetElement(refId)
    if isinstance(element, ViewPlan):
        selPlans.append(element)
if len(selPlans) == 0:
    ui.forms.Alert(
        'Предварительно выбери планы. Можно несколько',
        title="КПЛН_Декларации",
        header="Нужно выбрать хотя бы один план!",
        exit=True
    )
else:
    allViewTemplColl = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views)
    viewTemplDict = {
        v.Name: v for v in allViewTemplColl if v.IsTemplate
    }
    # Форма пользовательского ввода
    ComboBox = ui.forms.flexform.ComboBox
    Label = ui.forms.flexform.Label
    Button = ui.forms.flexform.Button
    CheckBox = ui.forms.flexform.CheckBox
    TextBox = ui.forms.flexform.TextBox
    Separator = ui.forms.flexform.Separator
    components = [Label("Узел ввода данных"),
        Label("1. Введи переменную марки на листе:"),
        TextBox('startName', Text="Место №"),
        Label("2. Введи стартовый номер листа:"),
        TextBox('startListNumb', Text="1"),
        Label("3. Введи номер цвета (через запятую):"),
        TextBox('colorNumb', Text="248,222,135"),
        Label("4. Введи применяемый шаблон для листа:"),
        ComboBox("vTemplate", viewTemplDict),
        Separator(),
        Button("Запуск")]
    inputForm = ui.forms.FlexForm("Декларации", components)
    inputForm.ShowDialog()
    if not inputForm.values:
        script.exit()
    # Данные для выбора
    userPhaseName = "АР_Проект"
    userTemplateViewMain = inputForm.values["vTemplate"]
    userColorNumb = inputForm.values["colorNumb"]
    userTitleName = "020_Основная надпись_ДДУ"
    userOffsetMain = 700 / 304.1
    userPartName = inputForm.values["startName"]
    userStartListNumber = int(inputForm.values["startListNumb"])
    # Объеденяющий параметр, например: Марка
    userSeparNumbParam = Guid("ae8ff999-1f22-4ed7-ad33-61503d85f0f4")
    # Параметр, для записи в метку листа
    userListMarkParam = Guid("e313f126-7e51-4a5d-a45a-7c6dfe02124a")

    # Поиск основной надписи
    vpr = ParameterValueProvider(ElementId(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM))
    eRule = FilterStringRule(vpr, FilterStringEquals(), userTitleName)
    eFilter = ElementParameterFilter(eRule)
    userTitleFamilySymbols = FilteredElementCollector(doc).OfClass(FamilySymbol).WherePasses(eFilter).ToElements()

    # Поиск помещений, с указанным пользователем именем типа
    for currentPlan in selPlans:
        parkSpacesColl = FilteredElementCollector(doc, currentPlan.Id).\
            OfCategory(BuiltInCategory.OST_Parking).\
            ToElements()
        # Создаю коллекцию DecParkSpaces
        separNumbsSet = set()
        for parkSpace in parkSpacesColl:
            cheker = checkElement(
                parkSpace,
                userSeparNumbParam
            )
            if not cheker:
                output.print_md("Не заполнены требуемые параметры (Комментарий).\
                    Ошибка в элементе {}\
                    **Внимание!!! Работа скрипта остановлена. Устрани ошибку!**".
                    format(output.linkify(parkSpace.Id))
                )
                raise
            separNumberParam = parkSpace.get_Parameter(userSeparNumbParam)

            # Создаю спец класс по группированию помещений
            decParkSpaces = DecParkSpaces(
                parkSpace,
                currentPlan,
                separNumberParam.AsString(),
                separNumberParam
            )
            decParkSpacesColl.append(decParkSpaces)

        # Создаю спец класс по подготовке к созданию планов
        try:
            decParkSpacesColl = sorted(
                decParkSpacesColl,
                key=lambda d: int(d.SeparNumb)
            )
        except:
            # Если номер состоит из цифры и текста
            decParkSpacesColl = sorted(
                decParkSpacesColl,
                key=lambda d: int((''.join(re.findall(r'\d', d.SeparNumb))))
            )

        if len(decParkSpacesColl) > 0:
            vCreator = ViewsCreator(
                doc,
                decParkSpacesColl
            )

            # Создаю планы, и получаю список видов на листы
            viewsColl = vCreator.createViews(
                userTemplateViewMain
            )

            # Создаю спец класс по подготовке к созданию листов
            lCreator = ListCreator(doc, viewsColl)
            # Создаю листы, и размещаю на них виды
            lCreator.createLists(userTitleFamilySymbols)
        else:
            ui.forms.Alert(
                'Под заданные параметры - не попало ни одного парковочное место',
                title="КПЛН_Декларации",
                header="Проверь корректность ввода.",
                exit=True
            )
