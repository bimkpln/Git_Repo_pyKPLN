# -*- coding: utf-8 -*-

__title__ = "Декларации для помещений"
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
    ElementParameterFilter, FilteredElementCollector, ScheduleFieldId, \
    BuiltInCategory, FilterStringRule, FilterStringEquals,\
    ParameterValueProvider,ElementId, Transaction, ScheduleFilter,\
    ViewPlan, BoundingBoxXYZ, XYZ, ViewSheet, FamilySymbol,\
    Viewport, OverrideGraphicSettings, Line, Curve, CurveLoop,\
    ScheduleFilterType, ScheduleSheetInstance, ViewSchedule
from datetime import date
from System import Guid
from System.Collections.Generic import List
from rpw import ui
from pyrevit import script, forms


class DecRooms:
    """ Класс для создания коллекции помещений, объедененных
        под одним параметром номера квартиры/аренд. помещения и секции
    """

    def __init__(self, unionRoomList, separRoomNumb, unionNumber, view, roomTypeParam, separRoomNumbParam, unionRoomNumbParam, roomDepartmentParam):
        self.RoomsDataColl = unionRoomList
        self.SeparRoomNumb = separRoomNumb
        self.UnionRoomNumb = unionNumber
        self.View = view
        self.RoomTypeParam = roomTypeParam
        self.SeparRoomNumbParam = separRoomNumbParam
        self.UnionRoomNumbParam = unionRoomNumbParam
        self.RoomDepartmentParam = roomDepartmentParam


class ViewsCreator:
    """ Класс для создание видов для декларации:
    """

    def __init__(self, doc, decRooms, allRooms):
        self.Doc = doc
        self.DecRooms = decRooms
        self.AllRooms = allRooms

    def __cropExpand(self, rColl, offset):
        xmax = ymax = zmax = - float("inf")
        xmin = ymin = zmin = float("inf")
        for room in rColl:
            bBox = BoundingBoxXYZ()
            geoElement = room.ClosedShell
            roomBBox = geoElement.GetBoundingBox()
            if roomBBox.Min.X < xmin:
                xmin = roomBBox.Min.X
            if roomBBox.Min.Y < ymin:
                ymin = roomBBox.Min.Y
            if roomBBox.Min.Z < zmin:
                zmin = roomBBox.Min.Z
            if roomBBox.Max.X > xmax:
                xmax = roomBBox.Max.X
            if roomBBox.Max.Y > ymax:
                ymax = roomBBox.Max.Y
            if roomBBox.Max.Z > zmax:
                zmax = roomBBox.Max.Z
        # Только для РПЯ
        tempP1_1 = XYZ(xmin, ymin, 0)
        tempP1_2 = XYZ(xmin, ymax, 0)
        tempP2_1 = XYZ(xmin, ymin, 0)
        tempP2_2 = XYZ(xmax, ymin, 0)
        tempLine1 = Line.CreateBound(tempP1_1, tempP1_2)
        tempLine2 = Line.CreateBound(tempP2_1, tempP2_2)
        if tempLine1.Distance > tempLine2.Distance:
            bBox.Min = XYZ(xmin - offset*3, ymin - offset, zmin)
            bBox.Max = XYZ(xmax + offset*3, ymax + offset, zmax)
        else:
            bBox.Min = XYZ(xmin - offset, ymin - offset*3, zmin)
            bBox.Max = XYZ(xmax + offset, ymax + offset*3, zmax)
        
        return bBox


    # def __doorMarker(self, decRooms, view):
    #     phase = filter(lambda p: p.Name == userPhaseName, list(self.Doc.Phases))[0]
    #     roomNumb = decRooms.RoomNumb
    #     vpr = ParameterValueProvider(ElementId(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,))
    #     eRule = FilterStringRule(vpr, FilterStringContains(), userDoorPartName, True)
    #     eFilter = ElementParameterFilter(eRule)
    #     doorsColl = FilteredElementCollector(self.Doc, view.Id).\
    #         OfCategory(BuiltInCategory.OST_Doors).\
    #         WherePasses(eFilter).\
    #         ToElements()
    #     for d in doorsColl:
    #         try:
    #             innerRoomNumb = d.\
    #                 FromRoom[phase].\
    #                 get_Parameter(userSeparNumbParam).\
    #                 AsString()
    #         except:
    #             continue
    #         if roomNumb == innerRoomNumb:
    #             # Указываю двери спец. номер помещения
    #             d.get_Parameter(userDoorMarkParam).Set(roomNumb)
    #             # Создаю марку
    #             ref = Reference(d)
    #             point = d.Location.Point
    #             IndependentTag.Create(self.Doc, userTag.GetTypeId(), view.Id, ref, False, TagOrientation.Horizontal, point)


    def __createCrop(self, decCol, offset):
        currentSepar = decCol.UnionRoomNumbParam
        equalSectRoomsColl = filter(
            lambda r: r.get_Parameter(userUnionNumbParam).AsString() == currentSepar.AsString(),
            self.AllRooms)
        sectBBox = self.__cropExpand(equalSectRoomsColl, offset)
        return sectBBox

    def __viewFilterCreator(self, name, param, separator, data):
        try:
            paramName = param.Name
            filterName = "Dec_{} {} {} {}".format(
                name,
                paramName,
                separator,
                data
            )
        except:
            paramName = param.Definition.Name
            filterName = "Dec_{} {} {} {}".format(
                name,
                paramName,
                separator,
                data
            )
        if separator == "НРВ":
            fRule = ParameterFilterRuleFactory.\
                CreateNotEqualsRule(param.Id, data, False)
        elif separator == "РВ":
            fRule = ParameterFilterRuleFactory.\
                CreateEqualsRule(param.Id, data, False)
        filterRules = ElementParameterFilter(fRule)
        IfilterCat = List[ElementId]()
        IfilterCat.Add(ElementId(BuiltInCategory.OST_Rooms))
        newFilter = ParameterFilterElement.Create(
            self.Doc,
            filterName,
            IfilterCat,
            filterRules
        )
        return newFilter

    def __planCopier(self, decRooms, rNumber, rNumberParam, viewTemplate, currentBBox, bugName, depFilter):
        # Создание вида
        try:
            newName = "KPLN_Декларация_{}_{}_{}".format(
                bugName,
                rNumberParam.Name,
                rNumber
            )
        except:
            newName = "KPLN_Декларация_{}_{}_{}_{}".format(
                bugName,
                rNumberParam.Name,
                rNumber,
                date.today().strftime("%d/%m/%y")
            )
        newViewId = decRooms.View.Duplicate(ViewDuplicateOption.WithDetailing)
        newView = doc.GetElement(newViewId)
        newView.ViewTemplateId = viewTemplate.Id
        newView.get_Parameter(BuiltInParameter.VIEW_NAME).Set(newName)
        # Предварительно нужно включить границы подрезки (для активации ViewCropRegionShapeManager)
        newView.CropBoxActive = True
        newView.CropBoxVisible = True
        self.Doc.Regenerate()
        # Примененеие подрезки
        newViewCropManager = newView.GetCropRegionShapeManager()
        if isCrop:
            # Создаю кривые для обрезки вида (по баудингбоксу)
            curveIList = List[Curve]()
            p1 = XYZ(currentBBox.Max.X, currentBBox.Max.Y, 0)
            p2 = XYZ(currentBBox.Max.X, currentBBox.Min.Y, 0)
            p3 = XYZ(currentBBox.Min.X, currentBBox.Min.Y, 0)
            p4 = XYZ(currentBBox.Min.X, currentBBox.Max.Y, 0)
            curveIList.Add(Line.CreateBound(p1, p2))
            curveIList.Add(Line.CreateBound(p2, p3))
            curveIList.Add(Line.CreateBound(p3, p4))
            curveIList.Add(Line.CreateBound(p4, p1))
            curveLoop = CurveLoop.Create(curveIList)
            # Добавляю подрезку
            newViewCropManager.SetCropShape(curveLoop)
        else:
            # Предварительно нужно включить границы подрезки (для активации ViewCropRegionShapeManager)
            decRooms.View.CropBoxActive = True
            decRooms.View.CropBoxVisible = True
            self.Doc.Regenerate()

            oldViewCropManager = decRooms.View.GetCropRegionShapeManager()
            newViewCropManager.SetCropShape(oldViewCropManager.GetCropShape()[0])
            newView.CropBoxActive = False
            newView.CropBoxVisible = False
        # Применение подрезки аннотаций (минимальная подрезка, равная 3 мм)
        newView.get_Parameter(
            BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE
        ).Set(1)
        newViewCropManager.LeftAnnotationCropOffset = 1 / 304.1
        newViewCropManager.RightAnnotationCropOffset = 1 / 304.1
        newViewCropManager.TopAnnotationCropOffset = 1 / 304.1
        newViewCropManager.BottomAnnotationCropOffset = 1 / 304.1
        # Создание переопределителя видимости фильтров (интересует прозрачность)
        ogs = OverrideGraphicSettings()
        ogs.SetSurfaceTransparency(100)
        # Применение фильтрации
        newFilterNumb = self.__viewFilterCreator(
            bugName,
            rNumberParam,
            "НРВ",
            rNumber
        )
        newView.AddFilter(newFilterNumb.Id)
        newView.SetFilterOverrides(newFilterNumb.Id, ogs)
        #newView.SetFilterVisibility(newFilterNumb.Id, False)
        newView.AddFilter(depFilter.Id)
        newView.SetFilterVisibility(depFilter.Id, False)
        return newView

    def createViews(self, viewTemplateMain, offsetMain, userTemplateScheduleMain):
        with Transaction(self.Doc, 'KPLN_Декларация. Копировать виды') as t:

            t.Start()
            decViewsDataList = list()
            # Создаю фильтр на категорию помещения
            newFilterDep = self.__viewFilterCreator(
                decRooms.RoomDepartmentParam.AsString(),
                decRooms.RoomTypeParam,
                "НРВ",
                decRooms.RoomTypeParam.AsString()
            )
            for currentDecRooms in self.DecRooms:
                separRoomNumber = currentDecRooms.SeparRoomNumb
                separRoomNumberParam = self.Doc.GetElement(
                    currentDecRooms.
                    SeparRoomNumbParam.
                    Id
                )
                # Создание подрезок по выбранным параметрам (BoundingBoxXYZ)
                # если в этом есть необходимость
                if isCrop:
                    mainBBox = self.__createCrop(
                        currentDecRooms,
                        offsetMain
                    )
                else:
                    mainBBox = None
                # Копирование планов
                newMainView = self.__planCopier(
                    currentDecRooms,
                    separRoomNumber,
                    separRoomNumberParam,
                    viewTemplateMain,
                    mainBBox,
                    "План",
                    newFilterDep
                )

                # Создание спецификации
                # Создание новой спецификации с выбранной категорией
                newMainSchedule = ViewSchedule.CreateSchedule(
                    self.Doc,
                    ElementId(BuiltInCategory.OST_Rooms)
                    )
                # Установка имени для новой спецификации
                newMainSchedule.Name = "KPLN_Декларация_{}_{}".format(
                    currentDecRooms.RoomTypeParam.AsString(),
                    currentDecRooms.SeparRoomNumb
                )
                newMainSchedule.ViewTemplateId = userTemplateScheduleMain.Id
                # Создаю фильтр на номер помещения
                for schedulableField in newMainSchedule.Definition.GetSchedulableFields():
                    if "Клад" in currentDecRooms.RoomTypeParam.AsValueString():
                        paramFieldName = currentDecRooms.RoomTypeParam.Definition.Name
                        paramFilterData = currentDecRooms.RoomTypeParam.AsString()
                    else:
                        paramFieldName = currentDecRooms.SeparRoomNumbParam.Definition.Name
                        paramFilterData = currentDecRooms.SeparRoomNumb

                    if schedulableField.GetName(self.Doc).Equals(paramFieldName):
                        scheduleField = newMainSchedule.Definition.AddField(schedulableField)

                newFilter = ScheduleFilter(
                    # Не работает, т.к. параметр может 2 раза встречаться (по имени). Лучше вручную указать стартовый id - 0
                    # scheduleField.FieldId
                    ScheduleFieldId(0),
                    ScheduleFilterType.Equal,
                    str(paramFilterData)
                )
                newMainSchedule.Definition.AddFilter(newFilter)

                # Маркировка дверей
                #self.__doorMarker(currentDecRooms, newMainView)
                # Создание коллекции видов
                newListName = "KPLN_Декларация_{}_{}".format(
                    currentDecRooms.RoomDepartmentParam.AsString(),
                    separRoomNumber
                )
                # Создание надписи на листе
                newViewsCollName = "{} - {}".format(
                    currentDecRooms.RoomDepartmentParam.AsString(),
                    separRoomNumber
                )
                decViewData = DecViewsData(
                    currentDecRooms,
                    newViewsCollName,
                    newListName,
                    newMainView,
                    newMainSchedule
                )
                decViewsDataList.append(decViewData)
            t.Commit()
            return decViewsDataList


class DecViewsData:
    """ Класс для создании коллекций видов для размещения на листе
    """

    def __init__(self, decRooms, viewsCollName, listName, mainView, mainSchedule):
        self.DecRooms = decRooms
        self.ViewsCollName = viewsCollName
        self.ListName = listName
        self.MainView = mainView
        self.MainSchedule = mainSchedule


class ListCreator:
    """ Класс для создание листов для декларации
    """

    def __init__(self, doc, decViews):
        self.Doc = doc
        self.DecViews = decViews


    def __dropView(self, newSheet, currentView, vpCenter):
        
        vp = Viewport.Create(
            self.Doc,
            newSheet.Id,
            currentView.Id,
            vpCenter
        )
        # Рунчая замена типа на "00_Без названия" из шаблона АР
        vp.ChangeTypeId(ElementId(2049736))


    def __dropSchedule(self, newSheet, currentView, vpCenter):
        ScheduleSheetInstance.Create(
            self.Doc,
            newSheet.Id,
            currentView.Id,
            vpCenter
        )

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
                        lambda t: "А2К" in t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), titleFamilySymbols
                    )[0]
                else:
                    tFS = filter(
                        lambda t: "А2А" in t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), titleFamilySymbols
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

                # Определяю центр основной надписи
                #vpCenter = XYZ(
                #    tBlockBBox.Min.X / 2,
                #    tBlockBBox.Max.Y / 2,
                #    0
                #)
                self.__dropView(
                    newSheet,
                    decView.MainView,
                    XYZ(-0.95, 0.9, 0)
                )
                self.__dropSchedule(
                    newSheet,
                    decView.MainSchedule,
                    XYZ(-1.883, 0.427, 0)
                )
            t.Commit()

class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value

    def __nonzero__(self):
        return self.state

    def __bool__(self):
        return self.state


def create_check_boxes_by_name(elements):
    # Create check boxes for elements if they have Name property
    elements_options = [
        CheckBoxOption(e.Name, e) for e in sorted(elements, key=lambda x: x.Name) if not e.GetLinkDocument() is None
    ]
    elements_checkboxes = forms.SelectFromList.show(
        elements_options,
        multiselect=True,
        title='Выбери подгруженные модели, из которой брать помещения:',
        width=600,
        button_name='Запуск!'
    )
    return elements_checkboxes

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
    eRule = FilterStringRule(vpr, fStringRule, strName)
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
decRoomsColl = list()
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
    # Данные для выбора
    #linkModelInstances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    #linkModelsCheckBox = create_check_boxes_by_name(linkModelInstances)
    #if not linkModelsCheckBox:
    #    forms.alert('Нужно выбрать хотя бы одну связанную модель!')
    #    script.exit()

    allRoomsColl = list()
    #linkModels = [c.value for c in linkModelsCheckBox if c.state is True]
    #for link in linkModels:
    #    linkDoc = link.GetLinkDocument()
    #    allRoomsColl.append(FilteredElementCollector(linkDoc).OfCategory(BuiltInCategory.OST_Rooms))
    allRoomsColl.extend(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms))
    
    roomDepartmentDict = {
        r.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString(): r.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT) for r in allRoomsColl
    }
    paramDict = dict()
    for r in allRoomsColl:
        params = r.Parameters
        for p in params:
            if p.IsShared:
                paramDict[p.Definition.Name] = p.GUID
        break

    allViewTemplColl = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views)
    viewTemplDict = {
        v.Name: v for v in allViewTemplColl if v.IsTemplate
    }

    allScheduleTemplColl = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Schedules)
    scheduleTemplDict = {
        v.Name: v for v in allScheduleTemplColl if v.IsTemplate
    }
    # Форма пользовательского ввода
    ComboBox = ui.forms.flexform.ComboBox
    Label = ui.forms.flexform.Label
    Button = ui.forms.flexform.Button
    CheckBox = ui.forms.flexform.CheckBox
    TextBox = ui.forms.flexform.TextBox
    Separator = ui.forms.flexform.Separator
    components = [Label("Узел ввода данных"),
        Label("1. Выбери типа помещения (из параметра 'Назначение'):"),
        ComboBox("roomType", roomDepartmentDict),
        Label("2. Выбери параметр, для деления на секции:"),
        ComboBox("unionParam", paramDict),
        Label("3. Выбери параметр для разделения помещений:"),
        ComboBox("separParam", paramDict),
        Label("4. Введи стартовый номер листа:"),
        TextBox('startListNumb', Text="1"),
        Label("5. Введи применяемый шаблон для плана:"),
        ComboBox("vTemplate", viewTemplDict),
        Label("6. Введи применяемый шаблон для спеки:"),
        ComboBox("sTemplate", scheduleTemplDict),
        CheckBox('isNoneIgnore',
            'Игнорировать пом., с пустыми параметрами?',
            default=False
        ),
        CheckBox('isCrop',
            'Подрезать вид при делении на секции?',
            default=True
        ),
        Separator(),
        Button("Запуск")]
    inputForm = ui.forms.FlexForm("Декларации", components)
    inputForm.ShowDialog()
    if not inputForm.values:
        script.exit()

    # Отработка с учетом пользовательского ввода
    userRoomTypeData = inputForm.values["roomType"]
    userDoorTagName = "050_Марка входной двери_Декларация_(Дверь): Тип 1"
    userDoorPartName = "Входная"
    userPhaseName = "АР_Проект"
    userTemplateViewMain = inputForm.values["vTemplate"]
    userTemplateScheduleMain = inputForm.values["sTemplate"]
    userTitleName = "020_Основная надпись_ДДУ"
    userOffsetMain = 700 / 304.1
    userStartListNumber = int(inputForm.values["startListNumb"])
    # Объеденяющий параметр, например: сквозной номер квартиры в здании
    userSeparNumbParam = inputForm.values["separParam"]
    # Объеденяющий параметр, например: номер секции
    userUnionNumbParam = inputForm.values["unionParam"]
    # Параметр, для записи в метку листа
    userListMarkParam = Guid("e313f126-7e51-4a5d-a45a-7c6dfe02124a")
    # Параметр, для записи в метку двери
    userDoorMarkParam = Guid("a2985e5c-b28e-416a-acf6-7ab7e4ee6d86")
    # Для игнорирования помещений, в которых выбранные параметры не заполнены
    isNoneIgnore = inputForm.values["isNoneIgnore"]
    # Для включения функции подрезки под параметр секции
    isCrop = inputForm.values["isCrop"]

    # Поиск основной надписи
    vpr = ParameterValueProvider(ElementId(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM))
    eRule = FilterStringRule(vpr, FilterStringEquals(), userTitleName)
    eFilter = ElementParameterFilter(eRule)
    userTitleFamilySymbols = FilteredElementCollector(doc).OfClass(FamilySymbol).WherePasses(eFilter).ToElements()

    # Поиск марки для дверей
    userTag = passesColl(
        doc,
        BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
        FilterStringEquals(),
        BuiltInCategory.OST_DoorTags,
        userDoorTagName
    ).WhereElementIsNotElementType().FirstElement()

    # Поиск помещений, с указанным пользователем именем типа
    for currentPlan in selPlans:
        roomsColl = passesColl(
            doc,
            BuiltInParameter.ROOM_DEPARTMENT,
            FilterStringEquals(),
            BuiltInCategory.OST_Rooms,
            userRoomTypeData.AsString(),
            currentPlan
        ).ToElements()
        # Создаю коллекцию DecRooms
        separRoomNumbsSet = set()
        for room in roomsColl:
            cheker = checkElement(
                room,
                userSeparNumbParam,
                userUnionNumbParam,
                userRoomTypeData
            )
            if not cheker and not isNoneIgnore:
                output.print_md("Не заполнены требуемые параметры.\
                    Ошибка в элементе {}\
                    **Внимание!!! Работа скрипта остановлена. Устрани ошибку!**".
                    format(output.linkify(room.Id))
                )
                raise
            currentThroughRoomNumb = room.get_Parameter(userSeparNumbParam).AsString()
            separRoomNumbsSet.add(currentThroughRoomNumb)
        for separRoomNumb in separRoomNumbsSet:
            unionRoomList = list()
            equalThroughRoomsColl = filter(
                lambda r: r.get_Parameter(userSeparNumbParam).AsString() == separRoomNumb,
                roomsColl)
            for room in equalThroughRoomsColl:
                unionRoomList.append(room)
            unionNumberParam = unionRoomList[0].get_Parameter(userUnionNumbParam)
            separNumberParam = unionRoomList[0].get_Parameter(userSeparNumbParam)
            roomTypeParam = unionRoomList[0].LookupParameter(userRoomTypeData.Definition.Name)
            roomDepartmentParam = unionRoomList[0].get_Parameter(BuiltInParameter.ROOM_DEPARTMENT)

            # Создаю спец класс по группированию помещений
            decRooms = DecRooms(
                unionRoomList,
                separRoomNumb,
                unionNumberParam.AsString(),
                currentPlan,
                roomTypeParam,
                separNumberParam,
                unionNumberParam,
                roomDepartmentParam
            )
            decRoomsColl.append(decRooms)

        # Создаю спец класс по подготовке к созданию планов
        try:
            decRoomsColl = sorted(
                decRoomsColl,
                key=lambda d: int(d.SeparRoomNumb)
            )
        except:
            # Если номер состоит из цифры и текста
            decRoomsColl = sorted(
                decRoomsColl,
                key=lambda d: int((''.join(re.findall(r'\d', d.SeparRoomNumb))))
            )

        if len(decRoomsColl) > 0:
            vCreator = ViewsCreator(
                doc,
                decRoomsColl,
                roomsColl
            )
            # Создаю планы, и получаю список видов на листы
            viewsColl = vCreator.createViews(
                userTemplateViewMain,
                userOffsetMain,
                userTemplateScheduleMain
            )

            # Очистка от лишних марок
            with Transaction(doc, 'KPLN_Декларация. Очистить виды') as t:
                t.Start()
                for viewColl in viewsColl:
                    roomTags = FilteredElementCollector(doc, viewColl.MainView.Id).OfCategory(BuiltInCategory.OST_RoomTags).ToElements()
                    roomIds = [x.Id for x in viewColl.DecRooms.RoomsDataColl]
                    for tag in roomTags:
                        if "050_Марка квартиры ГОСТ" in tag.Name and (not tag.Room.Id in roomIds):
                            doc.Delete(tag.Id)
                t.Commit()

            # Создаю спец класс по подготовке к созданию листов
            lCreator = ListCreator(doc, viewsColl)
            # Создаю листы, и размещаю на них виды
            lCreator.createLists(userTitleFamilySymbols)
        else:
            ui.forms.Alert(
                'Под заданные параметры - не попало ни одного помещения',
                title="КПЛН_Декларации",
                header="Проверь корректность ввода.",
                exit=True
            )
