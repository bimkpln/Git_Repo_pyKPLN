#coding: utf-8


__title__ = "Создание деклараций"
__author__ = 'Tima Kutsko'
__doc__ = 'Автоматическая генерация видов, листов, спецификаций\n'\
          'для для создания деклараций\n'\


import string
from tokenize import Double
import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import BuiltInParameter, ParameterFilterRuleFactory,\
    ParameterFilterElement, ViewDuplicateOption, ElementId,\
    ElementParameterFilter, FilterRule, FilteredElementCollector,\
    BuiltInCategory, FilterStringRule, FilterStringEquals,\
    ParameterValueProvider,ElementId, Transaction, ViewFamilyType, ViewFamily,\
    ViewPlan, BoundingBoxXYZ, XYZ, ViewSchedule, SchedulableField,\
    ScheduleFieldType, ScheduleFilter, ScheduleFilterType, ViewSheet,\
    Viewport, ScheduleSheetInstance, SectionType, UV, LinkElementId
from Autodesk.Revit.Creation import Document
from System import Guid
from System.Collections.Generic import List


class RoomData:
    """ Класс для объеденения помещений и их атрибутов
    """

    def __init__(self, room, numbParam, sectParam):
        self.room = room
        self.numbParam = numbParam
        self.sectParam = sectParam


class DecRooms:
    """ Класс для создания коллекции помещений типа RoomData, объедененных
        под одним параметром номера квартиры/аренд. помещения
    """

    def __init__(self, roomsDataColl, roomNumb):
        self.roomsDataColl = roomsDataColl
        self.roomNumb = roomNumb
        try:
            self.level = roomsDataColl[0].room.Level
            self.roomDepartment = roomsDataColl[0].room.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString()
        except Exception:
            raise Exception('Ничего не попало в коллекцию помещений')


class ViewsCreator:
    """ Класс для создание видов для декларации:
    """

    def __init__(self, doc, decRooms, roomNumbParam, sectionNumbParam, allRooms, roomTag):
        self.doc = doc
        self.decRooms = decRooms
        self.roomNumbParam = roomNumbParam
        self.sectionNumbParam = sectionNumbParam
        self.allRooms = allRooms
        self.roomTag = roomTag


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
        bBox.Min = XYZ(xmin - offset, ymin - offset, zmin)
        bBox.Max = XYZ(xmax + offset, ymax + offset, zmax)
        return bBox


    def __roomMarker(self, decRooms, view):
        for r in decRooms.roomsDataColl:
            room = r.room
            point = room.Location.Point
            if isinstance(point, XYZ):
                uv = UV(point.X, point.Y)
            elif isinstance(point, UV):
                uv = UV(point.U, point.V)
            roomTag = self.doc.Create.NewRoomTag(LinkElementId(room.Id), uv, view.Id)
            roomTag.RoomTagType = self.roomTag.RoomTagType


    def __createCrop(self, decCol, offset):
        currentSection = decCol.roomsDataColl[0].sectParam
        equalSectRoomsColl = filter(
            lambda r: r.get_Parameter(self.sectionNumbParam).AsString() == currentSection.AsString(),
            self.allRooms)
        roomNumbRoomsColl = list()
        for r in decCol.roomsDataColl:
            roomNumbRoomsColl.append(r.room)
        sectBBox = self.__cropExpand(equalSectRoomsColl, offset)
        roomNumbBBox = self.__cropExpand(roomNumbRoomsColl, offset)
        return sectBBox, roomNumbBBox


    def __creatorPlan(self, decRooms, rNumber, rNumberParam, viewTemplate, viewFamilyTypeId, currentBBox, bugName, isMain):
        # Создание вида
        newName = "KPLN_Декларация_{}_{}_{}".format(
            bugName,
            rNumberParam.Name,
            rNumber)
        currentLevelId = decRooms.level.Id
        newView = ViewPlan.Create(
            self.doc,
            viewFamilyTypeId,
            currentLevelId)
        newView.get_Parameter(BuiltInParameter.VIEW_NAME).Set(newName)
        # Примененеие подрезки
        newView.CropBox = currentBBox
        newView.CropBoxActive = True
        newView.CropBoxVisible = True
        # Применение заранее заготовленного шаблона
        newView.ViewTemplateId = viewTemplate.Id
        # Применение фильтрации (только для общего вида)
        if isMain:
            fRule = ParameterFilterRuleFactory.\
                CreateNotContainsRule(rNumberParam.Id, rNumber, False)
            filterRules = ElementParameterFilter(fRule)
            filterName = "{}_{}".format(rNumberParam.Name, rNumber)
            IfilterCat = List[ElementId]()
            IfilterCat.Add(ElementId(BuiltInCategory.OST_Rooms))
            newFilter = ParameterFilterElement.Create(
                self.doc,
                filterName,
                IfilterCat,
                filterRules)
            newView.AddFilter(newFilter.Id)
            newView.SetFilterVisibility(newFilter.Id, False)
        return newView


    def __creatorSchedule(self, decRooms, rNumber, rNumberParam, templateSchedule):
        newScheduleView = ViewSchedule.CreateSchedule(
            self.doc,
            decRooms.roomsDataColl[0].room.Category.Id
        )
        newScheduleViewName = "KPLN_Декларация_{}_{}".format(
            rNumberParam.Name,
            rNumber
        )
        newScheduleView.Name = newScheduleViewName
        schDef = newScheduleView.Definition

        schedulFieldNumber = SchedulableField(
            ScheduleFieldType.Instance,
            decRooms.roomsDataColl[0].room.get_Parameter(self.roomNumbParam).Id
        )
        schFieldNumber = schDef.AddField(schedulFieldNumber)
        schFType = ScheduleFilterType.Equal
        schFilter = ScheduleFilter(
            schFieldNumber.FieldId,
            schFType,
            decRooms.roomNumb
        )
        schDef.AddFilter(schFilter)

        # Применение заранее заготовленного шаблона. 
        # Данный шаблон уже содержит все поля, но поле с параметром номера - 
        # добавляется специально, для задания фильтрации. Иначе
        # нужно закрывать транзакцию, чтобы получить список полей спеки
        newScheduleView.ViewTemplateId = templateSchedule.Id
        self.doc.Regenerate()
        table = newScheduleView.GetTableData()
        section = table.GetSectionData(SectionType.Body)
        section.SetColumnWidth(0, 20 / 304.1)
        section.SetColumnWidth(1, 160 / 304.1)
        section.SetColumnWidth(2, 20 / 304.1)
        return newScheduleView


    def createViews(self, viewTemplateMain, viewTemplateFlat, offsetMain, offsetRoom, templateSchedule):
        with Transaction(self.doc, 'KPLN_Декларация. Создать виды') as t:
            viewFamTypeColl = FilteredElementCollector(self.doc).\
                OfClass(ViewFamilyType).\
                WhereElementIsElementType()
            for viewFamType in viewFamTypeColl:
                if viewFamType.ViewFamily == ViewFamily.FloorPlan:
                    viewFamilyTypeId = viewFamType.Id
                    break

            t.Start()
            decViewsDataList = list()
            for currentDecRooms in self.decRooms:
                roomNumber = currentDecRooms.roomNumb
                roomNumberParam = self.doc.GetElement(
                    currentDecRooms.
                    roomsDataColl[0].
                    numbParam.
                    Id
                )
                # Создание подрезок по выбранным параметрам (BoundingBoxXYZ)
                mainBBox = self.__createCrop(
                    currentDecRooms,
                    offsetMain
                )[0]
                roomBBox = self.__createCrop(
                    currentDecRooms,
                    offsetRoom
                )[1]
                # Создание общих планов
                newMainView = self.__creatorPlan(
                    currentDecRooms,
                    roomNumber,
                    roomNumberParam,
                    viewTemplateMain,
                    viewFamilyTypeId,
                    mainBBox,
                    "Жук",
                    True
                )
                # Создание поквартирных планов
                newRoomView = self.__creatorPlan(
                    currentDecRooms,
                    roomNumber,
                    roomNumberParam,
                    viewTemplateFlat,
                    viewFamilyTypeId,
                    roomBBox,
                    currentDecRooms.roomDepartment,
                    False
                )
                # Маркировка помещений
                self.__roomMarker(currentDecRooms, newRoomView)
                # Создание поквартирных спецификаций
                newScheduleView = self.__creatorSchedule(
                    currentDecRooms,
                    roomNumber,
                    roomNumberParam,
                    templateSchedule
                )
                # Создание коллекции видов
                newListName = "KPLN_Декларация_{}_{}".format(
                    roomNumberParam.Name,
                    roomNumber
                )
                newViewsCollName = "{} - {}".format(
                    currentDecRooms.roomDepartment,
                    roomNumber
                )
                decViewData = DecViewsData(
                    newViewsCollName,
                    newListName,
                    newMainView,
                    newRoomView,
                    newScheduleView
                )
                decViewsDataList.append(decViewData)
            t.Commit()
            return decViewsDataList


class DecViewsData:
    """ Класс для создании коллекций видов для размещения на листе
    """

    def __init__(self, viewsCollName, listName, mainView, roomView, scheduleView):
        self.viewsCollName = viewsCollName
        self.listName = listName
        self.mainView = mainView
        self.roomView = roomView
        self.scheduleView = scheduleView


class ListCreator:
    """ Класс для создание листов для декларации
    """

    def __init__(self, doc, decViews):
        self.doc = doc
        self.decViews = decViews


    def __bBoxCenter(bBox):
        return XYZ(
            (bBox.Max.X + bBox.Min.X) / 2,
            (bBox.Max.Y + bBox.Min.Y) / 2,
            0
        )


    def __dropViews(self, newSheet, currentView, tBlockBBox, isMain=False, isRoom=False):
        #vpTypeId = currentView.GetTypeId()
        if isinstance(currentView, ViewPlan):
            if isMain:
                # Определяю центр НИЖЕ средней точки основной надписи
                vpCenter = XYZ(
                    (tBlockBBox.Max.X + tBlockBBox.Min.X) / 2,
                    (tBlockBBox.Max.Y) * 1 / 6,
                    0
                )
            elif isRoom:
                # Определяю центр ВЫШЕ средней точки основной надписи
                vpCenter = XYZ(
                    (tBlockBBox.Max.X + tBlockBBox.Min.X) / 2,
                    (tBlockBBox.Max.Y) * 3 / 4,
                    0
                )
            Viewport.Create(self.doc,
                newSheet.Id,
                currentView.Id,
                vpCenter
            )
        elif isinstance(currentView, ViewSchedule):
            # Определяю центр основной надписи
            vpCenter = XYZ(
                # Подгонка по правому краю на 5 мм
                tBlockBBox.Min.X + 5 / 304.1,
                (tBlockBBox.Max.Y) * 3 / 5,
                0
            )
            ScheduleSheetInstance.Create(
                self.doc,
                newSheet.Id,
                currentView.Id,
                vpCenter
            )

    def createLists(self, titleFamily):
        with Transaction(self.doc, 'KPLN_Декларация. Создать листы') as t:
            t.Start()
            tFS = titleFamily.Symbol
            currentListNumber = 0
            for decView in self.decViews:
                currentListNumber += 1
                newSheet = ViewSheet.Create(self.doc, tFS.Id)
                newSheet.Name = decView.listName
                newSheet.SheetNumber = "DEC_{}".format(currentListNumber)
                # Заполняю параметр для маркировки листа (спец. метка)
                newSheet.get_Parameter(userListMarkParam).Set(decView.viewsCollName)
                self.doc.Regenerate()

                # Размещаю виды на листы
                sheetTBlock = FilteredElementCollector(self.doc, newSheet.Id).\
                    OfCategory(BuiltInCategory.OST_TitleBlocks).\
                    WhereElementIsNotElementType().\
                    FirstElement()
                tBlockBBox = sheetTBlock.get_BoundingBox(
                    self.doc.GetElement(newSheet.Id)
                )
                self.__dropViews(
                    newSheet,
                    decView.mainView,
                    tBlockBBox,
                    True,
                    False
                )
                self.__dropViews(
                    newSheet,
                    decView.roomView,
                    tBlockBBox,
                    False,
                    True
                )
                self.__dropViews(
                    newSheet,
                    decView.scheduleView,
                    tBlockBBox
                )
            t.Commit()


# TEST - нужно удалить перед выпуском в работу. Он нужен для фильтрации помещений ТОЛЬКО на активном виде
def passesColl(doc, builInParam, builInCat, strName, TEST=False):
    vpr = ParameterValueProvider(ElementId(builInParam))
    eRule = FilterStringRule(vpr, FilterStringEquals(), strName, True)
    eFilter = ElementParameterFilter(eRule)
    if TEST:
        elColl = FilteredElementCollector(doc, doc.ActiveView.Id).\
            OfCategory(builInCat).\
            WherePasses(eFilter)
    else:
        elColl = FilteredElementCollector(doc).\
            OfCategory(builInCat).\
            WherePasses(eFilter)
    return elColl

#input part
userRoomTypeData = "Квартира"
userRoomTagName = "041_Номер Помещения(Кваритры) GOST Common: Номер_2.5 мм"
userTemplateViewMainName = "KPLN_Декларации_Жук"
userTemplateViewFlatName = "KPLN_Декларации_Квартира"
userTemplateScheduleName = "KPLN_Декларации_Спецификация"
userTitleName = "DEC_100_A4"
userOffsetMain = 700 / 304.1
userOffsetRoom = 350 / 304.1
# Объеденяющий параметр, например: сквозной номер квартиры в здании
userRoomNumbParam = Guid("533266b9-e2bc-42b4-8c79-2c82f64bbee2")
# Объеденяющий параметр, например: номер секции
userSectionNumbParam = Guid("735250d9-82a2-4a0f-ac44-85fc69a96353")
# Параметр, для записи в метку листа
userListMarkParam = Guid("e313f126-7e51-4a5d-a45a-7c6dfe02124a")

#main part
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Поиск помещений, с указанным пользователем именем типа
roomsColl = passesColl(
    doc,
    BuiltInParameter.ROOM_DEPARTMENT,
    BuiltInCategory.OST_Rooms,
    userRoomTypeData,
    True).ToElements()

# Поиск шаблона для жука, с указанным пользователем именем
userTemplateViewMain = passesColl(
    doc,
    BuiltInParameter.VIEW_NAME,
    BuiltInCategory.OST_Views,
    userTemplateViewMainName).FirstElement()

# Поиск шаблона для квартиры, с указанным пользователем именем
userTemplateViewFlat = passesColl(
    doc,
    BuiltInParameter.VIEW_NAME,
    BuiltInCategory.OST_Views,
    userTemplateViewFlatName).FirstElement()

# Поиск шаблона для спецификации
userTemplateSchedule = passesColl(
    doc,
    BuiltInParameter.VIEW_NAME,
    BuiltInCategory.OST_Schedules,
    userTemplateScheduleName).FirstElement()

# Поиск основной надписи
userTitle = passesColl(
    doc,
    BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
    BuiltInCategory.OST_TitleBlocks,
    userTitleName).WhereElementIsNotElementType().FirstElement()
# Поиск марки для помещений
userTag = passesColl(
    doc,
    BuiltInParameter.ELEM_FAMILY_PARAM,
    BuiltInCategory.OST_RoomTags,
    userRoomTagName).WhereElementIsNotElementType().FirstElement()

# Создаю коллекцию RoomData, обедененных по параметру userRoomNumbParam
throughRoomNumbsSet = set()
decRoomsColl = list()
for room in roomsColl:
    currentThroughRoomNumb = room.get_Parameter(userRoomNumbParam).AsString()
    throughRoomNumbsSet.add(currentThroughRoomNumb)
for throughtRoomNumb in throughRoomNumbsSet:
    roomDataList = list()
    equalThroughRoomsColl = filter(
        lambda r: r.get_Parameter(userRoomNumbParam).AsString() == throughtRoomNumb,
        roomsColl)

    for room in equalThroughRoomsColl:
        roomDataList.append(RoomData(
            room,
            room.get_Parameter(userRoomNumbParam),
            room.get_Parameter(userSectionNumbParam)))

    decRooms = DecRooms(roomDataList, throughtRoomNumb)
    decRoomsColl.append(decRooms)

# Создаю спец класс по подготовке к созданию планов
vCreator = ViewsCreator(
    doc,
    decRoomsColl,
    userRoomNumbParam,
    userSectionNumbParam,
    roomsColl,
    userTag
)
# Создаю планы, и получаю список видов на листы
viewsColl = vCreator.createViews(
    userTemplateViewMain,
    userTemplateViewFlat,
    userOffsetMain,
    userOffsetRoom,
    userTemplateSchedule
)

# Создаю спец класс по подготовке к созданию листов
lCreator = ListCreator(doc, viewsColl)
# Создаю листы, и размещаю на них виды
lCreator.createLists(userTitle)
