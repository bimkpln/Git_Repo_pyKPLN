# -*- coding: utf-8 -*-


__author__ = 'Tima Kutsko'
__title__ = "Кол-во на этаж"
__doc__ = 'Подсчет кол-ва элементов на этаж с записью значений в выбранные параметры' \


from ast import Str
import os
import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows')
from Autodesk.Revit.DB import BuiltInParameter, ParameterFilterRuleFactory,\
    BuiltInParameterGroup, ParameterType, ElementId,\
    ElementParameterFilter, FilterRule, FilteredElementCollector,\
    BuiltInCategory, FilterStringRule, FilterStringEquals,\
    ParameterValueProvider, ElementId, Transaction, ViewFamilyType, ViewFamily,\
    ViewPlan, BoundingBoxXYZ, XYZ, ViewSchedule, SchedulableField,\
    ScheduleFieldType, ScheduleFilter, ScheduleFilterType, ViewSheet,\
    Viewport, ScheduleSheetInstance, SectionType, UV, LinkElementId, InstanceBinding, Category, ElementLevelFilter
from pyrevit import forms
from System.Windows import Window
from pyrevit.forms import WPFWindow
from System.Collections import IEnumerable
from rpw.ui.forms import Alert
from System.Collections.Generic import List
import webbrowser
import wpf


class NameStub():
    """Заглушка для пользовательского окна"""
    Name = "<Пусто>"


class DefinitionStub(IEnumerable):
    """Заглушка для пользовательского окна"""
    Definition = NameStub


class LevelData():
    """Класс для объеденения элементов на одном уровне"""
    global specialParamsGUIDList

    def __init__(self, inputLevel, inputElemColl):
        self.level = inputLevel
        self.elemCollection = inputElemColl
        self.sortParam = None
        self.sortParamDataSet = None
        self.bindingParam = None

        # Список параметров, который нужен для выбора
        # параметра для сортировки
        self.heapParameters = []
        # Список параметров, который подгружен из ФОП
        self.fopParameters = []
        self.__paramBuilder()

    def __paramBuilder(self):
        try:
            firstElem = self.elemCollection.FirstElement()
            paramsColl = firstElem.Parameters
            for param in paramsColl:
                if param.IsShared\
                        or param.Id.IntegerValue < 0\
                        and (param.Definition.ParameterType == ParameterType.Text):
                    try:
                        if param.GUID not in specialParamsGUIDList:
                            self.heapParameters.append(param)
                        else:
                            self.fopParameters.append(param)
                    except Exception:
                        self.heapParameters.append(param)
            self.heapParameters.sort(key=lambda x: x.Definition.Name)
            self.fopParameters.sort(key=lambda x: x.Definition.Name)
        except AttributeError:
            self.heapParameters.append(DefinitionStub())
            self.fopParameters.append(DefinitionStub())


class MyWindow(WPFWindow):
    """Форма пользовательского ввода"""
    def __init__(self, catList, doc):
        self.doc = doc
        self.isClosed = None
        # Выбранная категория
        self.selectedCategory = None
        # Выборка элементов по указанной категории
        self.selectedCatElems = None
        # Список сгенерированных экземплярова LevelData
        self.levelDataList = []

        # Загружаю окно
        WPFWindow.__init__(self, 'Form.xaml')

        # Добавляю категории для обработки
        categories = []
        for cat in catList:
            category = Category.GetCategory(doc, cat)
            categories.append(category)
        categories.sort(key=lambda c: c.Name)
        self.ComboBoxCategory.ItemsSource = categories

        # Добавляю уровни проекта для обработки
        self.levelsColl = FilteredElementCollector(self.doc).\
            OfCategory(BuiltInCategory.OST_Levels).\
            WhereElementIsNotElementType()

    def OnSelectedCategoryChanged(self, sender, e):
        """Метод для заполнения данными по выбранной категории"""
        self.levelDataList = []
        self.selectedCategory = self.ComboBoxCategory.SelectedItem

        # Добавляю общую коллекцию по данным из выбранного параметра
        self.selectedCatElems = FilteredElementCollector(self.doc).\
            OfCategoryId(self.selectedCategory.Id).\
            WhereElementIsNotElementType()

        # Добавляю коллекцию элементов по этажам
        for currentLevel in self.levelsColl:
            elemsColl = FilteredElementCollector(self.doc).\
                OfCategoryId(self.selectedCategory.Id)
            levelFilter = ElementLevelFilter(currentLevel.Id)
            elemsColl.WherePasses(levelFilter)
            if elemsColl.FirstElement():
                levelData = LevelData(currentLevel, elemsColl)
                self.levelDataList.append(levelData)

        # Добавляю параметры в зависимости от выбранной категории
        self.UpdateParametersDataOnWindow()

        # Добавляю уровни в зависимости от выбранной категории
        self.UpdateLevelDataElementsOnWindow()

    def UpdateParametersDataOnWindow(self):
        """Обновление актуального списка параметров"""
        if self.levelDataList:
            self.ParamToSort.ItemsSource = self.levelDataList[0].heapParameters
        else:
            self.ParamToSort.ItemsSource = [DefinitionStub()]

    def UpdateLevelDataElementsOnWindow(self):
        """Обновление актуального списка уровней"""
        if self.levelDataList:
            self.levelDataList.sort(key=lambda x: x.level.Name)
            self.iControll.ItemsSource = self.levelDataList
        else:
            stubList = [LevelData(NameStub(), [DefinitionStub()])]
            self.iControll.ItemsSource = stubList

    def OnSortedParameterChanged(self, sender, e):
        selectedParam = sender.SelectedItem
        sortParamDataSet = set()
        for elem in self.selectedCatElems:
            data = elem.\
                LookupParameter(selectedParam.Definition.Name).\
                AsString()
            sortParamDataSet.add(data)
        # Устанавливаю для каждой коллекции параметр для сортировки
        for levelData in self.levelDataList:
            levelData.sortParam = selectedParam
            levelData.sortParamDataSet = sortParamDataSet

    def OnSelectedParameterChanged(self, sender, e):
        self.btnApply.IsEnabled = True
        # Устанавливаю для каждой коллекции параметр для записи
        # (привязка именно к коллекции - благодаря биндингу в wpf)
        sender.DataContext.bindingParam = sender.SelectedItem
        for levelData in self.levelDataList:
            if levelData.bindingParam is None:
                self.btnApply.IsEnabled = False

    def OnButtonApply(self, sender, e):
        self.Close()

    def OnButtonClose(self, sender, e):
        self.isClosed = True
        self.Close()

    def OnButtonHelp(self, sender, e):
        webbrowser.open(
            'http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=758'
        )


# Ключевые переменные
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
comParamsFilePath = "X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Двери.txt"
# Список категорий, которые будут обрабатываться
catList = [
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_CurtainWallMullions,
    BuiltInCategory.OST_CurtainWallPanels
]
# Список параметров, которые нужно обработать (заполняется из ФОП)
specialParamsGUIDList = []


# Подгружаю параметры
if os.path.exists(comParamsFilePath):
    try:
        # Создаю спец класс CategorySet и добавляю в него зависимости
        # (категории)
        catSetElements = app.Create.NewCategorySet()
        for cat in catList:
            catSetElements.Insert(doc.Settings.Categories.get_Item(cat))

        # Забираю все парамтеры проекта в список
        prjParamsNamesList = []
        paramBind = doc.ParameterBindings
        fIterator = paramBind.ForwardIterator()
        fIterator.Reset()
        while fIterator.MoveNext():
            prjParamsNamesList.append(fIterator.Key.Name)

        # Забираю все параметры из спец ФОПа
        app.SharedParametersFilename = comParamsFilePath
        sharedParamsFile = app.OpenSharedParameterFile()

        # Добавляю недостающие парамтеры в проект
        with Transaction(doc, 'КП_Кол. на этажах_Добавить параметры') as t:
            t.Start()
            for defGroups in sharedParamsFile.Groups:
                for extDef in defGroups.Definitions:
                    # Список параметров (GUID) из ФОП, которые используются
                    # для плагина
                    specialParamsGUIDList.append(extDef.GUID)

                    # Добавляю параметры
                    if extDef.Name not in prjParamsNamesList:
                        paramBind = doc.ParameterBindings
                        newIB = app.Create.NewInstanceBinding(catSetElements)
                        paramBind.Insert(
                            extDef,
                            newIB,
                            BuiltInParameterGroup.PG_DATA
                        )

                        # Разворачиваю проход по параметрам проекта
                        revFIterator = doc.ParameterBindings.ReverseIterator()
                        while revFIterator.MoveNext():
                            if extDef.Name == revFIterator.Key.Name:

                                # Включаю вариативность между экземплярами
                                # групп в Revit
                                revFIterator.Key.SetAllowVaryBetweenGroups(
                                    doc,
                                    True
                                )
                                break
            t.Commit()
    except Exception as e:
        Alert(
            "Ошибка при загрузке параметров:\n[{}]".format(str(e)),
            title="Загрузчик параметров", header="Ошибка"
        )
else:
    Alert(
        "Файл общих параметров не найден:\nX:\\BIM\\4_ФОП\\00_Архив\\КП_Двери.txt",
        title="Загрузчик параметров",
        header="Ошибка"
    )


# Обрабатываю элементы
myWindow = MyWindow(catList, doc)
myWindow.ShowDialog()
if not myWindow.isClosed:
    with Transaction(doc, 'КП_Кол. на этажах_Заполнить параметры') as t:
        t.Start()
        # Чистка предыдущих запусков
        for elem in myWindow.selectedCatElems:
            for param in specialParamsGUIDList:
                elem.get_Parameter(param).Set(0)

        # Заполнение данными
        for levelData in myWindow.levelDataList:
            elemColl = levelData.elemCollection
            bindingParam = levelData.bindingParam
            sortParam = levelData.sortParam
            sortParamSet = levelData.sortParamDataSet

            # Получаю элементы по этажу
            for paramData in sortParamSet:
                trueElements = list(filter(
                    lambda x: (
                        x.LookupParameter(sortParam.Definition.Name).AsString() ==
                        paramData
                    ),
                    elemColl)
                )

                # Записываю новые данные
                for elem in trueElements:
                    elem.LookupParameter(bindingParam.Definition.Name).Set(1)
        t.Commit()

