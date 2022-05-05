# -*- coding: utf-8 -*-


__author__ = 'Tima Kutsko'
__title__ = "Кол-во на этаж"
__doc__ = 'Подсчет кол-ва элементов на этаж с записью значений в выбранные параметры' \


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
    Viewport, ScheduleSheetInstance, SectionType, UV, LinkElementId, InstanceBinding, Category
from pyrevit import forms
from System.Windows import Window
from System.Collections import IEnumerable
from rpw.ui.forms import Alert
import webbrowser
import wpf


class MyWindowStub(IEnumerable):
    """Заглушка для пользовательского окна"""
    name = "Пусто"

    @staticmethod
    def Defenition():
        pass

    @staticmethod
    def Name(self):
        return self.Name

    def GetEnumerator(self):
        pass


class MyWindow(Window):
    """Форма пользовательского ввода"""

    def __init__(self, catList, doc):
        self.doc = doc
        self.cmbStub = MyWindowStub()
        # Выбранная категория
        self.SelectedCategory = None
        # Выбранный парамтер
        self.SelectedParameter = None

        wpf.LoadComponent(
            self,
            'Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Инструменты.panel\\Doors.stack\\Двери.pulldown\\Кол-во.pushbutton\\Form.xaml'
        )

        # Добавляю категории для обработки
        categories = []
        for cat in catList:
            category = Category.GetCategory(doc, cat)
            categories.append(category)
        categories.sort(key=lambda c: c.Name)
        self.ComboBoxCategory.ItemsSource = categories

    def OnSelectedCategoryChanged(self, sender, e):
        """Метод для заполнения данными по выбранной категории"""
        self.SelectedCategory = self.ComboBoxCategory.SelectedItem
        elemsColl = FilteredElementCollector(self.doc).\
            OfCategoryId(self.SelectedCategory.Id).\
            WhereElementIsNotElementType().\
            ToElements()

        # Добавляю параметры в зависимости от выбранной категории
        self.UpdateParametersData(elemsColl)

        # Добавляю уровни для обработки
        self.UpdateLevelData(elemsColl)

    def OnSelectedParameterChanged(self, sender, e):
        try:
            sender.DataContext.SelectedIndex = sender.SelectedIndex
            if self.ComboBoxCategory.SelectedIndex == -1:
                self.btnApply.IsEnabled = False
            else:
                self.btnApply.IsEnabled = True
        except :
            pass

    def OnButtonApply(self, sender, e):

        self.Close()

    def OnButtonClose(self, sender, e):
        self.Close()

    def OnButtonHelp(self, sender, e):
        webbrowser.open(
            'http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=758'
        )

    def UpdateLevelData(self, elemsColl):
        """Обновление актуального списка уровней"""
        # Коллекция уровней в зависимости от выбранной категории
        self.levelCollection = []
        for elem in elemsColl:
            level = self.doc.GetElement(
                elem.
                get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).
                AsElementId()
            )
            levelCollNames = [lev.Name for lev in self.levelCollection]
            if level.Name not in levelCollNames:
                self.levelCollection.append(level)
        self.levelCollection.sort(key=lambda x: x.Elevation)
        self.iControll.ItemsSource = self.levelCollection

    def UpdateParametersData(self, elemsColl):
        """Обновление актуального списка параметров"""
        # Коллекция параметров в зависимости от выбранной категории
        print(self.cmbStub.Defenition.Na)
        self.parameters = []
        try:
            firstElem = elemsColl[0]
            paramsColl = firstElem.Parameters
            for param in paramsColl:
                if param.IsShared\
                        or param.Id.IntegerValue < 0\
                        and (param.Definition.ParameterType == ParameterType.Text):
                    self.parameters.append(param)
            self.parameters.sort(key=lambda x: x.Definition.Name)
            self.ParamToWrite.ItemsSource = self.parameters
        except Exception:
            self.ParamToWrite.ItemsSource = self.cmbStub


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
        with Transaction(doc, 'КП_Двери_Добавить параметры') as t:
            t.Start()
            for defGroups in sharedParamsFile.Groups:
                for extDef in defGroups.Definitions:
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
MyWindow(catList, doc).ShowDialog()
