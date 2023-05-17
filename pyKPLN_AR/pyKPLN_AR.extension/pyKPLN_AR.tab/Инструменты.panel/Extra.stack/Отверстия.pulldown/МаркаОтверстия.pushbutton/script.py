# -*- coding: utf-8 -*-


__author__ = 'Tsimafei Kutsko'
__title__ = "Маркировать проемы"
__doc__ = 'Присвоение марок для стеновых проемов'\
    '«199_Отверстие в стене прямоугольное»'


import math
import clr
clr.AddReference('RevitAPI')
clr.AddReference('System')
import os
from Autodesk.Revit.DB import BuiltInParameter, FilteredElementCollector,\
    BuiltInCategory, BuiltInParameterGroup, StorageType,\
    Category, InstanceBinding, Transaction, ParameterType
from rpw.ui.forms import CommandLink, TaskDialog
from rpw.ui.forms import TextInput, Alert
from System import Guid
from System.Collections.Generic import *
from System.Windows.Forms import Form, Button, ComboBox, FormBorderStyle,\
    ComboBoxStyle, ListView, HorizontalAlignment, DockStyle,\
    View, SortOrder, ListViewItem, Application, ColumnHeader
from System.Drawing import Point, Size, Icon
from collections import OrderedDict


class PickParameter(Form):
    def __init__(self):
        self.Name = "Параметр"
        self.Text = "Выберите параметр"
        self.Size = Size(205, 110)
        self.Icon = Icon(iconPath)
        self.button_apply = Button(Text="Применить")
        self.combo_box = ComboBox()
        self.ControlBox = False
        self.TopMost = True
        self.MinimumSize = Size(205, 110)
        self.MaximumSize = Size(205, 110)
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow
        self.ShowInTaskbar = False
        self.CenterToScreen()

        self.combo_box.Parent = self
        self.combo_box.Items.Add("<по умолчанию>")
        self.combo_box.Text = "<по умолчанию>"
        self.combo_box.DropDownStyle = ComboBoxStyle.DropDownList
        self.combo_box.Location = Point(12, 12)
        self.combo_box.Size = Size(166, 21)
        self.paramlist = []
        elemColl = FilteredElementCollector(doc).\
            OfCategory(BuiltInCategory.OST_MechanicalEquipment).\
            WhereElementIsNotElementType()
        for element in elemColl:
            famName = element.Symbol.FamilyName
            if famName.startswith("199_Отверстие"):
                for j in element.Parameters:
                    if j.IsShared\
                            and j.UserModifiable\
                            and j.StorageType == StorageType.String\
                            and not j.IsReadOnly:
                        self.paramlist.append(j.Definition.Name)
                break
        self.paramlist.sort()
        for i in self.paramlist:
            self.combo_box.Items.Add(i)
        self.button_apply.Parent = self
        self.button_apply.Location = Point(12, 40)
        self.button_apply.Size = Size(75, 23)
        self.button_apply.Text = "Применить"
        self.button_apply.Click += self.Commit

    def Commit(self, sender, args):
        global write_parameter

        write_parameter = self.combo_box.Text
        self.Close()


class CreateWindow(Form):
    def __init__(self):
        self.Name = "Связи для маркировки"
        self.Text = "Выберите связи с общей маркировкой"
        self.Size = Size(418, 608)
        self.Icon = Icon(iconPath)
        self.button_create = Button(Text="Ок")
        self.ControlBox = True
        self.TopMost = True
        self.MinimumSize = Size(418, 480)
        self.MaximumSize = Size(418, 480)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.CenterToScreen()

        self.listbox = ListView()
        self.c_cb = ColumnHeader()
        self.c_cb.Text = ""
        self.c_cb.Width = -2
        self.c_cb.TextAlign = HorizontalAlignment.Center
        self.c_name = ColumnHeader()
        self.c_name.Text = "Путь"
        self.c_name.Width = -2
        self.c_name.TextAlign = HorizontalAlignment.Left

        self.SuspendLayout()
        self.listbox.Dock = DockStyle.Fill
        self.listbox.View = View.Details

        self.listbox.Parent = self
        self.listbox.Size = Size(400, 400)
        self.listbox.Location = Point(1, 1)
        self.listbox.FullRowSelect = True
        self.listbox.GridLines = True
        self.listbox.AllowColumnReorder = True
        self.listbox.Sorting = SortOrder.Ascending
        self.listbox.Columns.Add(self.c_cb)
        self.listbox.Columns.Add(self.c_name)
        self.listbox.LabelEdit = True
        self.listbox.CheckBoxes = True
        self.listbox.MultiSelect = True

        self.button_ok = Button(Text="Ok")
        self.button_ok.Parent = self
        self.button_ok.Location = Point(10, 410)
        self.button_ok.Click += self.OnOk

        self.button_ok = Button(Text="Отмена")
        self.button_ok.Parent = self
        self.button_ok.Location = Point(100, 410)
        self.button_ok.Click += self.OnCancel

        self.item = []
        linkColl = FilteredElementCollector(doc).\
            OfCategory(BuiltInCategory.OST_RvtLinks).\
            WhereElementIsNotElementType()
        for link in linkColl:
            try:
                document = link.GetLinkDocument()
                self.item.append(ListViewItem())
                self.item[len(self.item)-1].Text = ""
                self.item[len(self.item)-1].Checked = False
                self.item[len(self.item)-1].SubItems.Add(
                    "{} ({})".format(document.Title, document.PathName)
                )
                self.listbox.Items.Add(self.item[len(self.item)-1])
            except Exception:
                pass

    def OnOk(self, sender, args):
        global paths
        for i in self.item:
            if i.Checked:
                si = i.SubItems
                paths.append(si[1].Text)
        self.Close()

    def OnCancel(self, sender, args):

        global next
        next = False
        self.Close()


class SSymbol():
    def __init__(self, name, width, height, offset, levElevation, elemElevation, element, host, islink):
        self.Type = name
        self.Width = self.zero(width)
        self.Height = self.zero(height)
        self.Offset = self.zero(offset)
        self.LevelElev = levElevation
        self.ElementElev = elemElevation
        self.Host = ""
        if host:
            self.Host = "1"
        else:
            self.Host = "0"
        self.Element = element
        self.Link = islink

    def IsLink(self):
        return self.Link

    def zero(self, integer):
        newint = math.fabs(integer)
        z = ""
        for i in range(0, 10 - len(str(newint))):
            z += "0"
        z += str(newint)
        if integer < 0:
            z = "-{}".format(z)
        return z


def InList(item, list):
    try:
        for i in list:
            if i == item:
                return True
        return False
    except Exception:
        return False


# Запись в класс SSymbol
def seterSSymbol(doc, isLink):
    global symbols

    elemColl = FilteredElementCollector(doc).\
        OfCategory(BuiltInCategory.OST_MechanicalEquipment).\
        WhereElementIsNotElementType()
    for element in elemColl:
        famSymb = element.Symbol
        famName = famSymb.FamilyName
        if famName.startswith("199_Отверстие")\
                or famName.startswith("Отверстие в стене"):
            try:
                name = famSymb.\
                    get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).\
                    AsString()
                width = int(round(
                        element.get_Parameter(guidWidth).AsDouble() * 3.048, 1
                    ) * 100
                )
                height = int(round(
                        element.get_Parameter(guidHeight).AsDouble() * 3.048, 1
                    ) * 100
                )
                offset = int(round(
                        element.get_Parameter(
                            BuiltInParameter.INSTANCE_ELEVATION_PARAM
                        ).AsDouble() * 3.048, 1) * 100
                )
                # Отметка уровня
                levElevation = int(round(
                        doc.GetElement(element.LevelId).Elevation * 3.048, 1
                    ) * 100
                )
                # Отметка элемента, относительно точки съемки
                elemElevation = int(round(
                        element.Location.Point.Z * 304.8 / 5, 0
                    ) * 5
                )
                host = element.Host.Name.startswith("00_")
                symbols.append(SSymbol(
                    name, width, height, offset,
                    levElevation, elemElevation, element,
                    host, isLink
                    )
                )
            except Exception as e_add:
                print(str(e_add) + "!!!")


def trueOrderDictCreator(sList, isKR):
    """ Возвращает упорядоченный словарь с экземплярами класса SSymbol,
    разделенные по заданным условиям. Ключ - условия группирования,
    значения - все отверстия под данной группой"""
    typesDict = OrderedDict()
    if isKR:
        trueSList = list(
            filter(
                lambda s: s.Element.Host.Name.startswith("00_"),
                sList
            )
        )
    else:
        trueSList = list(
            filter(
                lambda s: not s.Element.Host.Name.startswith("00_"),
                sList
            )
        )

    for symbol in trueSList:
        dictKey = symbol.Type +\
            "~" +\
            symbol.Width.ToString() +\
            "~" +\
            symbol.Height.ToString() +\
            "~" +\
            symbol.Offset.ToString() +\
            "~" +\
            symbol.ElementElev.ToString()
        try:
            typesDict[dictKey].append(symbol)
        except Exception:
            typesDict[dictKey] = [symbol]
    return typesDict


def mainValueSetter(tDict, isKR):
    """Для заполнения параметров в зависимости от основания отверстия"""
    global cnt
    if isKR:
        comment = "см. раздел КЖ"
    else:
        comment = ""
    for key, symbols in tDict.items():
        # Сортировка по ширине ВНУТРИ ключа
        symbols = sorted(symbols, key=lambda s: s.Width)
        cnt += 1
        for symbol in symbols:
            if not symbol.IsLink():
                if write_parameter == "None"\
                        or write_parameter == "<по умолчанию>":
                    symbol.\
                        Element.\
                        get_Parameter(BuiltInParameter.DOOR_NUMBER).\
                        Set(value + str(cnt))
                else:
                    symbol.\
                        Element.LookupParameter(write_parameter).\
                        Set(value + str(cnt))
                symbol.\
                    Element.\
                    LookupParameter("00_Комментарий").\
                    Set(comment)
                symbol.\
                    Element.\
                    LookupParameter("00_Фасад").\
                    Set(symbol.LevelElev.ToString())


# Ключевые переменные
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
comParamsFilePath = "X:\\BIM\\4_ФОП\\02_Для плагинов\\КП_Плагины_Общий.txt"
iconPath = "X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico"
paramsList = ["00_Комментарий", "00_Фасад"]
trueCategory = BuiltInCategory.OST_MechanicalEquipment
paths = []
documents = []
# Ширина
guidWidth = Guid("8f2e4f93-9472-4941-a65d-0ac468fd6a5d")
# Высота
guidHeight = Guid("da753fe3-ecfa-465b-9a2c-02f55d0c2ff1")

# Подгружаю параметры
if os.path.exists(comParamsFilePath):
    try:
        # Создаю спец класс CategorySet и добавляю в него зависимости
        # (категории)
        catSetElements = app.Create.NewCategorySet()
        catSetElements.Insert(
            doc.Settings.Categories.get_Item(trueCategory)
        )

        # Забираю все парамтеры проекта в список
        prjParamsNamesList = []
        paramBind = doc.ParameterBindings
        fIterator = paramBind.ForwardIterator()
        fIterator.Reset()
        while fIterator.MoveNext():
            d_Definition = fIterator.Key
            d_Name = fIterator.Key.Name
            d_Binding = fIterator.Current
            d_catSet = d_Binding.Categories
            if d_Name in paramsList\
                    and d_Binding.GetType() == InstanceBinding\
                    and str(d_Definition.ParameterType) == "Text"\
                    and d_catSet.Contains(
                        Category.GetCategory(
                            doc,
                            trueCategory
                        )
                    ):
                prjParamsNamesList.append(fIterator.Key.Name)

        # Забираю все параметры из спец ФОПа
        app.SharedParametersFilename = comParamsFilePath
        sharedParamsFile = app.OpenSharedParameterFile()

        # Добавляю недостающие парамтеры в проект
        with Transaction(doc, 'КП_Добавить параметры') as t:
            t.Start()
            for defGroups in sharedParamsFile.Groups:
                if defGroups.Name == "АР_Отверстия":
                    for extDef in defGroups.Definitions:
                        # Добавляю параметры (если они не были ранее загружены)
                        if extDef.Name not in prjParamsNamesList:
                            paramBind = doc.ParameterBindings
                            newIB = app.\
                                Create.\
                                NewInstanceBinding(catSetElements)
                            paramBind.Insert(
                                extDef,
                                newIB,
                                BuiltInParameterGroup.PG_DATA
                            )

                            # Разворачиваю проход по параметрам проекта
                            revFIterator = doc.\
                                ParameterBindings.\
                                ReverseIterator()
                            while revFIterator.MoveNext():
                                if extDef.Name == revFIterator.Key.Name\
                                        and extDef.ParameterType == ParameterType.Text:
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
        "Файл общих параметров не найден:{}".format(comParamsFilePath),
        title="Загрузчик параметров",
        header="Ошибка"
    )

write_parameter = "None"
commands = [
    CommandLink('Да', return_value=True),
    CommandLink('Нет', return_value=False),
    CommandLink('Отмена', return_value="Отмена")
]
dialog = TaskDialog(
    'Учитывать маркировку подгруженных связей во время маркировки?',
    title="Учет связей",
    title_prefix=False,
    content="Опция позволяет создавать общую систему марок для нескольких проектов.",
    commands=commands,
    footer='',
    show_close=False
)

dialog_par = TaskDialog(
    'Использовать параметр «по умолчанию» для записи значения марки?',
    title="Параметр для записи",
    title_prefix=False,
    content="",
    commands=commands,
    footer='',
    show_close=False
)

ShowForm = dialog.show()
if ShowForm != "Отмена":
    next = True
    if ShowForm:
        form = CreateWindow()
        Application.Run(form)
    if next:
        ParamForm = dialog_par.show()
        if not ParamForm:
            form2 = PickParameter()
            Application.Run(form2)
        if len(paths) != 0:
            for name in paths:
                linkColl = FilteredElementCollector(doc).\
                    OfCategory(BuiltInCategory.OST_RvtLinks).\
                    WhereElementIsNotElementType()
                for link in linkColl:
                    try:
                        document = link.GetLinkDocument()
                        if "{} ({})".format(
                                    document.Title, document.PathName
                                ) == name:
                            documents.append(document)
                    except Exception:
                        pass

        symbols = []
        value = str(
            TextInput(
                'Префикс для маркировки',
                default="",
                description="«[префикс][марка]»",
                exit_on_close=False
            )
        )

        if len(documents) != 0:
            for document in documents:
                try:
                    seterSSymbol(document, True)
                except Exception:
                    continue

        seterSSymbol(doc, False)
        # Сортировка id элемента, и финально по уровню. Тройная сортировка
        # нужна при использовании связанных файлов
        symbols = sorted(symbols, key=lambda s: s.Element.Id)
        symbols = sorted(symbols, key=lambda s: (s.Width, s.Height))
        symbols = sorted(symbols, key=lambda s: s.LevelElev)

        # Генерирую словари
        types_AR_Dict = trueOrderDictCreator(symbols, False)
        types_KR_Dict = trueOrderDictCreator(symbols, True)

        with Transaction(doc, 'КП_Нумерация отверстий') as t:
            t.Start()
            cnt = 0
            mainValueSetter(types_AR_Dict, False)
            mainValueSetter(types_KR_Dict, True)
            t.Commit()
