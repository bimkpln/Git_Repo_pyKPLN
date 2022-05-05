# coding: utf-8

__title__ = "Структурная схема"
__author__ = 'Kapustin Roman'
__doc__ = ''''''

from rpw import *
from System import Guid
from libKPLN.MEP_Elements import allMEP_Elements

from rpw import revit, db, ui
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import script, forms
from rpw.ui.forms import*
from Autodesk.Revit.ApplicationServices import *
from pyrevit.forms import WPFWindow
from System.Windows import Window
from pyrevit.framework import wpf
from System.Collections.Generic import List
from System.Windows.Forms import Form, Button, ComboBox, Label, TextBox
from System.Drawing import Point, Size

import os 
import re
import clr
import webbrowser
clr.AddReference("System")
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
flag = False
textToWrite = "System"
doc = revit.doc
app = doc.Application
uidoc = revit.uidoc
output = script.get_output()
view = doc.ActiveView
lnStart = XYZ(0,0,0)
output = script.get_output()
elemCatTuple = [BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
                    BuiltInCategory.OST_ElectricalEquipment, BuiltInCategory.OST_CableTrayFitting,
                    BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_LightingFixtures,
                    BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_DataDevices,
                    BuiltInCategory.OST_LightingDevices, BuiltInCategory.OST_NurseCallDevices,
                    BuiltInCategory.OST_SecurityDevices, BuiltInCategory.OST_FireAlarmDevices,
                    BuiltInCategory.OST_CommunicationDevices]
ElCircuitCat = BuiltInCategory.OST_ElectricalCircuit
AnotationCat = BuiltInCategory.OST_GenericAnnotation
sysParam = Guid("98b30303-a930-449c-b6c2-11604a0479cb")
# КП_И_Адрес текущий
sysParam = Guid("686ecb99-8191-425a-9cd1-65d3597c3f75")
# КП_О_Позиция
prefParam = Guid("ae8ff999-1f22-4ed7-ad33-61503d85f0f4")
# RBZ_Количество занимаемых адресов
AdressParam = Guid("1fd68ef8-37fe-4896-bf08-c54bf8751ff6")
# КП_И_Адрес предыдущий
ExElParam = Guid("07230eb2-e78d-4b00-9e54-3e4f55bbf21b")
# КП_И_Префикс системы
systemParam = Guid("bd9e5583-305b-495f-8702-de38d6e01b7e")
# КП_И_Адрес устройства- 
ElAdressParam = Guid("02beca6d-e93e-47aa-9cd7-cd5f2867145e")
# КП_И_Код УГО
YGOParam = Guid("493342a5-0f1a-4fe9-b8a2-ad1874a91104")
# КП_И_ID элемента
IDParam= Guid("2aa97f34-a1fc-489b-9371-ee666075e4ee")
class Setup(Form):
        def __init__(self):
            self.__x1 = 30
            self.__y1 = 70
            self.__x2 = 500
            self.__y2 = 70
            self.Size = Size(940,lenMenu)
            self.CenterToScreen()
            self.AutoScroll = True
            self._inintStart()
            self._inintHelp()
            self.dataList1 = list()
            self.dataList2 = list()
        def _inintStart(self):
            self._btnStr = Button()
            self._btnStr.Text = 'Выполнить'
            self._btnStr.Size = Size(400, 40)
            self._btnStr.Location = Point(30, 10)
            self.Controls.Add(self._btnStr)
            self._btnStr.Click += self._result
        def _inintHelp(self):
            self._btnHlp = Button()
            self._btnHlp.Text = 'Помощь'
            self._btnHlp.Size = Size(400, 40)
            self._btnHlp.Location = Point(500, 10)
            self.Controls.Add(self._btnHlp)
            self._btnHlp.Click += self._onHelpClick
        def _addGroup1(self):
            self._cmb1 = TextBox()
            self._cmb1.Size = Size(400, 40)
            self._cmb1.Text = Text1
            self._cmb1.Enabled = False
            self._cmb1.Location = Point(self.__x1, self.__y1)
            self._mark = Label()
            self._mark.Size = Size(10, 10)
            self._mark.Text = '-'
            self._mark.Location = Point(460, self.__y1 + 5)
            self.Controls.Add(self._mark)
            self.Controls.Add(self._cmb1)
            self.__y1 += 35
        def _addGroup2(self):
            self._cmb2 = ComboBox()
            self._cmb2.Size = Size(400, 40)
            self._cmb2.Location = Point(self.__x2, self.__y2)
            for annotation in annotationsDict:
                self._cmb2.Items.Add(annotation)
            self.__y2 += 35
            self.Controls.Add(self._cmb2)
        def _onHelpClick(self, sender, evant_args):
            webbrowser.open('https://kpln.bitrix24.ru/marketplace/app/16/')
        def _result(self, sender, evant_args):
            for ctrl in self.Controls:
                if isinstance(ctrl, ComboBox):
                    self.dataList1.append(ctrl.Text.ToString())
            self.Close()
# ___________________________________________________________
class ParametersForm(Window):
    flag = False
    FlipCheck = False
    Readress = False
    Set = False
    txtBox = "500"
    def __init__(self):
        wpf.LoadComponent(self, 'Z:\\pyRevit\\pyKPLN_MEP\KPLN.extension\\pyKPLN_MEP.tab\\ЭОМ_СС.panel\\Структурная схема.pushbutton\\Form.xaml')
    def _onSetClick(self, sender, evant_args):
        self.Set = True
        self.Close()
    def TextChanged(self, sender, evant_args):
        self.txtBox = self.textBox1.Text
    def _onExtClick(self, sender, evant_args):
        self.flag = False
        self.Close()
    def _onFlipCheckClick(self, sender, evant_args):
        self.FlipCheck = True
        self.Close()
    def _onConsClick(self, sender, evant_args):
        self.flag = True
        self.Close()
    def _onReadressClick(self, sender, evant_args):
        self.Readress = True
        self.Close()
    def _onHelpClick(self, sender, evant_args):
        webbrowser.open('https://articles.it-solution.ru/article/111966/')
window = ParametersForm()
window.ShowDialog()

if window.flag:
    els = FilteredElementCollector(doc).OfCategory(ElCircuitCat).WhereElementIsNotElementType().ToElements()
    elemLines = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).WhereElementIsNotElementType().ToElements()
    lines = dict()
    for line in elemLines:
        lines[line.LineStyle.Name] = line.LineStyle
    sysDict = dict()
    for element in els:
        try:
            Pref = element.get_Parameter(systemParam).AsString()
        except:
            continue
        if Pref in sysDict:
            sysDict[Pref].append(element)
        else:
            sysDict[Pref] = [element]
    components = [rpw.ui.forms.Label('Система для анализа:'),
                                rpw.ui.forms.ComboBox('combobox1', sysDict),
                                rpw.ui.forms.Separator(),
                                rpw.ui.forms.Button('Выбрать')]
    form = FlexForm('Система:', components)
    form.show()
    combobox = form.values.get('combobox1')
    components = [rpw.ui.forms.Label('Выбери тип линий:'),
                            rpw.ui.forms.ComboBox('combobox2', lines),
                            rpw.ui.forms.Separator(),
                            rpw.ui.forms.Button('Выбрать')]
    form = FlexForm('Тип линий:', components)
    form.show()
    Style = form.values.get('combobox2')
    for sys in combobox:
        index = sys.get_Parameter(ElAdressParam).AsDouble()
        if index == 1:
            element = sys.BaseEquipment
    outSysList = []
    sysList = [1,1]
    for iteration in range(128):
        curSys = element.MEPModel.AssignedElectricalSystems
        if not curSys and iteration > 0:
            if element not in outSysList:
                outSysList.append(element)
            break
        sysList = []
        for i in curSys:
            sysList.append(i)
        if element not in outSysList:
            outSysList.append(element)
        elements = sysList[0].Elements
        for el in elements:
            element = el
    point = uidoc.Selection.PickPoint("Pick Point and press Escap to finish")

    AllElements = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    annotations = []
    for element in AllElements:
        if str(element.Category.CategoryType) == "Annotation" and "УГО".lower() in element.Family.Name.lower():
            annotations.append(element)
    with db.Transaction('Структурная схема'):
        FindList = dict()
        scale = float(view.Scale)
        NewY = point.Y
        deltaX = 7*scale/100
        flagFlip = False
        iter = 1
        try:
            length = int(window.txtBox)
        except:
            ui.forms.Alert('Некорректная ширина!', title="Внимение!")
            script.exit()
        for elemSys in outSysList:
            UGOKod = (doc.GetElement(elemSys.GetTypeId()).get_Parameter(YGOParam).AsString())
            Mark = elemSys.get_Parameter(sysParam).AsString()
            if UGOKod in FindList.keys():
                famtype = FindList[UGOKod]
            else:
                for elemAnot in annotations:
                    try:
                        AnottationUGOKod = elemAnot.get_Parameter(YGOParam).AsString()
                    except:
                        continue
                    if AnottationUGOKod == UGOKod:
                        famtype =  elemAnot
                        FindList[UGOKod] = famtype
            newobj = doc.Create.NewFamilyInstance(point,famtype,view)
            newobj.get_Parameter(sysParam).Set(Mark)
            pref = combobox[0].get_Parameter(systemParam).AsString()
            newobj.get_Parameter(systemParam).Set(pref)
            newobj.get_Parameter(IDParam).Set(str(elemSys.Id))
            if iter*deltaX > (scale*length)/304.8 or iter*deltaX < -(scale*length)/304.8: 
                iter = 1
                NewY += 3.5*scale/100
                deltaX = -deltaX
                flagFlip = True
            exPoint = point
            if not flagFlip:
                point = XYZ(point.X+deltaX,NewY,0)
            else:
                point = XYZ(point.X,NewY,0)
                flagFlip = False
            if iter > 0:
                line = Line.CreateBound(exPoint, point)
                DetailCurve = doc.Create.NewDetailCurve(view,line) 
                DetailCurve.LineStyle = Style
            iter+=1
if window.FlipCheck:
    YGOList = []
    selectedReference = ui.Pick.pick_element(msg='Выбери уровни', multiple=True)
    with db.Transaction('Исключение зеркальных УГО'):
        for YgoReference in selectedReference:
            YgoEl = doc.GetElement(YgoReference.id)
            if YgoEl.Category.Name == "Типовые аннотации":
                if YgoEl.Mirrored:
                    Ids=List[ElementId]([YgoEl.Id])
                    PointYgoEl1 = YgoEl.Location.Point
                    PointYgoEl2 = XYZ(PointYgoEl1.X,PointYgoEl1.Y+1,PointYgoEl1.Z)
                    PointYgoEl3 = XYZ(PointYgoEl1.X,PointYgoEl1.Y,PointYgoEl1.Z+1)
                    myLine = Plane.CreateByThreePoints(PointYgoEl1, PointYgoEl2,PointYgoEl3)
                    ElementTransformUtils.MirrorElements(doc, Ids, myLine, False)
if window.Readress:
    CheckList = []
    els = FilteredElementCollector(doc).OfCategory(ElCircuitCat).WhereElementIsNotElementType().ToElements()
    elsAnat = FilteredElementCollector(doc,view.Id).OfCategory(AnotationCat).WhereElementIsNotElementType().ToElements()
    sysDictAnat = dict()
    for elementAnat in elsAnat:
        try:
            PrefAnat = elementAnat.get_Parameter(systemParam).AsString()
            prefSys = PrefAnat
        except:
            continue
        if PrefAnat in sysDictAnat:
            sysDictAnat[PrefAnat].append(elementAnat)
        else:
            sysDictAnat[PrefAnat] = [elementAnat]
    components = [rpw.ui.forms.Label('Система для анализа:'),
                                rpw.ui.forms.ComboBox('combobox3', sysDictAnat),
                                rpw.ui.forms.Separator(),
                                rpw.ui.forms.Button('Выбрать')]
    form = FlexForm('Система:', components)
    form.show()
    comboboxAnat = form.values.get('combobox3')
    FindList = dict()
    AllElements = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    annotations = []
    for element in els:
        prefToCheck = element.get_Parameter(systemParam).AsString()
        if prefToCheck == prefSys:
            sysEls = element.Elements
            ind = element.get_Parameter(ElAdressParam).AsDouble()
            for i in sysEls:
                sysEl = i
            CheckList.append(sysEl.Id)
            if ind == 1:
                sysEls = element.BaseEquipment
                CheckList.append(sysEls.Id)
    for element in AllElements:
        if str(element.Category.CategoryType) == "Annotation" and "УГО".lower() in element.Family.Name.lower():
            annotations.append(element)
    with db.Transaction('Обновление структурной схемы'):
        for AnatEl in comboboxAnat:
            IdOfMother = AnatEl.get_Parameter(IDParam).AsString()
            MotherEl = doc.GetElement(ElementId(int(IdOfMother)))
            MotherElId = MotherEl.Id
            if MotherElId in CheckList:
                CheckList.remove(MotherElId)
            if not MotherEl:
                doc.Delete(AnatEl.Id)
                continue
            AnatUGOKod = (doc.GetElement(AnatEl.GetTypeId()).get_Parameter(YGOParam).AsString())
            AnatMark = AnatEl.get_Parameter(sysParam).AsString()
            MotherMark = MotherEl.get_Parameter(sysParam).AsString()
            MotherUGOKod = doc.GetElement(MotherEl.GetTypeId()).get_Parameter(YGOParam).AsString()
            if AnatUGOKod == MotherUGOKod:
                if AnatMark == MotherMark:
                    prefSys = AnatEl.get_Parameter(systemParam).AsString()
                    if prefSys == "":
                        AnatEl.get_Parameter(systemParam).Set(str(prefSys))
                    continue
                else:
                    AnatEl.get_Parameter(sysParam).Set(MotherMark)
            else:
                point = AnatEl.Location.Point
                if MotherUGOKod in FindList.keys():
                    famtype = FindList[MotherUGOKod]
                else:
                    for elemAnot in annotations:
                        try:
                            AnottationUGOKod = elemAnot.get_Parameter(YGOParam).AsString()
                        except:
                            continue
                        if AnottationUGOKod == MotherUGOKod:
                            famtype =  elemAnot
                            FindList[MotherUGOKod] = famtype
                newobj = doc.Create.NewFamilyInstance(point,famtype,view)
                newobj.get_Parameter(sysParam).Set(MotherMark)
                newobj.get_Parameter(systemParam).Set(str(prefSys))
                newobj.get_Parameter(IDParam).Set(str(MotherEl.Id))
                doc.Delete(AnatEl.Id)
        if CheckList:
            ui.forms.Alert('В цепи новые элементы, выберите точку вставки!', title="Внимение!")
            pt = uidoc.Selection.PickPoint("Pick Point and press Escap to finish")
            for newElId in CheckList:
                newEl = doc.GetElement(newElId)
                NewUGOKod = doc.GetElement(newEl.GetTypeId()).get_Parameter(YGOParam).AsString()
                NewElMark = newEl.get_Parameter(sysParam).AsString()
                if NewUGOKod in FindList.keys():
                    famtype = FindList[NewUGOKod]
                else:
                    for elemAnot in annotations:
                        try:
                            AnottationUGOKod = elemAnot.get_Parameter(YGOParam).AsString()
                        except:
                            continue
                        if AnottationUGOKod == NewUGOKod:
                            famtype =  elemAnot
                            FindList[NewUGOKod] = famtype
                newobj = doc.Create.NewFamilyInstance(pt,famtype,view)
                newobj.get_Parameter(sysParam).Set(NewElMark)
                newobj.get_Parameter(systemParam).Set(str(prefSys))
                newobj.get_Parameter(IDParam).Set(str(newElId))
                pt = XYZ(pt.X + 0.05,pt.Y,pt.Z)
if window.Set:
    AllElements = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    annotationsDict = dict()
    for element in AllElements:
        if str(element.Category.CategoryType) == "Annotation" and "УГО".lower() in element.Family.Name.lower():
            elTypeName =  element.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            try:
                element.get_Parameter(YGOParam).AsValueString()
            except:
                continue
            if elTypeName not in annotationsDict:
                annotationsDict[elTypeName] = element

    els = FilteredElementCollector(doc).OfCategory(ElCircuitCat).WhereElementIsNotElementType().ToElements()
    sysDict = dict()
    for element in els:
        try:
            Pref = element.get_Parameter(systemParam).AsString()
        except:
            continue
        if Pref in sysDict:
            sysDict[Pref].append(element)
        else:
            sysDict[Pref] = [element]
    components = [rpw.ui.forms.Label('Система для анализа:'),
                                rpw.ui.forms.ComboBox('combobox3', sysDict),
                                rpw.ui.forms.Separator(),
                                rpw.ui.forms.Button('Выбрать')]
    form = FlexForm('Система:', components)
    form.show()
    combobox = form.values.get('combobox3')
    unequiEl = []
    for element in combobox:
        elBaseId = element.BaseEquipment.Symbol.Id
        elEls = element.Elements
        for i in elEls:
            elElId = i.Symbol.Id
        if elBaseId not in unequiEl:
            unequiEl.append(elBaseId)
        if elElId not in unequiEl:
            unequiEl.append(elElId)
    n = len(unequiEl)
    lenMenu = 100 + (n+1)*35
    win = Setup()
    for i in unequiEl:
        Text1 = doc.GetElement(i).get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        win._addGroup1()
        win._addGroup2()
    win.ShowDialog()
    itemsList = win.dataList1
    
    with db.Transaction('Настройка уго'):
        for i in range(len(win.dataList1)):
            elementType = doc.GetElement(unequiEl[i])
            elementYGO = annotationsDict[itemsList[i]]
            YgoKod = elementType.get_Parameter(YGOParam).Set(itemsList[i])
            YgoKodAnat = elementYGO.get_Parameter(YGOParam).Set(itemsList[i])