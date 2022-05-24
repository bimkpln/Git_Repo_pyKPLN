# coding: utf-8

__title__ = "ЭлЦепи ЭОМ"
__author__ = 'Kapustin Roman'
__doc__ = ''''''

from rpw import *
from System import Guid
from Autodesk.Revit.DB import *
from pyrevit import script, forms
from rpw.ui.forms import*
from Autodesk.Revit.DB.Structure import StructuralType
from System.Windows import Window
from pyrevit.forms import WPFWindow
from pyrevit.framework import wpf
from System.Windows.Forms import*
from System.Drawing import Point, Size
import webbrowser
doc = revit.doc
app = doc.Application
view = doc.ActiveView
WorkPlane = doc.ActiveView.GenLevel
systems = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ElectricalCircuit).WhereElementIsNotElementType().ToElements()
PanelParam = Guid('d6ce883f-81d1-4d81-a4e7-fb897c10944c')
SysNumParam = Guid('85c1b5f7-8b83-4533-8fc4-c19f72de752b')
IdParam = Guid('2aa97f34-a1fc-489b-9371-ee666075e4ee')
typeCabelYGOParam = Guid('bd1cc219-f982-4da5-80aa-76145e6b814c')
NameParam = Guid('e6e0f5cd-3e26-485b-9342-23882b20eb43')
TypeOfPipe = Guid('328ab0b4-dda7-4c74-9d25-bd7134cafad7')
SysNameParam = BuiltInParameter.RBS_ELEC_CIRCUIT_WIRE_TYPE_PARAM
DeltaParam = Guid("a37d7ee0-23b1-4d2c-8e8e-ecd2e5a36101")
allFam = FilteredElementCollector(doc).OfClass(Family).ToElements()
cabelFam = None

class Setup(Form):
    def __init__(self):
        self.num = 0
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
        self.num  += 1
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
        self._cmb2.Text = str(list(Elsyss)[self.num-1].get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString())
        self._cmb2.Size = Size(400, 40)
        self._cmb2.Location = Point(self.__x2, self.__y2)
        for i in Elsyss:
            self._cmb2.Items.Add(i.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString())
        self.__y2 += 35
        self.Controls.Add(self._cmb2)
    def _onHelpClick(self, sender, evant_args):
        webbrowser.open('https://kpln.bitrix24.ru/marketplace/app/16/')
    def _result(self, sender, evant_args):
        for ctrl in self.Controls:
            if isinstance(ctrl, ComboBox):
                self.dataList1.append(ctrl.Text.ToString())
        self.Close()

class ParametersForm(WPFWindow):
    flag = False
    Readress = False
    ReadPaint = False
    ReGroup = False
    AddFlag = False
    track = False
    trackByLine = False 

    def __init__(self):
        WPFWindow.__init__(self, 'Form.xaml')

    def _onExtClick(self, sender, evant_args):
        self.Close()

    def _onNewClick(self, sender, evant_args):
        self.flag = True
        self.Close()

    def _onTrackClick(self, sender, evant_args):
        self.track = True
        self.Close()

    def _onLineTrackClick(self, sender, evant_args):
        self.trackByLine = True
        self.Close()

    def _onAddClick(self, sender, evant_args):
        self.AddFlag = True
        self.Close()

    def _onPaintClick(self, sender, evant_args):
        self.ReadPaint = True
        self.Close()

    def _onReadressClick(self, sender, evant_args):
        self.Readress = True
        self.Close()

    def _onReNameGroup(self, sender, evant_args):
        self.ReGroup = True
        self.Close()

    def _onHelpClick(self, sender, evant_args):
        webbrowser.open('http://moodle.stinproject.local/mod/book/view.php?id=396&chapterid=542')
window = ParametersForm()
window.ShowDialog()

if window.flag: #Построение кабеля по траектории
    #1_______Поиск семейства проводов в проекте___________
    for fam in allFam:
        if "001_Кабель ЭОМ" in fam.Name:
            ids = fam.GetFamilySymbolIds()
            for i in ids:
                cabelFam =  doc.GetElement(i)
                break
            break
    if not cabelFam:
        ui.forms.Alert('Семейство кабеля не подгружено в проет!', title="Внимание!")
        script.exit()
    #1____________________________________________________
    #2_______Выбор панели у которой берется система_______
    while True:
        ui.forms.Alert('Выбери Панель!', title="Внимание!")
        selectedReference1 = ui.Pick.pick_element(msg='')
        PickedEl1 = doc.GetElement(selectedReference1.id)
        PickedSysDict = dict()
        PickedSys = PickedEl1.MEPModel.ElectricalSystems
        for i in PickedSys:
            if i.BaseEquipment.Id == PickedEl1.Id:
                PickedSysDict[i.Name] = i
        if PickedSysDict:
            break
        else:
            ui.forms.Alert('Выбрана не панель!', title="Внимание!")
    components = [rpw.ui.forms.Label('Система для анализа:'),
                            rpw.ui.forms.ComboBox('combobox3', PickedSysDict),
                            rpw.ui.forms.Separator(),
                            rpw.ui.forms.Button('Выбрать')]
    form = FlexForm('Система:', components)
    form.show()
    system = form.values.get('combobox3')
    
    PanelName = system.PanelName
    CircuitNumber = system.CircuitNumber
    pathPoints = system.GetCircuitPath()
    SysID = system.Id
    Base = system.BaseEquipment.Id
    #2____________________________________________________
    #3_______Удалить либо оставить существующие в системе провода_______
    lineInfoList = []
    nextFlag = False
    commands1= [CommandLink('Удалить предыдущие!', return_value='Dell'),
                CommandLink('Оставить предыдущие!', return_value='Add')]
    dialog1 = TaskDialog('Выберите дальнейшее действие:',
                commands=commands1,
                    show_close=True)
    SysElSet = system.Elements
    dialog_out1 = None
    #3__________________________________________________________________
    #4_______Определение корректного уровня для размещения проводов_____
    curLvlEl = None
    curLvl= None
    allLvl = FilteredElementCollector(doc).OfClass(Level).ToElements()
    minDelta = abs(0 - allLvl[0].Elevation)
    WorkPlane = allLvl[0]
    for lvl in allLvl:
        delta = abs(0 - lvl.Elevation)
        if delta < minDelta:
            WorkPlane = lvl
    #4__________________________________________________________________
    #3_______Удалить либо оставить существующие в системе провода_______
    for i in SysElSet:
        if "001_Кабель ЭОМ" in i.Name:
            ui.forms.Alert('В цепи уже есть провода!', title="Внимание!")
            dialog_out1 = dialog1.show()
            break
    CabelList = []
    delList = []
    if dialog_out1 == "Dell":
        with db.Transaction("Удвление старых проводов:"):
            for i in SysElSet:
                if "001_Кабель ЭОМ" in i.Name:
                    doc.Delete(i.Id)
    else:
        for i in SysElSet:
            if "001_Кабель ЭОМ" in i.Name:
                CabelList.append(i)
    #3__________________________________________________________________
    #5_______Выбор типа прокладки для провода (Визуальное отображение)_______
    commands= [CommandLink('Без типа прокладки', return_value='NoType'),
                CommandLink('Прокладка в гофротрубе', return_value='Gofr'),
                CommandLink('Прокладка в кабельном лотке', return_value='kabLot')]
    dialog = TaskDialog('Выберите тип прокладки:',
                commands=commands,
                    show_close=True)
    dialog_out = dialog.show()
    if not dialog_out:
        script.exit()
    #5_______________________________________________________________________
    with db.Transaction("Построение цепи:"):
        for i in range(len(pathPoints)):
            #6_______Исключение дублирования точек (первичное)_______________
            lnStart1 =  XYZ(pathPoints[i].X,pathPoints[i].Y,pathPoints[i].Z)
            try:
                lnEnd1 = XYZ(pathPoints[i+1].X,pathPoints[i+1].Y,pathPoints[i+1].Z)
                if not nextFlag:
                    deltaZ = abs(pathPoints[i].Z - pathPoints[i+1].Z)
                else:
                    deltaZ += abs(pathPoints[i].Z - pathPoints[i+1].Z)
                if lnEnd1 < lnStart1:
                    lnEnd = lnStart1
                    lnStart = lnEnd1
                else:
                    lnEnd = lnEnd1
                    lnStart = lnStart1
                if str(lnStart)+'|'+str(lnEnd) in lineInfoList:
                    continue
                lineInfoList.append(str(lnStart)+'|'+str(lnEnd))
                lineInfoList.append(str(lnEnd)+'|'+str(lnStart))
            except:
                break
            #6_______________________________________________________________
            #7_______Создание провода по точкам и задача параметров для него_______________
            try:
                line = Line.CreateBound(lnStart, lnEnd)
                newobj = doc.Create.NewFamilyInstance(line,cabelFam,WorkPlane,StructuralType.NonStructural)   
                newobj.get_Parameter(PanelParam).Set(str(PanelName))
                newobj.get_Parameter(SysNumParam).Set(str(CircuitNumber))
                newobj.get_Parameter(IdParam).Set(str(SysID))
                newobj.get_Parameter(DeltaParam).Set(round(deltaZ,1))
                newobj.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(pathPoints[i+1].Z - 10/304.8)
                nextFlag = False
                try:
                    cabelPipeName = system.LookupParameter('Наименование трубы_Электрические цепи').AsString()
                except:
                    cabelPipeName = None
                try:
                    cabelName = system.get_Parameter(SysNameParam).AsValueString()
                except:
                    cabelName = None
                if cabelName:
                    newobj.get_Parameter(NameParam).Set(cabelName)
                if cabelPipeName:
                    newobj.get_Parameter(TypeOfPipe).Set(cabelPipeName)
                if dialog_out == 'NoType':
                    newobj.get_Parameter(typeCabelYGOParam).Set(0)
                if dialog_out == 'Gofr':
                    newobj.get_Parameter(typeCabelYGOParam).Set(1)
                if dialog_out == 'kabLot': 
                    newobj.get_Parameter(typeCabelYGOParam).Set(2)
                ElSet = ElementSet()
                ElSet.Insert(newobj)
                CabelList.append(newobj)
                try:
                    system.AddToCircuit(ElSet)
                except:
                    continue
            except:
                nextFlag = True
                continue
            #7_____________________________________________________________________________________________
    #8_______Исключение дублирования точек (вторичное, удаление линий которые дублировались)_______________
    for Cabel1 in CabelList:
        for Cabel2 in CabelList:
            if Cabel1.Id != Cabel2.Id:
                Cabel1Z = Cabel1.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble()
                Cabel2Z = Cabel2.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsDouble()
                if Cabel1Z == Cabel2Z:
                    if round(Cabel1.Location.Curve.Tessellate()[0].X) == round(Cabel2.Location.Curve.Tessellate()[1].X):
                        if round(Cabel1.Location.Curve.Tessellate()[0].Y) == round(Cabel2.Location.Curve.Tessellate()[1].Y):
                            if round(Cabel1.Location.Curve.Tessellate()[1].X) == round(Cabel2.Location.Curve.Tessellate()[0].X):
                                if round(Cabel1.Location.Curve.Tessellate()[1].Y) == round(Cabel2.Location.Curve.Tessellate()[0].Y):
                                    if Cabel2 not in delList and  Cabel1 not in delList:
                                        delList.append(Cabel2)
                                        CabelList.remove(Cabel2)
                                        CabelList.remove(Cabel1)
                                        break
    with db.Transaction("Удаление дублирования:"):
        for i in delList:
            doc.Delete(i.Id)
#8_______________________________________________________________________________________________________________
#________________________________________________________________________________________________________________
if window.ReadPaint:  #Обновить траекторию провода
    #1_______Поиск семейства проводов в проекте___________
    for fam in allFam:
        if "001_Кабель ЭОМ" in fam.Name:
            ids = fam.GetFamilySymbolIds()
            for i in ids:
                cabelFam =  doc.GetElement(i)
                break
            break
    if not cabelFam:
        ui.forms.Alert('Семейство кабеля не подгружено в проет!', title="Внимание!")
        script.exit()
    #1____________________________________________________
    #2_______Выбор панели у которой берется система_______
    while True:
        ui.forms.Alert('Выбери Панель!', title="Внимание!")
        selectedReference1 = ui.Pick.pick_element(msg='')
        PickedEl1 = doc.GetElement(selectedReference1.id)
        PickedSysDict = dict()
        PickedSys = PickedEl1.MEPModel.ElectricalSystems
        for i in PickedSys:
            if i.BaseEquipment.Id == PickedEl1.Id:
                PickedSysDict[i.Name] = i
        if PickedSysDict:
            break
        else:
            ui.forms.Alert('Выбрана не панель!', title="Внимание!")
    components = [rpw.ui.forms.Label('Система для анализа:'),
                                rpw.ui.forms.ComboBox('combobox3', PickedSysDict),
                                rpw.ui.forms.Separator(),
                                rpw.ui.forms.Button('Выбрать')]
    form = FlexForm('Система:', components)
    form.show()
    CurSys = form.values.get('combobox3')
    pathPoints = CurSys.GetCircuitPath()
    cabelsList = []
    cabelsPointsList = []
    SysElSet = CurSys.Elements
    #2____________________________________________________
    #3_______Создание списка из точек у уже существующих в системе проводов для того, чтобы проверить, есть ли изменения в траектории_______
    for cabel in SysElSet:
        if "001_Кабель ЭОМ" in cabel.Name:
            if cabel not in cabelsList:
                cabelsPointsList.append([str(XYZ(round(cabel.Location.Curve.Tessellate()[0].X,5),round(cabel.Location.Curve.Tessellate()[0].Y,5),0)),str(XYZ(round(cabel.Location.Curve.Tessellate()[1].X,5),round(cabel.Location.Curve.Tessellate()[1].Y,5),0))])
                cabelsPointsList.append([str(XYZ(round(cabel.Location.Curve.Tessellate()[1].X,5),round(cabel.Location.Curve.Tessellate()[1].Y,5),0)),str(XYZ(round(cabel.Location.Curve.Tessellate()[0].X,5),round(cabel.Location.Curve.Tessellate()[0].Y,5),0))])
                cabelsList.append(cabel)
                cabelsList.append(cabel)
    PanelName = CurSys.PanelName
    CircuitNumber = CurSys.CircuitNumber
    pathPoints = CurSys.GetCircuitPath()
    SysID = CurSys.Id      
    #3________________________________________________________________________________________________________________________________________
    #4_______Выбор способа прокладки провода (Визуальное отображение)_______
    NotToDel = []
    commands= [CommandLink('Без типа прокладки', return_value='NoType'),
                CommandLink('Прокладка в гофротрубе', return_value='Gofr'),
                CommandLink('Прокладка в кабельном лотке', return_value='kabLot')]
    dialog = TaskDialog('Выберите тип прокладки:',
                commands=commands,
                    show_close=True)
    dialog_out = dialog.show()
    #4______________________________________________________________________
    with db.Transaction("Построение цепи:"):
        #5_______Определение точек траектории и исключение дублирования точек (первичное)_______
        iteration = 0
        lineInfoList = [] 
        newFlag = False
        for i in range(len(pathPoints)):
            lnStart1 =  XYZ(round(pathPoints[i].X,5),round(pathPoints[i].Y,5),0)
            try:
                lnEnd1 = XYZ(round(pathPoints[i+1].X,5),round(pathPoints[i+1].Y,5),0)
                if lnEnd1 < lnStart1:
                    lnEnd = lnStart1
                    lnStart = lnEnd1
                else:
                    lnEnd = lnEnd1
                    lnStart = lnStart1

                if [str(lnStart),str(lnEnd)] in cabelsPointsList:   
                    curCabInd = cabelsPointsList.index([str(lnStart),str(lnEnd)])
                    CurCab = cabelsList[curCabInd]
                    NotToDel.append(CurCab)
                    continue
                if str(lnStart)+'|'+str(lnEnd) in lineInfoList:  #Если точка уже есть в точке имеющихся в траектории проводов, то continue
                    continue
                lineInfoList.append(str(lnStart)+'|'+str(lnEnd))
                lineInfoList.append(str(lnEnd)+'|'+str(lnStart))
            except(Exception) as e:
                break
            #5__________________________________________________________________________________________________
            #6_______Создание проводов, которые имеют новые точки (старые остаются только в том случае, если их точки не менялись)_______
            try:
                line = Line.CreateBound(lnStart, lnEnd)
                newobj = doc.Create.NewFamilyInstance(line,cabelFam,WorkPlane,StructuralType.NonStructural)   
                newobj.get_Parameter(PanelParam).Set(str(PanelName))
                newobj.get_Parameter(SysNumParam).Set(str(CircuitNumber))
                newobj.get_Parameter(IdParam).Set(str(SysID))
                newobj.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(pathPoints[i+1].Z - 10/304.8)
                try:
                    cabelPipeName = CurSys.LookupParameter('Наименование трубы_Электрические цепи').AsString()
                except:
                    cabelPipeName = None
                try:
                    cabelName = CurSys.get_Parameter(SysNameParam).AsValueString()
                except:
                    cabelName = None
                if cabelName:
                    newobj.get_Parameter(NameParam).Set(cabelName)
                if cabelPipeName:
                    newobj.get_Parameter(TypeOfPipe).Set(cabelPipeName)
                if dialog_out == 'NoType':
                    newobj.get_Parameter(typeCabelYGOParam).Set(0)
                if dialog_out == 'Gofr':
                    newobj.get_Parameter(typeCabelYGOParam).Set(1)
                if dialog_out == 'kabLot': 
                    newobj.get_Parameter(typeCabelYGOParam).Set(2)
                ElSet = ElementSet()
                ElSet.Insert(newobj)
                try:
                    CurSys.AddToCircuit(ElSet)
                except:
                    continue    
            except:
                continue
        for i in cabelsList:
            if i not in NotToDel:
                try:
                    doc.Delete(i.Id)
                except:
                    continue
            #6_____________________________________________________________________________________________________________________________________
#__________________________________________________________________________________________________________________________________________________
if window.Readress:  #Обновить информацию о типе провода
    ui.forms.Alert('Тип кабеля будет перенесен из цепи в провод!', title="Внимание!")
    #1_______Перезаписывает тип кабеля из системы в провода, которые подключены к системе_______
    with db.Transaction("Обновление информации о кабеле:"):
        for sys in systems:
            cabelDict = []
            SysElSet = sys.Elements
            cabelName = sys.get_Parameter(SysNameParam).AsValueString()
            for i in SysElSet:
                if "001_Кабель ЭОМ" in i.Name:
                    cabelDict.append(i)
            for cabel in cabelDict:
                cabel.get_Parameter(NameParam).Set(cabelName)
    #1___________________________________________________________________________________________
#__________________________________________________________________________________________________________________________________________________
if window.ReGroup: #Обновить порядок групп
    #1_______Создание класса системы (объект группа системы)______
    class System:
        CircuitPath = None
        SysName = None
        SysId = None
        Sys = None
        Cabels = None
    #1_____________________________________________________________
    #2______Выбирается панель у которой берется система____________
    ui.forms.Alert('Выберите панель!', title="Внимание!")
    selectedReference = ui.Pick.pick_element(msg='Выбери панель!')
    PickedEl = doc.GetElement(selectedReference.id)
    mepModel = PickedEl.MEPModel
    Elsyss = mepModel.ElectricalSystems
    n = 0
    for i in Elsyss:
        if i.BaseEquipment.Id != PickedEl.Id:
            continue
        n += 1
    lenMenu = 100 + (n+1)*35
    win = Setup()
    #2_____________________________________________________________
    #3______Формирование менюшки с выбором групп___________________
    ElsyssDict = dict()
    for i in Elsyss:
        if i.BaseEquipment.Id != PickedEl.Id:
            continue
        #3.1______создание объекта система с записью в него траектории и информации о системе____________
        numOfsys = i.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString()
        system = System()
        system.Sys = i
        system.SysId = i.Id
        system.CircuitPath = i.GetCircuitPath()
        system.SysName = i.Name
        ElsyssDict[numOfsys] = system
        #3.1______________________________________________________________________________________________
    iter = 0
    for Elsys in Elsyss:
        iter+=1
        Text1 = str(iter)
        win._addGroup1()
        win._addGroup2()
    win.ShowDialog()
    itemsList = win.dataList1
    #3_____________________________________________________________
    #4______Отключение всех групп от панели_________________________
    with db.Transaction("Отключение:"):
        for sysObjKey in ElsyssDict.keys():
            sysObj = ElsyssDict[sysObjKey]
            curSys = sysObj.Sys
            curSys.DisconnectPanel()
    #4______________________________________________________________
    #5______Подключение групп к панели в последовательности, заданной выходными данными менюшки_________________________
    try:
        for i in itemsList:
            curSysObj = ElsyssDict[i]
            curSysName = curSysObj.SysName
            curSys = curSysObj.Sys
            curCircuitPath = curSysObj.CircuitPath
            with db.Transaction("Переподключение: {}".format(curSysName)):
                curSys.SelectPanel(PickedEl)
        #5_______________________________________________________________________________________________________________
        #6______Запись в цепь траектории из объекта системы_________________________
        for i in itemsList:
            curSysObj = ElsyssDict[i]
            curSysName = curSysObj.SysName
            curSys = curSysObj.Sys
            curCircuitPath = curSysObj.CircuitPath
            with db.Transaction("Настройка трассировки: {}".format(curSysName)):
                curSys.SetCircuitPath(curCircuitPath)
    except:
        ui.forms.Alert('Не корректный порядок', title="Внимание!")
        script.exit()
        #6____________________________________________________________________________
    cabels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
    curCabelsDict = dict()
        
    #7______Запись остальных параметров в провода (новое значение номера цепи)_________________________
    with db.Transaction("Переименование номера цепи:"):
        for i in itemsList:
            curSysObj = ElsyssDict[i]
            curSys = curSysObj.Sys
            curSysObj.SysName = curSys.get_Parameter(BuiltInParameter.RBS_ELEC_CIRCUIT_NUMBER).AsString()
            CurName = curSysObj.SysName
            curSysEls = curSys.Elements
            curCabelsListItems = []
            for i in curSysEls:
                if "001_Кабель ЭОМ" in i.Name:
                    curCabelsListItems.append(i)
            if curCabelsListItems == []:
                for cabel in cabels:
                    if "001_Кабель ЭОМ" in cabel.Name:
                        if str(curSys.Id) == cabel.get_Parameter(IdParam).AsString():
                            curCabelsListItems.append(cabel)
            curSysObj.Cabels = curCabelsListItems
            for j in curSysObj.Cabels:
                j.get_Parameter(SysNumParam).Set(CurName)
    #7____________________________________________________________________________________________________
#_____________________________________________________________________________________________________________________________________________________________
if window.AddFlag:  #Объеденение проводов в 1 линию
    while True:
        try:
            selectedReference = ui.Pick.pick_element(msg='Выбери провода, которые надо ровнять!',  multiple=True)
            PickedEl1 = doc.GetElement(selectedReference[0].id) 
            for i in selectedReference:
                if PickedEl1.Location.Curve.Length <  doc.GetElement(i.id).Location.Curve.Length:
                    PickedEl1 = doc.GetElement(i.id)
            lnStart = PickedEl1.Location.Curve.Tessellate()[0]
            lnEnd = PickedEl1.Location.Curve.Tessellate()[1]
            if lnStart.X != lnEnd.X and lnStart.Y != lnEnd.Y:
                if abs(lnStart.X - lnEnd.X) < 2: 
                    lnEnd = XYZ(lnStart.X,lnEnd.Y,lnEnd.Z)
                elif abs(lnStart.Y - lnEnd.Y) < 2:
                    lnEnd = XYZ(lnEnd.X,lnStart.Y,lnEnd.Z)
            line = Line.CreateBound(lnStart, lnEnd)
            with db.Transaction("Объединение проводов:"):
                for i in selectedReference:
                    PickedEl = doc.GetElement(i.id)
                    if PickedEl.Category.Name == "Обобщенные модели":
                            PickedEl.Location.Curve = line 
        except:
            ui.forms.Alert('Завершено!', title="Внимание!")
            break
#_____________________________________________________________________________________________________________________________________________________________

if window.trackByLine: #Плагин по прокладке траектории по точкам (по линии детализации)
    #1______Выбирается панель у которой берется система_________________________
    ui.forms.Alert('Выбери Панель!', title="Внимание!")
    selectedReference1 = ui.Pick.pick_element(msg='')
    PickedEl1 = doc.GetElement(selectedReference1.id)
    PickedSysDict = dict()
    PickedSys = PickedEl1.MEPModel.ElectricalSystems
    for i in PickedSys:
        PickedSysDict[i.Name] = i
    components = [rpw.ui.forms.Label('Система для анализа:'),
                                rpw.ui.forms.ComboBox('combobox3', PickedSysDict),
                                rpw.ui.forms.Separator(),
                                rpw.ui.forms.Button('Выбрать')]
    form = FlexForm('Система:', components)
    form.show()
    combobox = form.values.get('combobox3')
    sys = combobox
    pointsList = []
    PanelLoc = PickedEl1.Location.Point
    firstPointZ = PanelLoc.Z
    pointsList.append(PanelLoc)
    iter = 1
    point2 = XYZ(0,0,0)
    point = PanelLoc
    #1______________________________________________________________________________
    #2______Бесконеный цикл (выбор трактории по точкам на отметке, выбор элемента на след отметке и тд)_________________________
    while True:
        try:
            ui.forms.Alert('Выберите линии траектории на {} отметке'.format(iter), title="Внимание!")
            point11 = uidoc.Selection.PickPoint("Выберите точку 1")
            while True:
                try:
                    point22 = uidoc.Selection.PickPoint("Выберите точку 2")
                    pointsList.append(XYZ(point11.X,point11.Y,firstPointZ))
                    point11 = point22
                except:
                    pointsList.append(XYZ(point11.X,point11.Y,firstPointZ))
                    break
            ui.forms.Alert('Выбери элемент на {} отметке'.format(iter), title="Внимание!")
            selectedReference = ui.Pick.pick_element(msg='')
            PickedEl = doc.GetElement(selectedReference.id)

            firstPointZ = PickedEl.Location.Point.Z
            point2 = XYZ(point11.X,point11.Y,firstPointZ)
            pointsList.append(point2)
            iter += 1
        except:
            break
    if not sys:
        ui.forms.Alert('Выбери элемент у которого взять систему!', title="Внимание!")
        selectedReference = ui.Pick.pick_element(msg='')
        PickedEl = doc.GetElement(selectedReference.id)
        syss = PickedEl.MEPModel.ElectricalSystems
        for i in syss:
            sys = i
    with db.Transaction("Создание траектории:"):
        sys.SetCircuitPath(pointsList)
        ui.forms.Alert('Завершено!', title="Завершено!")
#2_______________________________________________________________________________________________________________________________
if window.track:  #Плагин по прокладке траектории выбирая последовательно элементы
    ui.forms.Alert('Выбери Панель!', title="Внимание!")
    selectedReference1 = ui.Pick.pick_element(msg='')
    PickedEl1 = doc.GetElement(selectedReference1.id)
    pointsList = []
    PanelLoc = PickedEl1.Location.Point
    pointsList.append(PanelLoc)
    iter = 1
    ui.forms.Alert('Выбери элементы на {} отметке'.format(iter), title="Внимание!")

    selectedReference = ui.Pick.pick_element(msg='')
    PanelLocZ = PanelLoc.Z
    firstPointZ = doc.GetElement(selectedReference.id).Location.Point.Z
    if PanelLocZ != firstPointZ:
        pointFirst = XYZ(PanelLoc.X,PanelLoc.Y,firstPointZ)
        pointsList.append(pointFirst)
    PickedEl = doc.GetElement(selectedReference.id)
    syss = PickedEl.MEPModel.ElectricalSystems
    for i in syss:
        sys = i
    Loc = PickedEl.Location.Point
    LocZ1 = Loc.Z
    LocZ2 = pointsList[-1].Z
    if LocZ1 != LocZ2:
        point3 = XYZ(Loc.X,Loc.Y,LocZ2)
        pointsList.append(point3)
    pointsList.append(Loc)
    sys2 = sys

    while True:
        try:
            selectedReference = ui.Pick.pick_element(msg='')
            PickedEl = doc.GetElement(selectedReference.id)
            syss = PickedEl.MEPModel.ElectricalSystems
            for i in syss:
                sys = i
            if sys.Id != sys2.Id:
                script.exit()
            Loc = PickedEl.Location.Point
            LocZ1 = Loc.Z
            LocZ2 = pointsList[-1].Z
            if LocZ1 != LocZ2:
                point3 = XYZ(Loc.X,Loc.Y,LocZ2)
                pointsList.append(point3)
            pointsList.append(Loc)
            sys2 = sys
        except:
            iter += 1 
            break
    while True:
        try:
            ui.forms.Alert('Выбери точку перепада высоты', title="Внимание!")
            LocZ = Loc.Z
            point = uidoc.Selection.PickPoint("Выберите точку перехода высоты")
            point1 = XYZ(point.X,point.Y,LocZ)
            pointsList.append(point1)
            
            ui.forms.Alert('Выбери элементы на {} отметке'.format(iter), title="Внимание!")
            selectedReference = ui.Pick.pick_element(msg='')
            PickedEl = doc.GetElement(selectedReference.id)
            syss = PickedEl.MEPModel.ElectricalSystems
            for i in syss:
                sys = i
            point2 = XYZ(point.X,point.Y,PickedEl.Location.Point.Z)
            pointsList.append(point2)
            if sys.Id != sys2.Id:
                script.exit()
            Loc = PickedEl.Location.Point
            LocZ1 = Loc.Z
            LocZ2 = pointsList[-1].Z
            if LocZ1 != LocZ2:
                point3 = XYZ(Loc.X,Loc.Y,LocZ2)
                pointsList.append(point3)
            pointsList.append(Loc)
            sys2 = sys
            while True:
                try:
                    selectedReference = ui.Pick.pick_element(msg='')
                    PickedEl = doc.GetElement(selectedReference.id)
                    syss = PickedEl.MEPModel.ElectricalSystems
                    for i in syss:
                        sys = i
                    if sys.Id != sys2.Id:
                        script.exit()
                    Loc = PickedEl.Location.Point
                    LocZ1 = Loc.Z
                    LocZ2 = pointsList[-1].Z
                    if LocZ1 != LocZ2:
                        point3 = XYZ(Loc.X,Loc.Y,LocZ2)
                        pointsList.append(point3)
                    pointsList.append(Loc)
                    sys2 = sys
                    
                except:
                    iter += 1 
                    break
        except:
            break
    with db.Transaction("Создание траектории:"):
        sys.SetCircuitPath(pointsList)
        ui.forms.Alert('Завершено!', title="Завершено!")
