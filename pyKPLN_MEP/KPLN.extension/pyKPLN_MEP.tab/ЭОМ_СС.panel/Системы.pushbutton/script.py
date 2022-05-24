# coding: utf-8

__title__ = "Слаботочные системы"
__author__ = 'Kapustin Roman'
__doc__ = ''''''

from rpw import *
from System import Guid
from libKPLN.MEP_Elements import allMEP_Elements
from System.Windows.Forms import Form, Button, ComboBox, Label
from System.Drawing import Point, Size
import webbrowser
import clr
import math
from Autodesk.Revit.DB import *
import re
from pyrevit import script, forms
from rpw.ui.forms import*
from Autodesk.Revit.ApplicationServices import *
from pyrevit.forms import WPFWindow
from System.Windows import Window
from pyrevit.framework import wpf
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.DB import MEPSystem
from Autodesk.Revit import DB
import AddEl
import Dell
import Readress
import ReadressHelp

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
# clr.AddReference('ProtoGeometry')
# clr.AddReference("RevitServices")
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
# КП_И_Адрес текущий
sysParam = Guid("686ecb99-8191-425a-9cd1-65d3597c3f75")
# КП_О_Позиция
prefParam = Guid("ae8ff999-1f22-4ed7-ad33-61503d85f0f4")
# КП_И_Количество занимаемых адресов
AdressParam = Guid("1fd68ef8-37fe-4896-bf08-c54bf8751ff6")
# КП_И_Длина линии
RangeParam = Guid("e6ccadc1-1d18-4168-aa21-7b3905835f0e")
# КП_И_Адрес предыдущий
ExElParam = Guid("07230eb2-e78d-4b00-9e54-3e4f55bbf21b")
# КП_И_Префикс системы
systemParam = Guid("bd9e5583-305b-495f-8702-de38d6e01b7e")
# КП_И_Адрес устройства- 
ElAdressParam = Guid("02beca6d-e93e-47aa-9cd7-cd5f2867145e")

sysTypes = [ ElectricalSystemType.FireAlarm, ElectricalSystemType.PowerCircuit,ElectricalSystemType.Data,ElectricalSystemType.Telephone,ElectricalSystemType.Security,ElectricalSystemType.NurseCall,ElectricalSystemType.Controls]
elemLines = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).WhereElementIsNotElementType().ToElements()
lines = dict()
if len(elemLines) < 1:
    ui.forms.Alert("В проекте нет линий!", title="Внимание")
    script.exit()
for line in elemLines:
    lines[line.LineStyle.Name] = line.LineStyle
components = [Label('Выбери тип линий:'),
                            ComboBox('combobox1', lines),
                            Separator(),
                            Button('Выбрать')]
form = FlexForm('Тип линий:', components)
class ParametersForm(WPFWindow):
    flag = True
    flagParal = False
    flagcons = True
    Mood = 'cons'
    flagLinePicked = False
    flagCheckBox1 = False
    flagNew = False
    dellFlag = False
    addFlag = False
    reculcFlag = False
    readressFlag =False
    readressHelpFlag = False
    txtBox1 = "1"
    txtBox2 = "АПС_1.1"
    TextAdress = 'Не определено'
    Style = 0
    SysType = ElectricalSystemType.FireAlarm

    def __init__(self):
        WPFWindow.__init__(self, 'Form.xaml')

    def TextChanged1(self, sender, evant_args):
        self.txtBox1 = self.textBox1.Text

    def TextChanged2(self, sender, evant_args):
        self.txtBox2 = self.textBox2.Text
    def TextChanged_adress(self, sender, evant_args):
        self.TextAdress = self.textBox_adress.Text

    def _onNewClick(self, sender, evant_args):
        self.flagNew = True
        self.Close()

    def _onExtClick(self, sender, evant_args):
        self.flag = False
        self.Close()

    def checkBox_Unchecked(self, sender, evant_args):
        self.flagCheckBox1 = False

    def checkBox_Checked(self, sender, evant_args):
        self.flagCheckBox1 = True

    def _onChangeClick(self, sender, evant_args):
        self.flagLinePicked = True
        components = [Label('Выбери тип линий:'),
                            ComboBox('combobox1', lines),
                            Separator(),
                            Button('Выбрать')]
        form = FlexForm('Тип линий:', components)
        form.show()
        self.Style = form.values.get('combobox1')
    def _onChangeSysClick(self, sender, evant_args):
        components = [Label('Выбери тип системы:'),
                            ComboBox('combobox2', sysTypes),
                            Separator(),
                            Button('Выбрать')]
        form = FlexForm('Тип системы:', components)
        form.show()
        self.SysType = form.values.get('combobox2')
    def _onConsClick(self, sender, evant_args):
        self.flag = True
        self.flagParal = False
        self.flagcons = True
        self.Mood = 'cons'
        self.Close()
    def _onParalClick(self, sender, evant_args):
        self.flag = True
        self.flagParal = True
        self.flagcons = False
        self.Mood = 'paral'
        # ui.forms.Alert('Данный тип построения пока находится в разработке!', title="приносим свои извинения!")
        # self.flag = False
        self.Close()
    def _onReculcClick(self, sender, evant_args):
        self.readressHelpFlag = True
        self.flag = False 
        self.Close()
    def _onAddClick(self, sender, evant_args):
        self.addFlag = True
        self.flag = False 
        self.Close()
    def _onDellClick(self, sender, evant_args):
        self.dellFlag = True
        self.flag = False 
        self.Close()
    def _onReadressClick(self, sender, evant_args):
        self.readressFlag = True
        self.flag = False 
        self.Close()
    def _onHelpClick(self, sender, evant_args):
        webbrowser.open('https://articles.it-solution.ru/article/111966/')

MarkCount = 0
window = ParametersForm()
window.ShowDialog()

ActivStyle = elemLines[0].LineStyle
category_set = app.Create.NewCategorySet()
insert_cat = doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_ElectricalCircuit)
category_set.Insert(insert_cat)
ParamsToAdd = ['КП_И_Адрес текущий','КП_И_Адрес предыдущий','КП_И_Префикс системы','КП_И_Адрес устройства','КП_И_Порядковый номер']
shared_file = app.OpenSharedParameterFile()	
shared_params_groups = []
shared_params_definitions = []
for i in shared_file.Groups:
    shared_params_groups.append(i) 
    for j in i.Definitions:
        for searchParam in ParamsToAdd:
            if j.Name == searchParam:
                shared_params_definitions.append(j)
map = doc.ParameterBindings
it = map.ForwardIterator()
it.Reset()
if len(shared_params_definitions) != len(ParamsToAdd):
    ui.forms.Alert("Некорректный ФОП!", title="Внимание")
    script.exit()
with db.Transaction("Добавление параметров проекта"):
    for curDef in shared_params_definitions:
        newIB = app.Create.NewInstanceBinding(category_set)
        doc.ParameterBindings.Insert(curDef, newIB, DB.BuiltInParameterGroup.PG_DATA)



if window.flagLinePicked:
    ActivStyle = window.Style
else:
    if window.flag and window.flagCheckBox1:
        ui.forms.Alert('Тип лийний не выбран!', title="Тип линий:")
        form.show()
        window.Style = form.values.get('combobox1')
        ActivStyle = window.Style
maxAdressFlag = False
MoodFlag = False
try:
    count = int(window.txtBox1)
except:
    ui.forms.Alert("Введено некорректное стартовое значение!", title="Внимание")
    window.ShowDialog()
numbering = 0
for iterations in range(1000):
    with db.Transaction("СИТИ_Система"):
        if window.flagLinePicked:
            ActivStyle = window.Style
        if window.flag:
            try:
                if window.flagcons:
                    selectedReference = ui.Pick.pick_element(msg='Выбери следующий элемент. По окончанию - нажми "Esc"!')
                    try:
                        FirstElement = doc.GetElement(selectedReference.id)
                        if 'Марки'.lower() in FirstElement.Category.Name.lower() or str(FirstElement.Category.Name.lower())=='оси' or str(FirstElement.Category.Name.lower())=='связанные файлы'  or str(FirstElement.Category.Name.lower())=='линии':
                            MarkCount += 1
                            continue

                        if FirstElement.MEPModel.ElectricalSystems and iterations != MarkCount:
                            commands= [CommandLink('Да', return_value='true'),
                                    CommandLink('Нет', return_value='false')]
                            dialog = TaskDialog('Элемент уже имеет систему, добавить его в текущую, удалив старую?',
                                        commands=commands,
                                            show_close=True)
                            dialog_out = dialog.show()
                            if dialog_out == 'true':
                                for sysToDEl in FirstElement.MEPModel.ElectricalSystems:
                                    doc.Delete(sysToDEl.Id)
                            else:
                                MarkCount += 1
                                continue
                        numbering += 1
                        if iterations == MarkCount:
                            lnEnd = XYZ(FirstElement.Location.Point.X ,FirstElement.Location.Point.Y,0)
                            lnEndSys = FirstElement
                            try:
                                prefOfMark = doc.GetElement(FirstElement.GetTypeId()).get_Parameter(prefParam).AsString()
                            except:
                                prefOfMark = ""
                            try:
                                Adress = int(doc.GetElement(FirstElement.GetTypeId()).get_Parameter(AdressParam).AsDouble())
                            except:
                                Adress = 0
                            if Adress > 10:
                                try:
                                    Mark = str(prefOfMark) + (window.txtBox2).split('_')[-1] 
                                except:
                                    ui.forms.Alert("Нет символа '_' !", title="Внимание")
                                    script.exit()
                                MaxAdress = Adress
                                maxAdressFlag = True
                                Adress = 0
                                MarkInd = 0
                            else:
                                try:
                                    Mark = str(prefOfMark) + (window.txtBox2).split('_')[-1] + '.' +str(count)
                                except:
                                    ui.forms.Alert("Нет символа '_' !", title="Внимание")
                                    script.exit()
                                MarkInd = int(Mark.split('.')[-1])
                            FirstElement.get_Parameter(sysParam).Set(Mark)
                            FirstElement.get_Parameter(ElAdressParam).Set(numbering)
                            count += Adress
                            continue
                        if MoodFlag:
                            window.Mood = 'cons'
                            Mark = FirstMark
                            MoodFlag = False
                        exMarkToSys = Mark
                        try:
                            prefOfMark = doc.GetElement(FirstElement.GetTypeId()).get_Parameter(prefParam).AsString()
                        except:
                            prefOfMark = ""
                        try:
                            Mark = str(prefOfMark) + (window.txtBox2).split('_')[-1] + '.' +str(count)
                        except:
                            ui.forms.Alert("Нет символа '_' !", title="Внимание")
                            script.exit()
                        FirstElement.get_Parameter(sysParam).Set(Mark)
                        FirstElement.get_Parameter(ElAdressParam).Set(numbering)
                        try:
                            Adress = int(doc.GetElement(FirstElement.GetTypeId()).get_Parameter(AdressParam).AsDouble())
                        except:
                            Adress = 0
                        count += Adress
                        lnStart = lnEnd
                        lnStartSys = lnEndSys
                        lnEndSys = FirstElement
# __________________________________________________________________
                        try:
                            sys =  ElectricalSystem.Create(doc, [lnEndSys.Id],window.SysType)
                        except:
                            ui.forms.Alert('Некорректный тип системы, завершите работу!', title="Внимение!")
                            break
                        try:
                            sys.SelectPanel(lnStartSys)
                        except:
                            ui.forms.Alert('Некорректный тип панели, завершите работу!', title="Внимение!")
                            break
                        sys.get_Parameter(sysParam).Set(Mark)
                        sys.get_Parameter(ElAdressParam).Set(int(Mark.split('.')[-1]))
                        sys.get_Parameter(systemParam).Set(window.txtBox2)
                        sys.get_Parameter(ExElParam).Set(exMarkToSys)
                        lnEndCheck = XYZ(FirstElement.Location.Point.X,FirstElement.Location.Point.Y,0)
                        if str(round(lnEndCheck.X,5)) < str(round(lnStart.X,5)):
                            addX = 0.5
                        else:
                            addX = -0.5
                        if str(round(lnEndCheck.Y,5)) < str(round(lnStart.Y,5)):
                            addY = +0.5
                        else:
                            addY = -0.5
                        if str(round(lnEndCheck.X,5)) == str(round(lnStart.X,5)):
                            lnStart = lnEnd
                            lnEnd = XYZ(FirstElement.Location.Point.X,FirstElement.Location.Point.Y+addY,0)
                            if window.flagCheckBox1:
                                line = Line.CreateBound(lnStart, lnEnd)
                                DetailCurve = doc.Create.  NewDetailCurve(view,line) 
                                DetailCurve.LineStyle = ActivStyle
                            lnEnd = XYZ(FirstElement.Location.Point.X + addX,FirstElement.Location.Point.Y,0)
                        elif str(round(lnEndCheck.Y,5)) == str(round(lnStart.Y,5)):
                            lnStart = lnEnd
                            lnEnd = XYZ(FirstElement.Location.Point.X - addX,FirstElement.Location.Point.Y,0)
                            if window.flagCheckBox1:
                                line = Line.CreateBound(lnStart, lnEnd)
                                DetailCurve = doc.Create.  NewDetailCurve(view,line) 
                                DetailCurve.LineStyle = ActivStyle
                            lnEnd = XYZ(FirstElement.Location.Point.X,FirstElement.Location.Point.Y,0)
                        else:  
                            lnEnd = XYZ(FirstElement.Location.Point.X,lnEnd[1],0)
                            if window.flagCheckBox1:
                                line = Line.CreateBound(lnStart, lnEnd)
                                DetailCurve = doc.Create.NewDetailCurve(view,line)
                                DetailCurve.LineStyle = ActivStyle
                            lnStart = lnEnd
                            lnEnd = XYZ(FirstElement.Location.Point.X,FirstElement.Location.Point.Y + addY,0)
                            if window.flagCheckBox1:
                                line = Line.CreateBound(lnStart, lnEnd)
                                DetailCurve = doc.Create.NewDetailCurve(view,line)
                                DetailCurve.LineStyle = ActivStyle
                            lnEnd = XYZ(FirstElement.Location.Point.X + addX,FirstElement.Location.Point.Y,0)

                    except:
                        script.exit()
                        print(element.Id)
                
                
                
                elif window.flagParal:
                    selectedReference = ui.Pick.pick_element(msg='Выбери следующий элемент. По окончанию - нажми "Esc"!')
                    try:
                        FirstElement = doc.GetElement(selectedReference.id)
                        if 'Марки'.lower() in FirstElement.Category.Name.lower() or str(FirstElement.Category.Name.lower())=='оси' or str(FirstElement.Category.Name.lower())=='связанные файлы' or str(FirstElement.Category.Name.lower())=='Линии':
                            MarkCount += 1
                            continue
                        if MoodFlag:
                            FirstlnEnd = lnEnd
                            FirstlnEndSys = lnEndSys
                            FirstMark = Mark
                            MoodFlag = False
                            window.Mood = 'paral'
                        elif iterations == MarkCount:
                            FirstlnEnd =  XYZ(FirstElement.Location.Point.X,FirstElement.Location.Point.Y,0)
                            FirstlnEndSys = FirstElement
                            lnEnd = XYZ(FirstElement.Location.Point.X,FirstElement.Location.Point.Y,0)
                            try:
                                prefOfMark = doc.GetElement(FirstElement.GetTypeId()).get_Parameter(prefParam).AsString()
                            except:
                                prefOfMark = ""
                            try:
                                Adress = int(doc.GetElement(FirstElement.GetTypeId()).get_Parameter(AdressParam).AsDouble())
                            except:
                                Adress = 0
                            if Adress > 10:
                                try:
                                    Mark = str(prefOfMark) + (window.txtBox2).split('_')[-1] 
                                except:
                                    ui.forms.Alert("Нет символа '_' !", title="Внимание")
                                    script.exit()
                                MaxAdress = Adress
                                Adress = 0
                                MarkInd = 0
                            else:
                                try:
                                    Mark = str(prefOfMark) + (window.txtBox2).split('_')[-1] + '.' +str(count)
                                except:
                                    ui.forms.Alert("Нет символа '_' !", title="Внимание")
                                    script.exit()
                                MarkInd = int(Mark.split('.')[-1])
                            FirstElement.get_Parameter(sysParam).Set(Mark)
                            FirstElement.get_Parameter(ElAdressParam).Set(numbering)
                            markToSys = Mark
                            FirstMark = Mark
                            count += Adress
                            continue
                        sys =  ElectricalSystem.Create(doc, [FirstElement.Id], ElectricalSystemType.FireAlarm)
                        sys.SelectPanel(FirstlnEndSys)
                        # _______
                        try:
                            prefOfMark = doc.GetElement(FirstElement.GetTypeId()).get_Parameter(prefParam).AsString()
                        except:
                            prefOfMark = ""
                        try:
                            Mark = str(prefOfMark) + (window.txtBox2).split('_')[-1] + '.' +str(count)
                        except:
                            ui.forms.Alert("Нет символа '_' !", title="Внимание")
                            script.exit()
                        FirstElement.get_Parameter(sysParam).Set(Mark)
                        sys.get_Parameter(sysParam).Set(Mark)
                        sys.get_Parameter(systemParam).Set(window.txtBox2)
                        sys.get_Parameter(ElAdressParam).Set((int(Mark.split('.')[-1]) + (float(FirstMark.split('.')[-1])/1000)))
                        try:
                            Adress = int(doc.GetElement(FirstElement.GetTypeId()).get_Parameter(AdressParam).AsDouble())
                        except:
                            Adress = 0
                        count += Adress
                        lnStart = lnEnd
                        lnEnd = XYZ(FirstElement.Location.Point.X,FirstElement.Location.Point.Y,0)
                        if window.flagCheckBox1:
                            line = Line.CreateBound(lnStart, lnEnd)
                            DetailCurve = doc.Create.NewDetailCurve(view,line) 
                            DetailCurve.LineStyle = ActivStyle
                        lnEnd = FirstlnEnd
                    except:
                        script.exit()
                        print(element.Id)
                else:
                    ui.forms.Alert("Не выбран вариант построения!", title="Внимание")
                    window.ShowDialog()

            except:
                if window.flag:
                    if maxAdressFlag:
                        adressRemain = str(MaxAdress - count)
                    mood = window.Mood
                    window.flagLinePicked = False
                    txtBox1 = str(count)
                    txtBox2 = str(window.txtBox2)
                    window = ParametersForm()
                    window.txtBox1 = txtBox1 
                    window.txtBox2 = txtBox2 
                    window.textBox1.Text = txtBox1 
                    window.textBox2.Text = txtBox2 
                    if maxAdressFlag: 
                        window.TextAdress = adressRemain
                        window.textBox_adress.Text = adressRemain
                    window.ShowDialog()
                    count = int(window.txtBox1)
                    if mood != window.Mood:
                        MoodFlag = True
                        mood = window.Mood
                    try:
                        txtBox2ToInt = window.txtBox2.split('.')
                        newTxtBox2 = txtBox2ToInt[0] + '.' +str(int(txtBox2ToInt[-1])+1)
                    except:
                        try:
                            newTxtBox2 = str(int(window.txtBox2)+1)
                        except:
                            newTxtBox2 = window.txtBox2
                    if window.flagNew:
                        maxAdressFlag = False
                        window.flagLinePicked = False
                        window = ParametersForm()
                        window.txtBox2 = newTxtBox2 
                        window.textBox2.Text = newTxtBox2 
                        MarkCount = iterations+1
                        window.ShowDialog()
                        count = int(window.txtBox1)
                        window.flagNew = False

                    
                else:
                    break
        elif window.addFlag:
            AddEl.Add()
            window.addFlag = False
            window = ParametersForm()
            window.ShowDialog()
        elif window.dellFlag:
            Dell.Dell()
            window.dellFlag = False
            window = ParametersForm()
            window.ShowDialog()
        elif window.readressFlag:
            Readress.Readress()
            window.readressFlag = False
            window = ParametersForm()
            window.ShowDialog()
        elif window.readressHelpFlag:
            ReadressHelp.ReadressHepl()
            window.readressHelpFlag = False
            commands= [CommandLink('Да', return_value='true'),
                    CommandLink('Нет', return_value='false')]
            dialog = TaskDialog('Обновить адресацию?',
                        commands=commands,
                            show_close=True)
            dialog_out = dialog.show()
            if dialog_out == 'true':
                Readress.Readress()
            window = ParametersForm()
            window.ShowDialog()
        else:
            break
                    
