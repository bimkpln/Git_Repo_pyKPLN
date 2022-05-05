# coding: utf-8

__title__ = "Сгруппировать системы"
__author__ = 'Tima Kutsko'
__doc__ = '''Группирует системы по параметрам КР_О_Имя Системы.\n
             Нужно для корректного вывода спецификации и генерации схем'''

from rpw import revit, db, ui
from System import Guid
from libKPLN.MEP_Elements import allMEP_Elements
from System.Windows.Forms import Form, Button, ComboBox, Label
from System.Drawing import Point, Size
import webbrowser
import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')


class Window(Form):
    def __init__(self):
        self.__x1 = 30
        self.__y1 = 60
        self.__x2 = 300
        self.__y2 = 60
        self.Size = Size(810, 200)
        self.CenterToScreen()
        self.AutoScroll = True
        self._inintAdd()
        self._inintStart()
        self._inintHelp()
        self.dataList = list()

    def _inintAdd(self):
        self._btnAdd = Button()
        self._btnAdd.Text = 'Добавить группу'
        self._btnAdd.Size = Size(200, 40)
        self._btnAdd.Location = Point(30, 0)
        self.Controls.Add(self._btnAdd)
        self._btnAdd.Click += self._addGroup1
        self._btnAdd.Click += self._addGroup2

    def _inintStart(self):
        self._btnStr = Button()
        self._btnStr.Text = 'Выполнить'
        self._btnStr.Size = Size(200, 40)
        self._btnStr.Location = Point(300, 0)
        self.Controls.Add(self._btnStr)
        self._btnStr.Click += self._result

    def _inintHelp(self):
        self._btnHlp = Button()
        self._btnHlp.Text = 'Помощь'
        self._btnHlp.Size = Size(200, 40)
        self._btnHlp.Location = Point(570, 0)
        self.Controls.Add(self._btnHlp)
        self._btnHlp.Click += self._onHelpClick

    def _addGroup1(self, sender, evant_args):
        self._cmb1 = ComboBox()
        self._cmb1.Size = Size(200, 40)
        self._cmb1.Location = Point(self.__x1, self.__y1)
        for sys in sorted(sysDataSet):
            self._cmb1.Items.Add(sys)
        self._mark = Label()
        self._mark.Size = Size(10, 10)
        self._mark.Text = "+"
        self._mark.Location = Point(260, self.__y1 + 5)
        self.Controls.Add(self._mark)
        self.Controls.Add(self._cmb1)
        self.__y1 += 25

    def _addGroup2(self, sender, evant_args):
        self._cmb2 = ComboBox()
        self._cmb2.Size = Size(200, 40)
        self._cmb2.Location = Point(self.__x2, self.__y2)
        for sys in sorted(sysDataSet):
            self._cmb2.Items.Add(sys)
        self.__y2 += 25
        self.Controls.Add(self._cmb2)

    def _onHelpClick(self, sender, evant_args):
        webbrowser.open('https://kpln.bitrix24.ru/marketplace/app/16/')

    def _result(self, sender, evant_args):
        for ctrl in self.Controls:
            if isinstance(ctrl, ComboBox):
                self.dataList.append(ctrl.Text.ToString())
        self.Close()


# Get systems collection in project
doc = revit.doc
sysParam = Guid("21213449-727b-4c1f-8f34-de7ba570973a")  # КП_О_Имя Системы
elemGetter = allMEP_Elements(doc).getOVVK()[0]
sysDataSet = set()
for curElem in elemGetter:
    param = curElem.get_Parameter(sysParam)
    paramData = param.AsString()
    if paramData:
        sysDataSet.add(str(paramData))

# Create example of input form
win = Window()
win.ShowDialog()

# Analizing data
with db.Transaction('pyKPLN_MEP: Объединение данных в КП_О_Имя Системы'):
    itemsList = win.dataList
    cnt = 0
    while cnt < len(itemsList)-1:
        filtColl_1 = list(filter(lambda x:
                                 itemsList[cnt] == x.get_Parameter(sysParam).
                                 AsString(), elemGetter))
        filtColl_2 = list(filter(lambda x:
                                 itemsList[cnt+1] == x.get_Parameter(sysParam).
                                 AsString(), elemGetter))
        filtColl = filtColl_1 + filtColl_2
        for item in filtColl:
            item.\
                get_Parameter(sysParam).\
                Set(itemsList[cnt] + '/' + itemsList[cnt+1])
        cnt += 1

ui.forms.Alert("Завершено!", title="Завершено")