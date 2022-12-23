# -*- coding: utf-8 -*-
"""
KPLN:DIV:ROOM:COUNTER

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Номер квартиры"
__doc__ = 'Увиличивает или уменьшает номера всех квартир (кроме номера «0») на заданное значение. Необходимо для сквозной нумерации между отдельными проектами rvt' \

"""
Архитекурное бюро KPLN

"""


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,\
                              BuiltInParameter, StorageType
from pyrevit.framework import clr
import wpf
from System.Windows import Window
from rpw import doc, db
from pyrevit import script
from System.Collections.Generic import *
import clr
clr.AddReference('RevitAPI')


def reNumbering(currRoom, currInd, currNum):
    numParam = currRoom.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
    numParamList = numParam.split(".")
    newNum = ""
    strLen = len(numParamList)
    if currInd < strLen:
        for i in range(strLen):
            if i == currInd:
                newNum += str(currNum)
                if i == strLen-1:
                    continue
                else:
                    newNum += "."
            elif i < strLen-1:
                newNum += str(numParamList[i])
                newNum += "."
            else:
                newNum += str(numParamList[i])
        currRoom.get_Parameter(BuiltInParameter.ROOM_NUMBER).Set(newNum)
    else:
        raise Exception("Индекс выше положенного! Работа остановлена!")


class MyWindow(Window):
    def __init__(self):
        wpf.LoadComponent(self, 'X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Квартиры.panel\\Инструменты.stack\\Tools.pulldown\\RoomNumber.pushbutton\\Form.xaml')
        self.Parameters = []
        self.Value = 0
        room = FilteredElementCollector(doc).\
            OfCategory(BuiltInCategory.OST_Rooms).\
            WhereElementIsNotElementType().\
            FirstElement()
        for j in room.Parameters:
            if j.IsShared and j.StorageType == StorageType.String:
                self.Parameters.append(j)
        self.Parameters.sort(key=lambda x: x.Definition.Name)
        self.cbxPameter.ItemsSource = self.Parameters

    def Update(self):
        if self.cbxPameter.SelectedIndex == -1:
            self.OnOk.IsEnabled = False
            return
        try:
            self.Value = int(self.tbStartNumber.Text)
            self.OnOk.IsEnabled = True
        except Exception:
            self.OnOk.IsEnabled = False

    def OnSelectionChanged(self, sender, e):
        self.Update()

    def OnTextChanged(self, sender, e):
        self.Update()

    def numOnTextChanged(self, sender, e):
        self.Update()

    def OnButtonApply(self, sender, e):
        with db.Transaction(name="KPLN: Нумерация квартир"):
            try:
                for room in FilteredElementCollector(doc).\
                        OfCategory(BuiltInCategory.OST_Rooms).\
                        WhereElementIsNotElementType().\
                        ToElements():
                    value = ""
                    param = None
                    for j in room.Parameters:
                        SPName = self.cbxPameter.SelectedItem.Definition.Name
                        if j.IsShared and j.Definition.Name == SPName\
                                and j.StorageType == StorageType.String:
                            value = j.AsString()
                            param = j
                            break
                    if value != "" and value != "0" and value is not None\
                            and param is not None:
                        try:
                            number = int(value)
                            trueNum = number + self.Value
                            if self.isRename.IsChecked:
                                trueInd = self.numIndex.Text
                                try:
                                    reNumbering(room, int(trueInd)-1, trueNum)
                                except Exception as exc:
                                    print(exc)
                                    script.exit()
                            param.Set(str(trueNum))
                        except Exception as exc:
                            print(str(exc))
                            pass
            except Exception as exc:
                print(str(exc))
        self.Close()


MyWindow().ShowDialog()