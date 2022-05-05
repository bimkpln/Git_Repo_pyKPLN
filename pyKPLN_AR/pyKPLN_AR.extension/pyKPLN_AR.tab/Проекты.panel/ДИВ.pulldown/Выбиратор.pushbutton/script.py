# -*- coding: utf-8 -*-
"""
KPLN:DIV:SELECTOR

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Выбрать + Заполнить"
__doc__ = 'Заполнить параметры выбранных элементов. Инструмент позволяет быстро выбирать элементы, для выбора которым необходимо использовать клавишу TAB' \

"""
Архитекурное бюро KPLN

"""


from pyrevit.framework import wpf
from System.Windows import Window
from rpw import doc, uidoc, DB, UI, db


class WPFCategory():
    def __init__(self, category):
        self.IsChecked = False
        self.Category = category


class MassSelectorFilter (UI.Selection.ISelectionFilter):
    def AllowElement(self, element=DB.Element):
        if element.Category.Id.IntegerValue != -2000095:
            return True
        else:
            return False
    def AllowReference(self, refer, point):
        return False


class ParametersForm(Window):
    def __init__(self, parameters):
        wpf.LoadComponent(self, 'Z:\\pyRevit\\pyKPLN_AR (alpha)\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\Проекты.panel\\ДИВ.pulldown\\Выбиратор (формы).pushbutton\\Parameters.xaml')
        self.cbxParameters.ItemsSource = parameters

    def OnBtnApply(self, sender, e):
        global pickedParameter
        global pickedValue
        pickedParameter = self.cbxParameters.SelectedItem
        pickedValue = self.tbxValue.Text
        self.Close()


def PickElement():
    elements = []
    try:
        pick_filter = MassSelectorFilter()
        ref_elements = uidoc.Selection.PickObjects(UI.Selection.ObjectType.Element, pick_filter, "Выберите элементы кроме : <Группы модели>")
        for i in ref_elements:
            try:
                elements.append(doc.GetElement(i))
            except: pass
    except:
        pass
    return elements


def GetGeometry(element):
    solids = []
    if element.Category.Id.IntegerValue == -2000160:
        SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Finish
        calculator = DB.SpatialElementGeometryCalculator(doc, DB.SpatialElementBoundaryOptions())
        results = calculator.CalculateSpatialElementGeometry(element)
        room_solid = results.GetGeometry()
        if room_solid != None:
            solids.append(room_solid)
    else:
        options = DB.Options()
        options.DetailLevel = DB.ViewDetailLevel.Fine
        options.IncludeNonVisibleObjects = True
        for i in element.get_Geometry(options):
            try:
                if type(i) == DB.Solid:
                    if i.SurfaceArea != 0:
                        solids.append(i)
            except: pass
            try:
                for g in i.GetInstanceGeometry():
                    if type(g) == DB.Solid:
                        if g.SurfaceArea != 0:
                            solids.append(g)
            except: pass
    return solids


def writeParam(param, value):
    if param.IsShared == pickedParameter.IsShared\
            and param.Definition.Name == pickedParameter.Definition.Name\
            and param.StorageType == DB.StorageType.String:
        param.Set(value)


cats = []
pickedElements = PickElement()
if len(pickedElements) > 0:
    params = []
    Hash = []
    for i in pickedElements:
        for j in i.Parameters:
            if j.StorageType == DB.StorageType.String\
                    and j.Definition.Name not in Hash:
                params.append(j)
                Hash.append(j.Definition.Name)
    params.sort(key=lambda x: x.Definition.Name, reverse=False)
    pickedParameter = None
    pickedValue = None
    ParametersForm(params).ShowDialog()
    if pickedParameter is not None and pickedValue is not None:
        with db.Transaction(name="КПЛН_Запись в параметр"):
            for i in pickedElements:
                try:
                    for j in i.Parameters:
                        writeParam(j, pickedValue)
                    if i.GetSubComponentIds():
                        for ID in i.GetSubComponentIds():
                            sub_element = doc.GetElement(ID)
                            for j in sub_element.Parameters:
                                writeParam(j, pickedValue)
                except Exception as exc:
                    if "no attribute 'GetSubComponentIds'" not in str(exc)\
                            and "The parameter is read-only." not in str(exc):
                        print(exc)
