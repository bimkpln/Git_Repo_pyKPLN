#coding: utf-8

__title__ = "Генерация видов"
__author__ = 'Tima Kutsko'
__doc__ = 'Автоматическая генерация видов (одна система - один вид)\n'\
          'для элементов на открытом виде и с применением настроек\n'\
          'данного вида'

from Autodesk.Revit.DB import BuiltInParameter, ParameterFilterRuleFactory,\
                              ParameterFilterElement, ViewDuplicateOption,\
                              ElementId, ElementParameterFilter, FilterRule
from rpw import revit, db
from pyrevit import script, HOST_APP
from pyrevit.framework import wpf
from System.Collections.Generic import List
from System import Guid
from libKPLN.MEP_Elements import allMEP_Elements
from System.Windows import Window
import webbrowser
import clr
clr.AddReference('RevitAPI')
clr.AddReference('System.Windows.Forms')


# first form
class ParametersForm(Window):
    def __init__(self):
        wpf.LoadComponent(self, 'Z:\\pyRevit\\pyKPLN_MEP\\KPLN.extension\\pyKPLN_MEP.tab\\ОВ_ВК.panel\\Системы и виды.pulldown\\3_viewsCr.pushbutton\\startForm.xaml')

    def _onBtnClick(self, sender, evant_args):
        self.Close()

    def _onExtClick(self, sender, evant_args):
        script.exit()
        self.Close()

    def _onHelpClick(self, sender, evant_args):
        webbrowser.open('https://kpln.bitrix24.ru/marketplace/app/16/')


# main code
ParametersForm().ShowDialog()
sysParam = Guid("21213449-727b-4c1f-8f34-de7ba570973a")  # КП_О_Имя Системы
output = script.get_output()
doc = revit.doc
actView = doc.ActiveView
isTemplView = actView.\
              get_Parameter(BuiltInParameter.VIEW_TEMPLATE_FOR_SCHEDULE).\
              AsElementId()
elemGetter = allMEP_Elements(doc, False, True, actView.Id)
IfilterCat = List[ElementId]()
sysNames = set()
isRev2020 = int(HOST_APP.version) >= 2020

with db.Transaction("pyKPLN_MEP: Генерация видов"):
    ovvkElems = elemGetter.getOVVK()[0]
    catOvvk = elemGetter.getOVVK()[1]
    if isTemplView != ElementId(-1):
        output.print_md("**Только для видов без шаблонов!**")
        script.exit()

    for cat in catOvvk:
        IfilterCat.Add(ElementId(cat))

    for curElem in ovvkElems:
        param = curElem.get_Parameter(sysParam)
        try:
            system_name = param.AsString()
            if system_name:
                sysNames.add(system_name)
        except Exception:
            param = None
            output.print_md("{} нет системы. Устрани ошибки и перезапусти".
                            format(output.linkify(curElem.Id)))
            script.exit()

    for sysName in sysNames:
        if isRev2020:
            fRule = ParameterFilterRuleFactory.CreateNotContainsRule(param.Id,
                                                                     sysName,
                                                                     False
                                                                     )
            filterRules = ElementParameterFilter(fRule)
        else:
            filterRules = List[FilterRule]()
            fRule = ParameterFilterRuleFactory.CreateNotContainsRule(param.Id,
                                                                     sysName,
                                                                     False
                                                                     )
            filterRules.Add(fRule)
        filterName = "prog_КП_О_Имя системы_!*" + sysName + "*!"
        try:
            newFilter = ParameterFilterElement.Create(doc,
                                                      filterName,
                                                      IfilterCat,
                                                      filterRules
                                                      )
            newViewId = doc.ActiveView.Duplicate(ViewDuplicateOption.Duplicate)
            newView = doc.GetElement(newViewId)
            newView.AddFilter(newFilter.Id)
            newView.SetFilterVisibility(newFilter.Id, False)
            if "/" in sysName:
                newView.Name = "Схема систем " + sysName
            else:
                newView.Name = "Схема системы " + sysName
        except Exception as exc:
            print(exc)
