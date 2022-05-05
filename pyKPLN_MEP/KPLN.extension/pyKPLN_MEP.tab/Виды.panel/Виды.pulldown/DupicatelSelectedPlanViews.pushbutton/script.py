# -*- coding: utf-8 -*-
__doc__ = 'Копирвание планов с детализацией'
__title__ = "Копирвание планов с детализацией"


import re
from collections import namedtuple, OrderedDict
from rpw import doc, ui, db, DB
from pyrevit import script
from pyrevit import revit
import viewedit as ve
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger


view_templates = db.Collector(of_category='OST_Views',
                                is_type=False,
                                where=lambda x: x.IsTemplate).get_elements(False)
templates_dict = {e.Name: e for e in view_templates}


selection = ui.Selection()
if len(selection) == 0:
    ui.forms.Alert('Ничего не выбранно', exit=True)

# form
ComboBox = ui.forms.flexform.ComboBox
Label = ui.forms.flexform.Label
Button = ui.forms.flexform.Button
CheckBox = ui.forms.flexform.CheckBox
TextBox = ui.forms.flexform.TextBox
components = [Label("Выбери шаблон вида:"),
              ComboBox("template", templates_dict),
              Label("Приставка:"),
              TextBox('preffix', Text="ОВ.О_"),
              Label("Суффикс:"),
              TextBox('suffix', Text="_План"),
              CheckBox('general', 'Общий план'),
              Button("Выбрать")]

form = ui.forms.FlexForm("Копирвание планов с детализацией", components)
form.ShowDialog()

try:
    template = form.values['template']
    prefix = form.values['preffix']
    suffix = form.values['suffix']
    general = form.values['general']

    if len(selection) > 0:
        #getting info logger about user
        log_name = "Виды_Копирование планов с детализацией"
        InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
        #main part of code
        with db.Transaction(name='pyKPLN_Копирование выбранных видов'):
            for el in selection:
                if ve.is_plan_view(el):
                    view = el
                    new_view = ve.duplicate_plan_view_with_detailing(view, template.Id)
                    ve.rename_plan_view(new_view, prefix, suffix, general)
                    print('{} создан.'.format(new_view.Name))
except:
    print("Обратись к разработчику")