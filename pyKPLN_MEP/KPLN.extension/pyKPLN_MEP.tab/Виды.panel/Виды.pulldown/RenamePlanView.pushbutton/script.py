# coding: utf8

__title__ = "Замена имени вида (план этажа)"
__doc__ = "Переименовать выбранный план этажа"

from datetime import datetime
import re
from System.Collections.Generic import List
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from viewedit import rename_plan_view
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger

selection = ui.Selection()
is_plan_views = all(v.unwrap().GetType().Name == 'ViewPlan' for v in selection)

ComboBox = ui.forms.flexform.ComboBox
Label = ui.forms.flexform.Label
Button = ui.forms.flexform.Button
CheckBox = ui.forms.flexform.CheckBox
TextBox = ui.forms.flexform.TextBox
try:
    if len(selection) == 0:
            print("Выбери (выдели через shift/ctrl) в модели план(-ы) для замены имени!")
    else:
        components = [Label("Узел ввода данных"),
                    Label("Приставка имени вида:"),
                    TextBox('preffix', Text="ОВ.О_"),
                    Label("Суффикс имени вида:"),
                    TextBox('suffix', Text="_План"),
                    CheckBox('general', 'Общий план', default=False),              
                    Button("Запуск")]
    form = ui.forms.FlexForm("Переименовать выбранный план", components)
    form.ShowDialog()
    if is_plan_views:
        #getting info logger about user
        log_name = "Виды_Замена имени вида"
        InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
        #main part of code
        views = selection
        prefix = form.values['preffix']
        suffix = form.values['suffix'] 
        general = form.values['general']   
        with db.Transaction('pyKPLN_Перименование планов'):
            for view in views:
                rename_plan_view(view, prefix, suffix, full_lvl_name=general, tag=True) 
    else:
        print('Работает только с планами этажей!')
except:
    script.exit()
