# -*- coding: utf-8 -*-
"""
RenameViews

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Переименовать выбранные (Имя уровня)"
__doc__ = 'Переименовать виды с указанием префикса и постфикса: [префикс] <имя уровня целиком> [постфикс]' \

"""
Архитекурное бюро KPLN

"""
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms

def rename_plan_view(view, prefix, suffix):
    try: 
        try:
            number = str(view.get_Parameter(DB.BuiltInParameter.VIEWER_DETAIL_NUMBER).AsString())
            sheet_number = str(view.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER).AsString())
            if sheet_number != "None":
                try:
                    view.Name = '{}{}{}_{}/{}'.format(prefix, view.get_Parameter(DB.BuiltInParameter.PLAN_VIEW_LEVEL).AsString(), suffix, sheet_number, number)
                except:
                    view.Name = '{}{}{}_{}/{}_{}'.format(prefix, view.get_Parameter(DB.BuiltInParameter.PLAN_VIEW_LEVEL).AsString(), suffix, sheet_number, number, view.Id.ToString())
            else:
                try:
                    view.Name = '{}{}{}'.format(prefix, view.get_Parameter(DB.BuiltInParameter.PLAN_VIEW_LEVEL).AsString(), suffix)
                except:
                    view.Name = '{}{}{}_{}'.format(prefix, view.get_Parameter(DB.BuiltInParameter.PLAN_VIEW_LEVEL).AsString(), suffix, view.Id.ToString())
        except:
            try:
                template = doc.GetElement(view.ViewTemplateId)
                if template != None:
                    name = '"{}"'.format(template.Name)
                else:
                    name = doc.GetElement(view.GetTypeId()).FamilyName
                number = str(view.get_Parameter(DB.BuiltInParameter.VIEWER_DETAIL_NUMBER).AsString())
                sheet_number = str(view.get_Parameter(DB.BuiltInParameter.VIEWPORT_SHEET_NUMBER).AsString())
                if sheet_number == "None":
                    view.Name = '{}{}{}_{}'.format(prefix, name, suffix, view.Id.ToString())
                else:
                    try:
                        view.Name = '{}{}{}_{}/{}'.format(prefix, name, suffix, sheet_number, number)
                    except:
                        view.Name = '{}{}{}_{}/{}_{}'.format(prefix, name, suffix, sheet_number, number, view.Id.ToString())
            except:
                template = doc.GetElement(view.ViewTemplateId)
                if template != None:
                    name = '"{}"'.format(template.Name)
                else:
                    name = doc.GetElement(view.GetTypeId()).FamilyName
                view.Name = '{}{}{}_{}'.format(prefix, name, suffix, view.Id.ToString())
    except Exception as e:
        print("Ошибка: {}".format(str(e)))

views = []
others = []
selection = ui.Selection()
for element in selection:
    try:
        if element.unwrap().Category.Id.IntegerValue == -2000279:
          views.append(element.unwrap())
    except: 
        pass
try:
    if len(views) == 0:
        forms.alert("Не выбрано ни одного подходящего вида!")
    else:
        components = [ui.forms.flexform.Label("[Префикс] [Имя уровня] [Постфикс]\n"), ui.forms.flexform.Label("Префикс:"),ui.forms.flexform.TextBox('preffix', Text="АР.#_"), ui.forms.flexform.Label("Постфикс:"), ui.forms.flexform.TextBox('suffix', Text=""), ui.forms.flexform.Button("Запуск")]
    if len(views) != 0:
        form = ui.forms.FlexForm("Переименовать выбранный план", components)
        form.ShowDialog()
        prefix = form.values['preffix']
        postfix = form.values['suffix'] 
        with db.Transaction('Переименовать вид'):
            for view in views:
                rename_plan_view(view, prefix, postfix)
except: pass
