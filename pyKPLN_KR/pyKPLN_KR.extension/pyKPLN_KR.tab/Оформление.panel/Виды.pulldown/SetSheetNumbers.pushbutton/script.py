# -*- coding: utf-8 -*-
"""
NumberSheets

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Пронумеровать листы"
__doc__ = 'Номерация листов с указанием префикса и начала номерации.' \

"""
Архитекурное бюро KPLN

"""
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms

def set_name_number(sheet, prefix, number, general):
    SheetNumber = ""
    try: 
        if not general:
            sheet_number = str(number)
        else:
            if number < 10:
                sheet_number = "00{}".format(str(number))
            elif number <100:
                sheet_number = "0{}".format(str(number))
            else:
                sheet_number = "{}".format(str(number))
        sheet.SheetNumber = '{} - {}'.format(prefix, sheet_number)
        try:
            sheet.LookupParameter("КП_Номер листа").Set('{}'.format(sheet_number))
        except: pass
        return sheet
    except Exception as e:
        print("Ошибка: {}\n\tПопытка записать значение «{}»".format(str(e), SheetNumber))
        return sheet

sheets = []
sheets_numbers = []
sheets_sorted = []
others = []
selection = ui.Selection()
for element in selection:
    try:
        if element.unwrap().Category.Id.IntegerValue == -2003100:
          sheets.append(element.unwrap())
          sheets_numbers.append(element.unwrap().SheetNumber)
    except: 
        pass
sheets_numbers.sort()
for num in sheets_numbers:
    for sheet in sheets:
        if sheet.SheetNumber == num:
            sheets_sorted.append(sheet)

try:
    if len(sheets_sorted) == 0:
        forms.alert("Не выбрано ни одного подходящего вида!")
    else:
        components = [ui.forms.flexform.Label("[Том] - [Номер]\n"), ui.forms.flexform.Label("Том:"),ui.forms.flexform.TextBox('preffix', Text="КР1"), ui.forms.flexform.Label("Начать номерацию с:"), ui.forms.flexform.TextBox('num', Text="1"), ui.forms.flexform.Label("Нумерация:"), ui.forms.flexform.CheckBox('general', '00# вместо #', default=False), ui.forms.flexform.Button("Запуск")]
    if len(sheets_sorted) != 0:
        form = ui.forms.FlexForm("Пронумеровать выбранный вид", components)
        form.ShowDialog()
        prefix = form.values['preffix']
        suffix = form.values['num'] 
        general = form.values['general'] 
        number = int(suffix)
        newsheets = []
        with db.Transaction('Переименовать вид TEMP'):
            for sheet in sheets_sorted:
                newsheets.append(set_name_number(sheet, "TEMP{}".format(str(sheet.Id.IntegerValue)), number, False))
                number += 1
        number = int(suffix)
        with db.Transaction('Переименовать вид'):
            for sheet in newsheets:
                set_name_number(sheet, prefix, number, general)
                number += 1
        doc.Regenerate()
except: pass
