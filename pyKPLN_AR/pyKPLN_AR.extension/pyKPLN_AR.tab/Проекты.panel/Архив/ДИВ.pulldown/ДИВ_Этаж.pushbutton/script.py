# -*- coding: utf-8 -*-
"""
DIV_Level

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "ДИВ_Этаж"
__doc__ = 'Заполнение параметра ДИВ_Этаж и ДИВ_Этаж_Текст для всех элементов НА АКТИВНОМ ВИДЕ' \
"""
"""
from rpw import doc, uidoc, DB, UI, db, ui, revit as Revit
from pyrevit import revit, DB, UI
from System.Drawing import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert


def GetLevelParameter(el, par):
    bp = None
    try: bp = doc.GetElement(el.LevelId).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(el.Host.LevelId).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(el.SuperComponent.LevelId).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(el.SuperComponent.Host.LevelId).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(el.SuperComponent).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(el.SuperComponent.Host).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(doc.GetElement(el.GroupId).LevelId).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(doc.GetElement(el.Host.GroupId).LevelId).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(doc.GetElement(el.SuperComponent.GroupId).LevelId).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(doc.GetElement(el.SuperComponent.Host.GroupId).LevelId).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(el.get_Parameter(DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId()).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(doc.GetElement(el.HostId).get_Parameter(DB.BuiltInParameter.STAIRS_BASE_LEVEL_PARAM).AsElementId()).LookupParameter(par)
    except: pass
    if bp != None: return bp
    #
    try: bp = doc.GetElement(el.Host).LookupParameter(par)
    except: pass
    if bp != None: return bp
    return bp


with db.Transaction(name="DIV"):
    for element in DB.FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType().ToElements():
        if element != None:
            try:
                for p in ["ДИВ_Этаж", "ДИВ_Этаж_Текст"]:
                    bp = GetLevelParameter(element, p)
                    if bp != None:
                        par = element.LookupParameter(p)
                        if par.CanBeAssociatedWithGlobalParameters():
                            par.Set(bp.AsString())
            except: pass
