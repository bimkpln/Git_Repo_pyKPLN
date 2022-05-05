# coding: utf8
from datetime import datetime
import re
from decimal import Decimal
from System.Collections.Generic import List

from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script, output
logger = script.get_logger()
output = output.get_output()


def create_level_name(level):
    if isinstance(level, str):
        level_name = level
        level = db.Collector(of_class='Level', is_type=False,
                             where=lambda x: x.Name == level_name).get_first()
    else:
        level_name = level.Name

    # extract level index: -01, 01, 02
    lvl_pat = re.compile("(?:(\-)?(\d\d))\_{1}")
    lvl_index = re.findall(lvl_pat, str(level_name))
    index = "".join(lvl_index[0])

    elevation = level.Elevation * 304.8
    f_elevation = format_elevation(elevation)

    new_name = "{}_({})".format(index, f_elevation)

    return new_name


def format_elevation(num):
    """ Convert elevation from mm to meters """
    num_m = Decimal(num / 1000)
    num_str = format(num_m, '.3f')
    if num > 0:
        num_str = '+' + num_str
    return num_str

def rename_plan_view(view, prefix, suffix, full_lvl_name=True, tag=False):

    level_name = view.get_Parameter(DB.BuiltInParameter.PLAN_VIEW_LEVEL).AsString()
    if full_lvl_name:
        level_name = create_level_name(level_name)

    new_name = '{}{}{}'.format(prefix, level_name, suffix)
    try:   
        view.Name = new_name
        if tag:
            output.print_md('План **{}** переименован на **{}**'.format(view.Name, new_name))
    except:
        output.print_md("План с видом **{}** уже есть в проекте. Переименуй вручную!".format(new_name))
    


def duplicate_plan_view(view, template_id):
    new_view_id = view.Duplicate(DB.ViewDuplicateOption.Duplicate)
    new_view = doc.GetElement(new_view_id)
    new_view.ViewTemplateId = template_id
    return new_view

def duplicate_plan_view_with_detailing(view, template_id):
    new_view_id = view.Duplicate(DB.ViewDuplicateOption.WithDetailing)
    new_view = doc.GetElement(new_view_id)
    new_view.ViewTemplateId = template_id
    return new_view    


def is_plan_view(view):
    return view.unwrap().GetType().Name == 'ViewPlan'


def set_view_range(view, top_off=None, cut_off=None,
                   bot_off=None, dep_off=None):

    range = view.GetViewRange()
    d = 304.8
    if top_off:
        range.SetOffset(DB.PlanViewPlane.TopClipPlane, top_off / d)
    if cut_off:
        range.SetOffset(DB.PlanViewPlane.CutPlane, cut_off / d)
    if bot_off:
        range.SetOffset(DB.PlanViewPlane.BottomClipPlane, bot_off / d)
    if dep_off:
        range.SetOffset(DB.PlanViewPlane.ViewDepthPlane, dep_off / d)

    if any([top_off, cut_off, bot_off, dep_off]):
        view.SetViewRange(range)
    return view