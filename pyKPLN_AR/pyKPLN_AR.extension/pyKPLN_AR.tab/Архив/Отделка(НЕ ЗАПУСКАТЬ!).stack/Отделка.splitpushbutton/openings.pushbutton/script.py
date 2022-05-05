# -*- coding: utf-8 -*-
"""
Test

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "DO NOT PUSH!"
__doc__ = 'Не нажимать!' \

"""
Архитекурное бюро KPLN

"""
import math
from pyrevit.framework import clr

# clr.AddReference("RevitNodes")

import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
from rpw.ui.forms import CommandLink, TaskDialog, Alert
import datetime

for pt in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory. OST_SharedBasePoint).WhereElementIsNotElementType().ToElements():
    shared = pt.Position
    q = 3.28084
    point = DB.XYZ(-8.847 * q + shared.X, 78.373 * q + shared.Y, 157.076 * q + shared.Z)
    print(point)