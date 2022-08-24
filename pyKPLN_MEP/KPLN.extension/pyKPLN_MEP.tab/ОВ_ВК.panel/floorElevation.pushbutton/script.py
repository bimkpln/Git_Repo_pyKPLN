# -*- coding: utf-8 -*-

__title__ = "Отметка по этажу"
__author__ = 'Tima Kutsko'
__doc__ = "Ставит спец. элементы по стоякам,"\
          "для маркировки перекрытий на схемах"


from Autodesk.Revit.DB import FilteredElementCollector, Level,\
    FilterStringRule, ParameterValueProvider, FilterStringContains,\
    ElementId, ElementParameterFilter, BuiltInParameter,\
    XYZ, Outline, BoundingBoxIntersectsFilter, FamilySymbol
from Autodesk.Revit.DB.Mechanical import Duct
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.DB.Structure import StructuralType
from rpw import revit, db
from System import Guid
from rpw.ui.forms import FlexForm, Label, Button, TextBox, Separator
from pyrevit import script


# Form
try:
    components = [Label("Узел ввода данных", h_align="Center"),
                  Label("Часть имени уровня, для фильтрации"),
                  TextBox("filterPart", Text=""),
                  Label(
                        "Индекс/-ы символа/-ов этажа """
                        """(через запятую, начиная с 0!).""",
                        MinWidth=350
                        ),
                  TextBox("indexPart", Text=""),
                  Separator(),
                  Label(
                        "Пример: имя этажа К2-02-(+3.000). "
                        "'К'=0,'2'=1,'-'=2,'0'=3,'2'=4",
                        MinWidth=350
                        ),
                  Label(
                        """Индексы символов этажа (02) """
                        """3 и 4 (записать в формате '3,4').""",
                        MinWidth=350
                        ),
                  Separator(),
                  Button("Запуск")
                  ]
    form = FlexForm("Отметка по этажу",
                    components
                    )
    form.show()
    filterPart = form.values["filterPart"]
    indexPart = form.values["indexPart"]
except KeyError:
    script.exit()

# Main code
doc = revit.doc
viewId = doc.ActiveView.Id
# КП_И_Отметка уровня
elemLevElevParam = Guid("9bf9520b-59aa-4c4d-a52e-5d8209041792")

# Getting true levels by part of name
valueProvider = ParameterValueProvider(ElementId(-1008000))
evaluator = FilterStringContains()
ruleString = filterPart
string_filter = FilterStringRule(valueProvider, evaluator, ruleString, False)
el_filter = ElementParameterFilter(string_filter)
levels_collector = FilteredElementCollector(doc).OfClass(Level).\
                    WhereElementIsNotElementType()
true_levels = levels_collector.WherePasses(el_filter).ToElements()


# Getting true family symbol of tag
famSymb_collector = FilteredElementCollector(doc).\
            OfClass(FamilySymbol).\
            ToElements()
for famSyb in famSymb_collector:
    if famSyb.FamilyName == '503_Технический_ОтметкаУровня(Об)':
        true_famSyb = famSyb


# Creating dict with levels name and elevation
true_elements_list = list()
sotringKey = lambda lev: lev.\
    get_Parameter(BuiltInParameter.LEVEL_ELEV).\
    AsDouble()

with db.Transaction('pyKPLN_Отметка по этажу'):
    for level in sorted(true_levels, key=sotringKey):
        level_elevation = level.\
            get_Parameter(BuiltInParameter.LEVEL_ELEV).\
            AsDouble()
        outline = Outline(
                        XYZ(-1000, -1000, level_elevation),
                        XYZ(1000, 1000, level_elevation)
                        )
        level_elevation = level_elevation * 304.8 / 1000
        bounding_box_filter = BoundingBoxIntersectsFilter(outline)
        true_pipes = FilteredElementCollector(doc, viewId).\
            OfClass(Pipe).\
            WhereElementIsNotElementType().\
            WherePasses(bounding_box_filter).\
            ToElements()
        true_ducts = FilteredElementCollector(doc, viewId).\
            OfClass(Duct).\
            WhereElementIsNotElementType().\
            WherePasses(bounding_box_filter).\
            ToElements()
        for pipe in true_pipes:
            origin = pipe.Location.Curve.Origin
            true_point = XYZ(origin.X, origin.Y, 0)
            familySymbol = true_famSyb
            # Активация семейства, если не размещено ни одного экземпляра
            if not familySymbol.IsActive:
                familySymbol.Activate()
            structuralType = StructuralType.NonStructural
            new_element = doc.Create.NewFamilyInstance(
                true_point,
                familySymbol,
                level,
                structuralType
                )
            pipeOutDiam = pipe.\
                get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).\
                AsDouble()
            if len(indexPart):
                trueName = ''
                for n in indexPart.split(','):
                    nPart = level.Name[int(n)]
                    if nPart != "_" and nPart != " " and nPart != "+":
                        trueName += nPart
            else:
                trueName = level.Name
            new_element.\
                LookupParameter('КП_И_Имя уровня').\
                Set(trueName)
            pipe_system = pipe.\
                get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).\
                AsString()
            new_element.\
                get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).\
                Set(pipe_system)
            if level_elevation == 0 or level_elevation > 0:
                new_element.\
                    get_Parameter(elemLevElevParam).\
                    Set('+' + '{:.3f}'.format(level_elevation).ToString())
            else:
                new_element.\
                    get_Parameter(elemLevElevParam).\
                    Set('{:.3f}'.format(level_elevation).ToString())
