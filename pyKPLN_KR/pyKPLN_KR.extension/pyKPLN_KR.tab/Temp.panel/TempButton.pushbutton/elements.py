# module elements.py
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB

doc = __revit__.ActiveUIDocument.Document  # noqa

def get_type_name(element):
    if hasattr(element, 'GetTypeId'):
        element_type = doc.GetElement(element.GetTypeId())
        return element_type.Parameter[DB.BuiltInParameter.SYMBOL_NAME_PARAM].AsString()

def is_of_type(element, type_name):
    return type_name == get_type_name(element)
