# module parameters.py
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit import DB

def get_parameter_value_v1(parameter):
    if isinstance(parameter, DB.Parameter):
        storage_type = parameter.StorageType
        if storage_type == DB.StorageType.Integer:
            return parameter.AsInteger()
        elif storage_type == DB.StorageType.Double:
            return DB.UnitUtils.ConvertFromInternalUnits(parameter.AsDouble(), parameter.DisplayUnitType)
        elif storage_type == DB.StorageType.String:
            return parameter.AsString()
        elif storage_type == DB.StorageType.ElementId:
            return parameter.AsElementId()
