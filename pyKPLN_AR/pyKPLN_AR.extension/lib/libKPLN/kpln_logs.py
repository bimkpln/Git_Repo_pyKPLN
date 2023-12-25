# -*- coding: utf-8 -*-
from Autodesk.Revit import DB
import System
import time


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


class KPLN_Logs:
    def __init__(self, guid, field_name, field_type, storage_name):

        self.guid = System.Guid(guid)
        self.field_name = field_name
        self.field_type = field_type
        self.storage_name = storage_name

        self.schema, self.field = self._create_schema()
        try:
            self.read()
        except:
            self.write("---Начало логов---")

    def _create_schema(self):
        schema_builder = DB.ExtensibleStorage.SchemaBuilder(self.guid)
        schema_builder.SetReadAccessLevel(DB.ExtensibleStorage.AccessLevel.Public)
        field = schema_builder.AddSimpleField(self.field_name, self.field_type)
        schema_builder.SetSchemaName(self.storage_name)
        schema = schema_builder.Finish()

        return schema, field
    
    def _get_element(self):
        element = doc.ProjectInformation
        return element
    
    def write(self, data):
        entity = DB.ExtensibleStorage.Entity(self.schema)
        entity.Set(self.field_name, data)
        self._get_element().SetEntity(entity)

    def read(self):
        entity = self._get_element().GetEntity(self.schema)
        field_name = self.schema.GetField(self.field_name)
        data = entity.Get[str](field_name)

        return data

def create_obj(guid, field_name, storage_name):
    obj = KPLN_Logs(guid, field_name, str, storage_name)
    return obj

def write_log(obj, data):
    string = obj.read()
    string += "\n" + str(time.ctime()) + " // " + str(__revit__.Application.Username) + " // " + data
    obj.write(string)

def read_log(obj):
    print(obj.read())

