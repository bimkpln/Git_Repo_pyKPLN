# coding: utf-8
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory


def elGetter(document, catList, curView, onM=True, onV=False):
    elemList = list()
    trueCatList = list()
    for cat in catList:
        if onM:
            coll = FilteredElementCollector(document).\
                   OfCategory(cat).\
                   WhereElementIsNotElementType().\
                   ToElements()
            elemList.extend(coll)
            if coll:
                trueCatList.append(cat)
        elif onV:
            coll = FilteredElementCollector(document, curView).\
                   OfCategory(cat).\
                   WhereElementIsNotElementType().\
                   ToElements()
            elemList.extend(coll)
            if coll:
                trueCatList.append(cat)
        else:
            print("Error in code: on view OR on model. Nohing else!!!")
    return elemList, trueCatList


class allMEP_Elements:
    '''
    For initialization it needs document, bool parts\n
    (on the model OR on the view) and view Id
    '''

    def __init__(self, doc, onModel=True, onView=False, view=None):
        self.__doc = doc
        self.__onModel = onModel
        self.__onView = onView
        self._view = view
        self._builtInCat_OVVK = [BuiltInCategory.OST_DuctCurves,
                                 BuiltInCategory.OST_FlexDuctCurves,
                                 BuiltInCategory.OST_DuctAccessory,
                                 BuiltInCategory.OST_DuctFitting,
                                 BuiltInCategory.OST_DuctTerminal,
                                 BuiltInCategory.OST_PipeCurves,
                                 BuiltInCategory.OST_FlexPipeCurves,
                                 BuiltInCategory.OST_PipeAccessory,
                                 BuiltInCategory.OST_PipeFitting,
                                 BuiltInCategory.OST_MechanicalEquipment,
                                 BuiltInCategory.OST_Sprinklers,
                                 BuiltInCategory.OST_PlumbingFixtures,
                                 BuiltInCategory.OST_DuctInsulations,
                                 BuiltInCategory.OST_PipeInsulations
                                 ]
        self._builtInCat_SSEOM = [BuiltInCategory.OST_CableTray,
                                  BuiltInCategory.OST_Conduit,
                                  BuiltInCategory.OST_ElectricalEquipment,
                                  BuiltInCategory.OST_CableTrayFitting,
                                  BuiltInCategory.OST_ConduitFitting,
                                  BuiltInCategory.OST_LightingFixtures,
                                  BuiltInCategory.OST_ElectricalFixtures,
                                  BuiltInCategory.OST_DataDevices,
                                  BuiltInCategory.OST_LightingDevices,
                                  BuiltInCategory.OST_NurseCallDevices,
                                  BuiltInCategory.OST_SecurityDevices,
                                  BuiltInCategory.OST_FireAlarmDevices,
                                  ]

    def getOVVK(self):
        itemColl = elGetter(self.__doc,
                            self._builtInCat_OVVK,
                            self._view,
                            self.__onModel,
                            self.__onView
                            )
        return itemColl

    def getSSEOM(self):
        itemColl = elGetter(self.__doc,
                            self._builtInCat_SSEOM,
                            self._view,
                            self.__onModel,
                            self.__onView
                            )
        return itemColl

