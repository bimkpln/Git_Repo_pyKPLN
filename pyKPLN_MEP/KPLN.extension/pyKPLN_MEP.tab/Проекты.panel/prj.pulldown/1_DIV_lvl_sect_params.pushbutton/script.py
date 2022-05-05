# -*- coding: utf-8 -*-

__title__ = "СИТИ_Заполнение параметров 'Этаж', 'Секция'"
__author__ = 'Kapustin Roman'
__doc__ = "Автоматическое заполнение параметров для классификатора.\n"\
        "Заполянет параметры 'Этаж', 'Секция' для элементов ИОС, с условием взаимоисключения параметров 'Этаж' и 'Система'\n"\
        "Актуально для всех разделов ИОС!"

from Autodesk.Revit.DB import *
import System
from System import Enum, Guid
import re
from rpw import revit, db, ui
from pyrevit import script, forms
from rpw.ui.forms import*
# main code
doc = revit.doc
output = script.get_output()
# СИТИ_Система
guid_system_param = Guid("98b30303-a930-449c-b6c2-11604a0479cb")
# СИТИ_Этаж
guid_level_param = Guid("79c6dee3-bd92-4fdf-ae12-c06683c61946")
# СИТИ_Секция
guid_section_param = Guid("4e8c4436-b231-4f9c-a659-181a2b55eda4")
# Имя оси 
guid_grid_name = Guid("1cacba6e-69a7-4fb5-a10b-613ea01c1188")

# get all elements of current project
categoryList = [BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_FlexPipeCurves, BuiltInCategory.OST_FlexDuctCurves,
                BuiltInCategory.OST_MechanicalEquipment, BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_CableTray, BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ElectricalEquipment, BuiltInCategory.OST_CableTrayFitting,
                BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_ElectricalFixtures, BuiltInCategory.OST_DataDevices,
                BuiltInCategory.OST_LightingDevices, BuiltInCategory.OST_NurseCallDevices,
                BuiltInCategory.OST_SecurityDevices, BuiltInCategory.OST_FireAlarmDevices,
                BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_CommunicationDevices]
catARCHgrid = BuiltInCategory.OST_Grids
catARCHlvl = BuiltInCategory.OST_Levels
# _________________________________ОБЪЯВЛЕНИЕ КЛАССОВ________________________________________________________________________
class LvlBox:
    Sect = int()
    MinPointZ = float()
    Lvl = int()
class SectBox():
    def __init__(self, name):
        self.name = name  
    MinPointX= float()
    MinPointY= float()
    MaxPointX= float()
    MaxPointY= float()
    MinXYPoint = [MinPointX,MinPointY]
    MaxXYPoint = [MaxPointX,MaxPointY]
    GridSect = int()
class Box(LvlBox,SectBox): 
    MaxPointZ = float()
#________________________Функция проверки линий на пересечения __________________________________________________________________              
def line_intersection(line1, line2):
    xdiff = (line1[0][0] - line1[1][0], line2[0][0] - line2[1][0])
    ydiff = (line1[0][1] - line1[1][1], line2[0][1] - line2[1][1]) #Typo was here
    def det(a, b):
        return a[0] * b[1] - a[1] * b[0]
    div = det(xdiff, ydiff)
    if div == 0:
        return None
    d = (det(*line1), det(*line2))
    x = det(d, xdiff) / div
    y = det(d, ydiff) / div
    return x, y

linkModelInstances = FilteredElementCollector(doc).OfClass(RevitLinkInstance)
linkModelsCord = [i for i in linkModelInstances if 'Разб'.upper() in i.Name.upper()]
if len(linkModelsCord) < 1:
    ui.forms.Alert("В проекте нет разбивочного файла, обратитесь в BIM отдел!", title="Ошибка")
    script.exit()
elif len(linkModelsCord) > 1:
    ui.forms.Alert("В проекте больше 1 разбивочного файла, обратитесь в BIM отдел!", title="Ошибка")
    script.exit()
else:
    linkModelsCord = linkModelsCord[0]
try:
    lvlInRpoject = FilteredElementCollector(linkModelsCord.GetLinkDocument()).OfCategory(catARCHlvl).WhereElementIsNotElementType().ToElements()
except:
    ui.forms.Alert("В проекте нет разбивочного файла!", title="Ошибка")
    script.exit()
gridsInRpoject = FilteredElementCollector(linkModelsCord.GetLinkDocument()).OfCategory(catARCHgrid).WhereElementIsNotElementType().ToElements()
project_base_point_Arch =FilteredElementCollector(linkModelsCord.GetLinkDocument()).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()

try:
    for base_point_Arch in project_base_point_Arch:
        boundBoxPoint = XYZ(round(base_point_Arch.get_BoundingBox(doc.ActiveView).Max.X, 7),round(base_point_Arch.get_BoundingBox(doc.ActiveView).Max.Y, 7),round(base_point_Arch.get_BoundingBox(doc.ActiveView).Max.Z, 7))
except:
    boundBoxPoint = XYZ(round(project_base_point_Arch.get_BoundingBox(doc.ActiveView).Max.X, 7),round(project_base_point_Arch.get_BoundingBox(doc.ActiveView).Max.Y, 7),round(project_base_point_Arch.get_BoundingBox(doc.ActiveView).Max.Z, 7))
BoxObjList = []
LvlBoxObjList = []
uncurrectLvlsFlag = True
for curLvl in lvlInRpoject:
    noSectFlag = False
    lvlName = curLvl.Name
    if '_' in lvlName:
        lvlNameSplited = lvlName.split('_')
    else:
        continue
    try:
        lvlVal = int(lvlNameSplited[1])
        uncurrectLvlsFlag = False
    except:
        try:
            noSectFlag = True
            lvlVal = int(lvlNameSplited[0])
            uncurrectLvlsFlag = False
        except:
            continue
    if 'кровля' in str(lvlVal).lower():
        continue
    if len(lvlNameSplited) < 2:
        continue
# ________________________ЕСЛИ НУЖНО ИСКЛЮЧИТЬ УРОВНИ КРОВЛИ, ДОБАВЬ ЗАКОММЕНТИРОВАННОЕ НИЖЕ________________________________________________________
    # lvlValAfter  = lvlNameSplited[2]
    # if 'Кровля'.upper() in lvlVal.upper() or 'Кровля'.upper() in lvlValAfter.upper():
    #     continue
    # lvlZCord = curLvl.Elevation
# ___________________________________________________Генерация списка экземпляров класса LvlBox_______________________________________________________
    if not noSectFlag:
        lvlSectVal = lvlNameSplited[0]
        shortLvlName = lvlName[:lvlName.find('_')] 
        if 'С' in lvlName.upper():
            lvlCurSplitName = shortLvlName.upper().split('С')[1]
        elif 'C' in lvlName.upper():
            lvlCurSplitName = shortLvlName.upper().split('C')[1]
        elif 'ПАР' in lvlName.upper():
            lvlCurSplitName = shortLvlName.upper().split('ПАР')[1]
        else:
            lvlCurSplitName = lvlName
        if '.' in shortLvlName:
            lvlCurSplited = lvlCurSplitName.split('.') 
        elif '-' in shortLvlName:
            lvlCurSplited = []
            lvlCurSplitedItem = lvlCurSplitName.split('-')
            for longRangeItem in range(int(lvlCurSplitedItem[0]),int(lvlCurSplitedItem[1])+1):
                lvlCurSplited.append(str(longRangeItem))
        else:
            lvlCurSplited = [str(lvlCurSplitName)]
    else:
        lvlCurSplited = ['1','2','3','4','5','6','7','8','9','10','11','12','13','14','15']
    for SectForLvlBox in lvlCurSplited:
        lvlBoxObj = LvlBox()
        lvlBoxObj.MinPointZ = curLvl.Elevation - 0.66 #ЗАДАЕТСЯ СМЕЩЕНИЕ НА 200ММ ДЛЯ УЧЕТА ЭЛЕМЕНТОВ В СТЯЖКЕ
        lvlBoxObj.Sect = int(SectForLvlBox)
        lvlBoxObj.Lvl = int(lvlVal)
        LvlBoxObjList.append(lvlBoxObj)
if uncurrectLvlsFlag:
    ui.forms.Alert("Не настроен разбивочный файл, обратитесь в BIM отдел! Некорректные уровни!", title="Ошибка!")
    script.exit()
# ПРЕДПОЛОЖЕНО, ЧТО СЕКЦИЙ БУДЕТ НЕ БОЛЬШЕ 15!!!
fourGridList = []
fourGridColisionList = []
SectBoxObjList = []
for intSectVal in range(15): 
    for grid in gridsInRpoject:

#_______________________________________________________КОЛИБРОВКА КООРДИНАТ ОСЕЙ_________________________________________________________________
        if grid.Name.upper() == 'А' or grid.Name.upper() == 'A':
            AGridCordArch = grid.GetExtents().MinimumPoint
        if grid.Name.upper() == '1':
            OneGridCordArch = grid.GetExtents().MinimumPoint
# ___________________________________________________Генерация списка экземпляров класса SectBox______________________________________________________
        try:
            gridName = grid.get_Parameter(guid_grid_name).AsString()
        except:
            ui.forms.Alert("Не настроен разбивочный файл, обратитесь в BIM отдел! (Нет параметра)!", title="Ошибка!")
            script.exit()
        if str(intSectVal) in str(gridName):
            fourGridList.append(grid)
        if len(fourGridList) == 4:
            for FirtOfFoutGrig in fourGridList:
                for SecondOfFoutGrig in fourGridList:
                    lineFirstStartX = round(FirtOfFoutGrig.GetExtents().MinimumPoint.X,10)
                    lineFirstStartY = round(FirtOfFoutGrig.GetExtents().MinimumPoint.Y,10)
                    lineFirstEndX = round(FirtOfFoutGrig.GetExtents().MaximumPoint.X,10)
                    lineFirstEndY = round(FirtOfFoutGrig.GetExtents().MaximumPoint.Y,10)
                    lineSecondStartX = round(SecondOfFoutGrig.GetExtents().MinimumPoint.X,10)
                    lineSecondStartY = round(SecondOfFoutGrig.GetExtents().MinimumPoint.Y,10)
                    lineSecondEndX = round(SecondOfFoutGrig.GetExtents().MaximumPoint.X,10)
                    lineSecondEndY = round(SecondOfFoutGrig.GetExtents().MaximumPoint.Y,10)
                    
                    if lineFirstStartY == lineFirstEndY:
                        if lineFirstStartX > lineFirstEndX:
                            add = -500
                        else:
                            add = 500
                        lineFirstStartX = lineFirstStartX + add
                        lineFirstEndX = lineFirstEndX  - add
                    elif lineFirstStartX == lineFirstEndX:
                        if lineFirstStartY > lineFirstEndY:
                            add = -500
                        else:
                            add = 500
                        lineFirstStartY = lineFirstStartY +add
                        lineFirstEndY = lineFirstEndY - add
                    if lineSecondStartY == lineSecondEndY:
                        if lineSecondStartX > lineSecondEndX:
                            add = -500
                        else:
                            add = 500
                        lineSecondStartX = lineSecondStartX + add
                        lineSecondEndX = lineSecondEndX - add
                    elif lineSecondStartX == lineSecondEndX:
                        if lineFirstStartY > lineSecondEndY:
                            add = -500
                        else:
                            add = 500
                        lineSecondStartY = lineSecondStartY + add
                        lineSecondEndY = lineSecondEndY - add

                    lineFirstStart = (lineFirstStartX,lineFirstStartY)
                    lineFirstEnd = (lineFirstEndX,lineFirstEndY)
                    lineSecondStart = (lineSecondStartX,lineSecondStartY)
                    lineSecondEnd = (lineSecondEndX,lineSecondEndY)
                    lineFirst = (lineFirstStart,lineFirstEnd)
                    lineSecond = (lineSecondStart,lineSecondEnd)
                    lineIntersection = line_intersection(lineFirst,lineSecond)
                    if lineIntersection:
                        if float(str(lineIntersection)[1:-1].split(',')[0]) > 100000 or float(str(lineIntersection)[1:-1].split(',')[1]) > 100000:
                            lineIntersection = None
                    if lineIntersection and list(lineIntersection) not in fourGridColisionList:
                        fourGridColisionList.append(list(lineIntersection))
            SectObjBox = SectBox(intSectVal)
            SectObjBox.MinXYPoint = min(list(fourGridColisionList))
            SectObjBox.MaxXYPoint = max(list(fourGridColisionList))
# ____________________________________________Точка проверки мин и мах координат X, Y для boundingBox_______________________________________________________________
            # print(min(list(fourGridColisionList)))
            # print(max(list(fourGridColisionList)))
            # print('---------------------------'+str(intSectVal))
# _________________________________________________________________________________________________________________________________________________________________
            SectObjBox.GridSect = intSectVal
            SectBoxObjList.append(SectObjBox)
            fourGridList = []
            fourGridColisionList = []
    if len(fourGridList) != 0:
        ui.forms.Alert("Не настроен разбивочный файл, обратитесь в BIM отдел!", title="Ошибка!")
        script.exit()
BoundingBoxObjList = []
for ObjFromLvlBox in LvlBoxObjList:
    for ObjFromSectBox in SectBoxObjList:
        if ObjFromSectBox.name == ObjFromLvlBox.Sect:
            BoundingBoxObj = Box(str(ObjFromLvlBox.Lvl)+';'+str(ObjFromSectBox.name))
            BoundingBoxObj.Sect = ObjFromSectBox.name
            BoundingBoxObj.Lvl = ObjFromLvlBox.Lvl
            BoundingBoxObj.MaxXYPoint = ObjFromSectBox.MaxXYPoint
            BoundingBoxObj.MinXYPoint = ObjFromSectBox.MinXYPoint
            BoundingBoxObj.MinPointZ = ObjFromLvlBox.MinPointZ
            BoundingBoxObjList.append(BoundingBoxObj)
for FirstBoundingBoxObjToCordZ in BoundingBoxObjList:
    if FirstBoundingBoxObjToCordZ.Lvl == -1:
        excrVAl = 2
    else:
        excrVAl = 1
    for SecondBoundingBoxObjToCordZ in BoundingBoxObjList:
        if (SecondBoundingBoxObjToCordZ.Lvl == FirstBoundingBoxObjToCordZ.Lvl + excrVAl) and (SecondBoundingBoxObjToCordZ.Sect == FirstBoundingBoxObjToCordZ.Sect):
            FirstBoundingBoxObjToCordZ.MaxPointZ = SecondBoundingBoxObjToCordZ.MinPointZ
for lastANDungraundLvls in BoundingBoxObjList:
    if '-' in str(lastANDungraundLvls.Lvl):
        lastANDungraundLvls.MinPointZ = -100
    elif lastANDungraundLvls.MaxPointZ == 0:
        lastANDungraundLvls.MaxPointZ = 150
# ____________________________________________Точка проверки мин и мах координаты Z для boundingBox_________________________________________
    # print(lastANDungraundLvls.MinPointZ,lastANDungraundLvls.MaxPointZ)
    # print(str(lastANDungraundLvls.Sect) +'--------------------|'+ str(lastANDungraundLvls.Lvl))
# _______________________________________________________ПРОВЕРКА КООРДИНАТ_____________________________________________________________________
DeltaX = 0
DeltaY = 0
gridsInEngRpoject = FilteredElementCollector(doc).OfCategory(catARCHgrid).WhereElementIsNotElementType().ToElements()
project_base_point_Enj =FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()
for base_point_Enj in project_base_point_Enj:
    boundBoxPointEnj = XYZ(round(base_point_Enj.get_BoundingBox(doc.ActiveView).Max.X, 7),round(base_point_Enj.get_BoundingBox(doc.ActiveView).Max.Y, 7),round(base_point_Enj.get_BoundingBox(doc.ActiveView).Max.Z, 7))
if str(boundBoxPointEnj) == str(boundBoxPoint):
    for gridInEngPrj in gridsInEngRpoject:
        if gridInEngPrj.Name.upper() == 'A' or gridInEngPrj.Name.upper() == 'А':
            AGridCord = gridInEngPrj.GetExtents().MinimumPoint
            DeltaY = AGridCordArch.Y - AGridCord.Y
        if gridInEngPrj.Name.upper() == '1':
            OneGridCord = gridInEngPrj.GetExtents().MinimumPoint
            DeltaX = OneGridCordArch.X - OneGridCord.X
        # if DeltaX > 10 or DeltaY > 10:
        #     ui.forms.Alert("Проблемы с координатами, обратитесь в BIM отдел!", title="Ошибка!")
        #     script.exit()
else:
    ui.forms.Alert("Проблемы с координатами, обратитесь в BIM отдел!", title="Ошибка!")
    script.exit()
# _________________________________________________________MAIN CODE______________________________________________________________
errorCount = 0
OutOfModelElList = []
OutOfModelAutoSectLvlFindFlag = False
errorEl = []
with db.Transaction("СИТИ_Заполнение пар-ров этажа и секции"):
    for curCat in categoryList:
        elementsList = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(curCat).WhereElementIsNotElementType().ToElements()
        for element in elementsList:
            SysFlag = False
            try:
                sysFilter = element.get_Parameter(guid_system_param).AsString()
                lvlFilter = element.get_Parameter(guid_level_param).AsString()
            except:
                output.print_md(
                '''У элемента **{} с id {}** нет параметров секции и системы '''
                '''скрипта остановлена. **Обратитесь в BIM-отдел!**'''.
                format(element.Name,output.linkify(element.Id))
                )
                script.exit()
            if sysFilter:
                if lvlFilter:
                    output.print_md('У элемента **{} с id {}** заполнена и система и этаж!'.format(element.Name,output.linkify(element.Id)))
                    errorEl.append(element)
                SysFlag = True
                OutOfModelFlag = False
            try:
                superComponent = element.SuperComponent
            except:
                superComponent = None
            if superComponent is None:
                OutOfModelFlag = True
                try:
                    centerPoint = XYZ(element.GetSpatialElementCalculationPoint().X + DeltaX,element.GetSpatialElementCalculationPoint().Y + DeltaY,element.GetSpatialElementCalculationPoint().Z)
                except:
                    try:
                        centerPoint = XYZ(((element.BoundingBox[doc.ActiveView].Max.X + element.BoundingBox[doc.ActiveView].Min.X) / 2)+DeltaX, ((element.BoundingBox[doc.ActiveView].Max.Y
                                    + element.BoundingBox[doc.ActiveView].Min.Y) / 2)+DeltaY, (element.BoundingBox[doc.ActiveView].Max.Z + element.BoundingBox[doc.ActiveView].Min.Z) / 2)
                    except:
                        try:
                            centerPoint = XYZ(element.Location.Point.X + DeltaX,element.Location.Point.Y + DeltaY,element.Location.Point.Z)
                        except:
                            output.print_md('Элемент **с id {}** левый! Проверьте по id!'.format(output.linkify(element.Id)))
                            errorEl.append(element)

                # print(centerPoint)
                for BoindingBox in BoundingBoxObjList:
                    MinCordX = BoindingBox.MinXYPoint[0]
                    MinCordY = BoindingBox.MinXYPoint[1]
                    MinCordZ = BoindingBox.MinPointZ
                    MaxCordX = BoindingBox.MaxXYPoint[0]
                    MaxCordY = BoindingBox.MaxXYPoint[1]
                    MaxCordZ = BoindingBox.MaxPointZ
                    CurCordX = float(centerPoint.X)
                    CurCordY = float(centerPoint.Y)
                    CurCordZ = float(centerPoint.Z)
                    if MinCordX <= CurCordX and MaxCordX > CurCordX:
                        if MinCordY <= CurCordY and MaxCordY > CurCordY:
                            if MinCordZ <= CurCordZ and MaxCordZ > CurCordZ:                              
                                if '-' in str(BoindingBox.Lvl):
                                    try:
                                        element.get_Parameter(guid_section_param).Set('1')
                                        if not SysFlag:
                                            element.get_Parameter(guid_level_param).Set('м1')
                                    except:
                                        output.print_md(
                                            '''У элемента **{} с id {}** нет параметров, работа '''
                                            '''скрипта остановлена. **Обратитесь в BIM-отдел!**'''.
                                            format(element.Name,output.linkify(element.Id))
                                            )
                                        script.exit()
                                else:
                                    try:
                                        element.get_Parameter(guid_section_param).Set(str(BoindingBox.Sect))
                                        if not SysFlag:
                                            element.get_Parameter(guid_level_param).Set(str(BoindingBox.Lvl))
                                    except:
                                        output.print_md(
                                            '''У элемента **{} с id {}** нет параметров, работа '''
                                            '''скрипта остановлена. **Обратитесь в BIM-отдел!**'''.
                                            format(element.Name,output.linkify(element.Id))
                                            )
                                        script.exit()
                                OutOfModelFlag = False
                                SysFlag = False
                                break
                if OutOfModelFlag:
                    OutOfModelElList.append(element)
                    output.print_md('Элемент **{} с id {}** вне модели! Проверьте вручную!'.format(element.Name,output.linkify(element.Id)))
                    errorCount +=1
                    OutOfModelFlag = True
    flagOut = False
    if errorCount > 0:
        commands= [CommandLink('ДА', return_value='да'),CommandLink('НЕТ', return_value='нет')]   
        dialog = TaskDialog('В проекте есть элементы вне осей.\nОпределить положение аналитически?', commands=commands)
        dialog_out = dialog.show()
        flagOut = True
        if dialog_out == 'да':
            OutOfModelAutoSectLvlFindFlag = True
        elif dialog_out == 'нет':
            OutOfModelAutoSectLvlFindFlag = False
        else:
            script.exit()  
    if errorCount == 0:
        ui.forms.Alert("Завершено!", title="Завершено")
    else:
        print('Количество ошибок - {} шт.'.format(errorCount))
        if dialog_out == 'да':
            print('_________________________________________________________________________')
            print('Результаты аналитического поиска этажа и секции:')  
        #_______________________________________________________АНАЛИТИЧЕСКИЙ ПОИСК СЕКЦИИ И ЭТАЖА ДЛЯ ЭЛ ВНЕ МОДЕЛИ___________________________________________________________________ 
        if OutOfModelAutoSectLvlFindFlag:
            for element in OutOfModelElList:   
                SysFlag = False
                sysFilter = element.get_Parameter(guid_system_param).AsString()
                if sysFilter:
                    SysFlag = True
                ElFoungSectLvlFlag = False
                for delta in [[0,2],[0,-2],[2,0],[-2,0]]:
                    try:
                        superComponent = element.SuperComponent
                    except:
                        superComponent = None
                    if superComponent is None and not ElFoungSectLvlFlag:
                        OutOfModelFlag = True
                        try:
                            centerPoint = XYZ(element.GetSpatialElementCalculationPoint().X + DeltaX + delta[0],element.GetSpatialElementCalculationPoint().Y + delta[1] + DeltaY,element.GetSpatialElementCalculationPoint().Z)
                        except:
                                try:
                                    centerPoint = XYZ(((element.BoundingBox[doc.ActiveView].Max.X + element.BoundingBox[doc.ActiveView].Min.X) / 2)+DeltaX + delta[0], ((element.BoundingBox[doc.ActiveView].Max.Y
                                                + element.BoundingBox[doc.ActiveView].Min.Y) / 2)+DeltaY + delta[1], (element.BoundingBox[doc.ActiveView].Max.Z + element.BoundingBox[doc.ActiveView].Min.Z) / 2)
                                except:
                                    centerPoint = XYZ(element.Location.Point.X + DeltaX + delta[0],element.Location.Point.Y + DeltaY + delta[1],element.Location.Point.Z)
                        for BoindingBox in BoundingBoxObjList:
                            MinCordX = BoindingBox.MinXYPoint[0]
                            MinCordY = BoindingBox.MinXYPoint[1]
                            MinCordZ = BoindingBox.MinPointZ
                            MaxCordX = BoindingBox.MaxXYPoint[0]
                            MaxCordY = BoindingBox.MaxXYPoint[1]
                            MaxCordZ = BoindingBox.MaxPointZ
                            CurCordX = float(centerPoint.X)
                            CurCordY = float(centerPoint.Y)
                            CurCordZ = float(centerPoint.Z)
                            if MinCordX <= CurCordX and MaxCordX > CurCordX:
                                if MinCordY <= CurCordY and MaxCordY > CurCordY:
                                    if MinCordZ <= CurCordZ and MaxCordZ > CurCordZ:
                                        if '-' in str(BoindingBox.Lvl):
                                            try:
                                                element.get_Parameter(guid_section_param).Set('1')
                                                if not SysFlag:
                                                    element.get_Parameter(guid_level_param).Set('м1')
                                            except:
                                                output.print_md(
                                                    '''У элемента **{} с id {}** нет параметров, работа '''
                                                    '''скрипта остановлена. **Обратитесь в BIM-отдел!**'''.
                                                    format(element.Name,output.linkify(element.Id))
                                                    )
                                                script.exit()
                                        else:
                                            try:
                                                element.get_Parameter(guid_section_param).Set(str(BoindingBox.Sect))
                                                if not SysFlag:
                                                    element.get_Parameter(guid_level_param).Set(str(BoindingBox.Lvl))
                                            except:
                                                output.print_md(
                                                    '''У элемента **{} с id {}** нет параметров, работа '''
                                                    '''скрипта остановлена. **Обратитесь в BIM-отдел!**'''.
                                                    format(element.Name,output.linkify(element.Id))
                                                    )
                                                script.exit()
                                        OutOfModelFlag = False
                                        ElFoungSectLvlFlag = True
                                        errorCount -=1
                                        break
                if OutOfModelFlag:
                    output.print_md('Элемент **{} с id {}** так и не нашел основу:( Проверьте вручную!'.format(element.Name,output.linkify(element.Id)))
                    OutOfModelFlag = True

if flagOut:
    if dialog_out == 'да':
        if errorCount == 0 and len(errorEl)==0:
            print('Ошибок нет, устранено ошибок: - {} шт.'.format(len(OutOfModelElList)))
            ui.forms.Alert("Завершено! Все ошибки исправлены!", title="Завершено")
        else:
            for err in errorEl:
                output.print_md('Элемент **{} с id {}** Либо левый, либо имеет и систему и этаж!'.format(err.Name,output.linkify(err.Id)))
            print('Так и не решилось ошибок: - {} шт.'.format(errorCount+len(errorEl)))
