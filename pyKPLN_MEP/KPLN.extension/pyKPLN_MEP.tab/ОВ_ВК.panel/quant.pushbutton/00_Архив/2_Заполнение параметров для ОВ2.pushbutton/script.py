# coding: utf-8

__title__ = "ОВ2. Заполнение параметров для спецификации по нескольким категориям v1.1"
__author__ = 'Kapustin Roman'
__doc__ = "Заполняет:\n 1.Для воздуховодв длину и габариты, считая воздуховоды MxN и NxM как MxN \n 2.Для элементов в шт. число. \n 3.Для изоляции площадь и толщину стенки.\n 4. Единицы измерения \n 5. Запас для линейных элементов \n 6. Порядок по категориям для сортировки"


from Autodesk.Revit.Exceptions import InvalidOperationException
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory,\
    BuiltInParameter, Mechanical, ConnectorProfileType, XYZ
import System
import enum
from rpw import revit, db
from pyrevit import script
from System import Guid
from rpw.ui.forms import CommandLink, TaskDialog, Label, CheckBox, TextBox,\
    Separator, Button, FlexForm, Alert


# main code: input part
commands = [CommandLink('Электролитный', return_value='Электролитный'),
            CommandLink('Новые проекты', return_value='Новые проекты')]
dialog = TaskDialog('Выберите тип файла:',
                    commands=commands,
                    show_close=True)
dialog_out = dialog.show()

if dialog_out == 'Новые проекты':
    # КП_И_Запас
    supply = Guid("62ff7f1d-f5cd-4e5e-9f50-b7b3e18f2c06")
    # КП_И_КолСпецификация
    guidLengthParam = Guid("df684fd7-6956-4d85-8475-709481838383")
    # КП_Размер_Текст
    guidSizeParam = Guid("7d38dcf5-47ef-46d0-b4a8-9babee9cc11b")
    # КП_И_Толщина стенки
    guidFitSizeParam = Guid("381b467b-3518-42bb-b183-35169c9bdfb3")
    # КП_О_Единица измерения
    guidUnitParam = Guid("4289cb19-9517-45de-9c02-5a74ebf5c86d")
    # КП_И_Порядковый номер
    guidNumParam = Guid("7a44592a-c698-4276-af5f-75be9b131b40")
    # КП_О_Сортировка
    guidSortParam = Guid("82fad8c7-4f46-4da8-a69a-035a34b5f015")
elif dialog_out == 'Электролитный':
    supply = "Запас"
    # КП_Длина погонная - параметр проекта
    guidLengthParam = Guid("2aa6d2a0-e0a6-468e-8a04-e1d10daea4f4")
    # КП_Размер_Текст
    guidSizeParam = Guid("7d38dcf5-47ef-46d0-b4a8-9babee9cc11b")
    # КП_И_Толщина стенки
    guidFitSizeParam = Guid("381b467b-3518-42bb-b183-35169c9bdfb3")
    # КП_О_Единица измерения
    guidUnitParam = Guid("4289cb19-9517-45de-9c02-5a74ebf5c86d")
    # КП_И_Порядковый номер
    guidNumParam = Guid("7a44592a-c698-4276-af5f-75be9b131b40")
    # КП_О_Сортировка
    guidSortParam = Guid("82fad8c7-4f46-4da8-a69a-035a34b5f015")
else:
    script.exit()

catLitsRange = [
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_FlexDuctCurves]

catList_quantity = [
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_DuctTerminal,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctFitting]

catListInsulation = [
    BuiltInCategory.OST_DuctLinings,
    BuiltInCategory.OST_DuctInsulations]

doc = revit.doc
output = script.get_output()
rectDuctSepar = Mechanical\
    .DuctSettings\
    .GetDuctSettings(doc)\
    .RectangularDuctSizeSeparator
if not rectDuctSepar:
    output.print_md("В проекте отсутсвует разделитель для воздуховодов")
    script.exit()

ductConnectorSepar = Mechanical\
    .DuctSettings\
    .GetDuctSettings(doc)\
    .ConnectorSeparator
if not ductConnectorSepar:
    output.print_md("В проекте отсутсвует разделитель для соед. воздуховодов")
    script.exit()

with db.Transaction('Заполнение корректной длины для спецификации по нескольким категориям'):
    components = [
        Label('Выберите параметры для заполнения:'),
        CheckBox('checkbox1', 'КП_И_Количество в спецификацию', default=True),
        CheckBox('checkbox2', 'КП_Размер_Текст', default=True),
        CheckBox('checkbox3', 'КП_И_Порядковый номер', default=True),
        CheckBox('checkbox4', 'КП_И_Запас', default=True),
        Label('Размер запаса:'),
        TextBox('textbox1', Text="1.15"),
        CheckBox('checkbox5', 'КП_И_Толщина стенки', default=True),
        CheckBox('checkbox6', 'КП_О_Единица измерения', default=True),
        CheckBox('checkbox7', 'КП_О_Сортирование', default=False),
        Separator(),
        Button('Далее')]
    form = FlexForm('ОВ_Спецификации', components)
    form.show()
    supply_value = float(form.values.get('textbox1'))
    if supply_value < 1 or supply_value >= 2:
        Alert('Ввден некорректный запас!',
              title='Ошибка',
              header='',
              exit=True)
    checkbox1 = form.values.get('checkbox1')
    checkbox2 = form.values.get('checkbox2')
    checkbox3 = form.values.get('checkbox3')
    checkbox4 = form.values.get('checkbox4')
    checkbox5 = form.values.get('checkbox5')
    checkbox6 = form.values.get('checkbox6')
    checkbox7 = form.values.get('checkbox7')

    if form.values != {}:
        # Перенос длины в категории в метрах
        # Воздуховоды, гибкие воздуховоды
        for cat in catLitsRange:
            elList_range = FilteredElementCollector(doc)\
                .OfCategory(cat)\
                .WhereElementIsNotElementType()\
                .ToElements()

            for el in elList_range:
                if checkbox1:
                    elRange_range = el\
                        .get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)\
                        .AsValueString()

                    Range = (float(elRange_range))/1000
                    if dialog_out == 'Электролитный':
                        el.get_Parameter(guidLengthParam).Set(Range/304.8)
                    else:
                        el.get_Parameter(guidLengthParam).Set(Range)

            # Запись в параметр КП_Размер_Текст размера воздуховодов
                if checkbox2:
                    elRangeSizeDuct = el\
                        .get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)\
                        .AsString()

                    if rectDuctSepar in elRangeSizeDuct:
                        sizeDuctFit = elRangeSizeDuct\
                            .split(rectDuctSepar)
                        duct_width = sizeDuctFit[0]
                        duct_height = sizeDuctFit[1]
                        if int(duct_width) > int(duct_height):
                            duct_width_correct = duct_width
                            duct_height_correct = duct_height
                        else:
                            duct_width_correct = duct_height
                            duct_height_correct = duct_width
                        sizeDuctFit = duct_width_correct\
                            + "x"\
                            + duct_height_correct

                        el\
                            .get_Parameter(guidSizeParam)\
                            .Set(sizeDuctFit)
                    else:
                        el\
                            .get_Parameter(guidSizeParam)\
                            .Set(elRangeSizeDuct)
                if checkbox4:

                    # Запись запаса для длины
                    try:
                        if type(supply) == System.String:
                            new_supply_duct = doc\
                                .GetElement(el.GetTypeId())\
                                .LookupParameter(supply)\
                                .AsDouble()
                        else:
                            new_supply_duct = doc\
                                .GetElement(el.GetTypeId())\
                                .get_Parameter(supply)\
                                .AsDouble()
                    except AttributeError:
                        print(el.Id)
                    if new_supply_duct < 1:
                        new_supply_duct = supply_value
                        doc\
                            .GetElement(el.GetTypeId())\
                            .get_Parameter(supply)\
                            .Set(new_supply_duct)

                # Запись единиц измерения
                if checkbox6:
                    unitDuctParam = doc\
                        .GetElement(el.GetTypeId())\
                        .get_Parameter(guidUnitParam)

                    if unitDuctParam:
                        unit_duct = unitDuctParam.AsString()
                        if not unit_duct:
                            try:
                                doc\
                                    .GetElement(el.GetTypeId())\
                                    .get_Parameter(guidUnitParam)\
                                    .Set("м")
                            except AttributeError:
                                el.get_Parameter(guidUnitParam).Set("м")
                    else:
                        output.print_md(
                            'У элемента **{} с id {}** нет параметра КП_О_Единица измерения'
                            .format(el.Name, output.linkify(el.Id)))

                # Запись номера для спецификации
                if checkbox3:
                    num_duct = doc\
                        .GetElement(el.GetTypeId())\
                        .get_Parameter(guidNumParam)\
                        .AsDouble()

                    if num_duct < 1:
                        new_num_duct = 4
                        doc\
                            .GetElement(el.GetTypeId())\
                            .get_Parameter(guidNumParam)\
                            .Set(new_num_duct)

        # Задает 1 для категорий:
        # Оборудование, воздухораспределители, арматура воздуховодов,
        # соеденительные детали воздуховодов.
        num = 1
        for cat in catList_quantity:
            elListQuantity = FilteredElementCollector(doc)\
                .OfCategory(cat)\
                .WhereElementIsNotElementType()\
                .ToElements()

            for el in elListQuantity:
                if checkbox1:
                    if dialog_out == 'Электролитный':
                        el.get_Parameter(guidLengthParam).Set(1/304.8)
                    else:
                        el.get_Parameter(guidLengthParam).Set(1)

                # Запись в параметр КП_Размер_Текст размера соед. воздуховодов
                currentBic = BuiltInCategory\
                    .Parse(
                        BuiltInCategory,
                        el.Category.Id.IntegerValue.ToString())

                # Запись в параметр КП_Размер_Текст размера соед. воздуховодов
                if BuiltInCategory.OST_DuctFitting.__eq__(currentBic)\
                        and checkbox2:
                    elRangeSizeDuct = el\
                        .get_Parameter(BuiltInParameter.RBS_CALCULATED_SIZE)\
                        .AsString()
                    sizeDuctFit = elRangeSizeDuct\
                        .split(ductConnectorSepar)

                    trueSizeDuctFit = ""
                    connectManager = el.MEPModel.ConnectorManager
                    # Заглушки
                    if connectManager.Connectors.Size == 1:
                        for c in connectManager.Connectors:
                            if all(ConnectorProfileType.Round.__eq__(ConnectorProfileType.Parse(ConnectorProfileType, c.Shape.ToString())) for c in connectManager.Connectors):
                                currentRadius = c.Radius
                                trueSizeDuctFit = "ø{}".format(
                                    (currentRadius * 304.8 * 2).ToString())
                            elif all(ConnectorProfileType.Rectangular.__eq__(ConnectorProfileType.Parse(ConnectorProfileType, c.Shape.ToString())) for c in connectManager.Connectors):
                                trueSizeDuctFit = sizeDuctFit[0]

                    # Отводы, врезки переходы
                    elif connectManager.Connectors.Size == 2:
                        if len(sizeDuctFit) != 2:
                            print(
                                "Ошибка в применении разделителя соед. деталей {}. Нужно обновить семейство, чтобы разделитель присвоился".
                                format(el.Id))
                            continue

                        for c in connectManager.Connectors:
                            # Соединители круглые
                            if all(ConnectorProfileType.Round.__eq__(ConnectorProfileType.Parse(ConnectorProfileType, c.Shape.ToString())) for c in connectManager.Connectors):
                                currentRadius = c.Radius
                                # Отводы, врезки
                                if all(c.Radius == currentRadius for c in connectManager.Connectors):
                                    trueSizeDuctFit = sizeDuctFit[0]
                                    try:
                                        conAngle = max(c.Angle for c in connectManager.Connectors)
                                        if (conAngle != 0):
                                            trueSizeDuctFit = "ø{}, {}°".format(
                                                (currentRadius * 304.8 * 2).ToString(),
                                                round(conAngle / 0.01745, 0).ToString())
                                        else:
                                            trueSizeDuctFit = "ø{}".format(
                                                (currentRadius * 304.8 * 2).ToString())
                                    except Exception:
                                        trueSizeDuctFit = "ø{}".format(
                                            (currentRadius * 304.8 * 2).ToString())
                                    break
                                # Переходы
                                else:
                                    trueSizeDuctFit = "{}/{}".format(
                                        sizeDuctFit[0],
                                        sizeDuctFit[1])
                                    break
                            # Соединители квадратные
                            elif all(ConnectorProfileType.Rectangular.__eq__(ConnectorProfileType.Parse(ConnectorProfileType, c.Shape.ToString())) for c in connectManager.Connectors):
                                currentHeight = c.Height
                                currentWidth = c.Width
                                # Отводы, врезки
                                if all(c.Height == currentHeight for c in connectManager.Connectors)\
                                        and all(c.Width == currentWidth for c in connectManager.Connectors):
                                    trueSizeDuctFit = sizeDuctFit[0]
                                    try:
                                        conAngle = max(c.Angle for c in connectManager.Connectors)
                                        if (conAngle != 0):
                                            trueSizeDuctFit = "{}x{}, {}°".format(
                                                (currentWidth * 304.8 ).ToString(),
                                                (currentHeight * 304.8 ).ToString(),
                                                round(conAngle / 0.01745, 0).ToString())
                                        else:
                                            trueSizeDuctFit = "{}x{}".format(
                                                (currentWidth * 304.8 ).ToString(),
                                                (currentHeight * 304.8 ).ToString())
                                    except Exception:
                                        trueSizeDuctFit = "{}x{}".format(
                                                (currentWidth * 304.8 ).ToString(),
                                                (currentHeight * 304.8 ).ToString())
                                    break
                                # Переходы
                                else:
                                    trueSizeDuctFit = sizeDuctFit[0] + "/" + sizeDuctFit[1]
                                    break
                            # Соединители смешанные
                            else:
                                trueSizeDuctFit = sizeDuctFit[0] + "/" + sizeDuctFit[1]
                                break
                    # Тройники
                    elif connectManager.Connectors.Size == 3:
                        rectSizeList = []
                        lineConList = []
                        branchConList = []
                        for con in connectManager.Connectors:
                            conAngle = con.Angle
                            if conAngle == 0:
                                lineConList.append(con)
                            else:
                                branchConList.append(con)

                        # Проверка на корректные тройники (заполненность параметра Угол)
                        if len(branchConList) == 0:
                            print(
                                "Для тройников обязательно на ответвлениях указывать Угол для соединителя, иначе анализ невозможен. Элемент на котором ошибка - {}. Анализ тройника - **НЕ ВЫПОЛНЕН**!"
                                .format(el.Id))
                            trueSizeDuctFit = "ОШИБКА! Не указан Угол у соединителя"
                        else:
                            # Соединители круглые
                            if all(ConnectorProfileType.Round.__eq__(ConnectorProfileType.Parse(ConnectorProfileType, c.Shape.ToString())) for c in connectManager.Connectors):
                                lineRadius = lineConList[0].Radius
                                branchRadius = branchConList[0].Radius
                                trueSizeDuctFit = "ø{0}/ø{1}/ø{0}"\
                                    .format(
                                        (lineRadius * 304.8 * 2).ToString(),
                                        (branchRadius * 304.8 * 2).ToString())
                            # Соединители квадратные
                            else:
                                lineWidth = lineConList[0].Width
                                lineHeight = lineConList[0].Height
                                branchWidth = branchConList[0].Width
                                branchHeight = branchConList[0].Height
                                trueSizeDuctFit = "{0}x{1}/{2}x{3}/{0}x{1}"\
                                    .format(
                                        (lineWidth*304.8).ToString(),
                                        (lineHeight*304.8).ToString(),
                                        (branchWidth*304.8).ToString(),
                                        (branchHeight*304.8).ToString())

                    # Соединители смешанные
                    else:
                        output.print_md(
                            'Элемент **{} с id {}** не обработываемый. Скинь в BIM-отдел'
                            .format(el.Name, output.linkify(el.Id)))
                    el.get_Parameter(guidSizeParam).Set(trueSizeDuctFit)

                # Запись запаса для длины
                if checkbox4:
                    if type(supply) == System.String:
                        new_supply_quantity = doc\
                            .GetElement(el.GetTypeId())\
                            .LookupParameter(supply)\
                            .AsDouble()
                    else:
                        new_supply_quantity = doc\
                            .GetElement(el.GetTypeId())\
                            .get_Parameter(supply)\
                            .AsDouble()
                    if new_supply_quantity < 1:
                        new_supply_quantity = 1
                        if type(supply) == System.String:
                            doc\
                                .GetElement(el.GetTypeId())\
                                .LookupParameter(supply)\
                                .Set(new_supply_quantity)
                        else:
                            doc\
                                .GetElement(el.GetTypeId())\
                                .get_Parameter(supply)\
                                .Set(new_supply_quantity)

                # Запись единиц измерения
                if checkbox6:
                    unitQuantityParam = doc\
                        .GetElement(el.GetTypeId())\
                        .get_Parameter(guidUnitParam)
                    if unitQuantityParam:
                        unitQuantity = unitQuantityParam.AsString()
                        if not unitQuantity:
                            try:
                                doc\
                                    .GetElement(el.GetTypeId())\
                                    .get_Parameter(guidUnitParam)\
                                    .Set("шт.")
                            except AttributeError:
                                el.get_Parameter(guidUnitParam).Set("шт.")
                    else:
                        output.print_md(
                            'У элемента **{} с id {}** нет параметра КП_О_Единица измерения'
                            .format(el.Name, output.linkify(el.Id)))

                # Запись номера для спецификации
                if checkbox3:
                    num_all = doc\
                        .GetElement(el.GetTypeId())\
                        .get_Parameter(guidNumParam)\
                        .AsDouble()
                    if num_all < 1:
                        new_num_all = num
                        try:
                            doc\
                                .GetElement(el.GetTypeId())\
                                .get_Parameter(guidNumParam)\
                                .Set(new_num_all)
                        except:
                            output.print_md(
                                'У элемента **{} с id {}** параметр КП_И_Порядковый номер заблочен!'
                                .format(el.Name, output.linkify(el.Id)))
            if num < 3:
                num += 1
            else:
                num = 6

        # Задает для изоляции площадь:
        for cat in catListInsulation:
            elListInsulations = FilteredElementCollector(doc)\
                .OfCategory(cat)\
                .WhereElementIsNotElementType()\
                .ToElements()
            for el in elListInsulations:
                if checkbox1:
                    elRange_size_insulations = el\
                        .get_Parameter(BuiltInParameter.RBS_DUCT_CALCULATED_SIZE)\
                        .AsString()
                    elRange_range_insulations = el\
                        .get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)\
                        .AsValueString()
                    elRange_range_fiting = el\
                        .get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA)\
                        .AsDouble()
                    if elRange_range_fiting > 0:
                        if "x" in elRange_size_insulations:
                            elSize_rectangle = elRange_size_insulations.split('x')
                            square_rectangle = ((float(elSize_rectangle[0])/1000)+(float(elSize_rectangle[1])/1000))*2*(float(elRange_range_insulations))/1000
                            if dialog_out == 'Электролитный':
                                el.get_Parameter(guidLengthParam).Set(square_rectangle/304.8)
                            else:
                                el.get_Parameter(guidLengthParam).Set(square_rectangle)
                        else:
                            elSize_circle = elRange_size_insulations.split('ø')
                            square_circle = (float(elSize_circle[1])/1000)*3.14*(float(elRange_range_insulations))/1000
                            if dialog_out == 'Электролитный':
                                el\
                                    .get_Parameter(guidLengthParam)\
                                    .Set(square_circle/304.8)
                            else:
                                el\
                                    .get_Parameter(guidLengthParam)\
                                    .Set(square_circle)

                # Перенос толщины изоляции в КП_И_Толщина стенки
                if checkbox5:
                    if cat == BuiltInCategory.OST_DuctInsulations:
                        elRange_fit_insulations = el.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS_FOR_DUCT).AsValueString()
                    else:
                        elRange_fit_insulations = el.get_Parameter(BuiltInParameter.RBS_LINING_THICKNESS_FOR_DUCT).AsValueString()
                    elRange_fit_insulations_int = int(elRange_fit_insulations.split(' ')[0])/304.8
                    if elRange_fit_insulations_int > 0:
                        el.get_Parameter(guidFitSizeParam).Set(elRange_fit_insulations_int)

                # Запись запаса для длины
                if checkbox4:
                    if type(supply) == System.String:
                        new_supply_Insulations = doc.\
                                GetElement(el.GetTypeId()).\
                                LookupParameter(supply).\
                                AsDouble()
                    else:
                        new_supply_Insulations = doc.GetElement(el.GetTypeId()).get_Parameter(supply).AsDouble()
                    if new_supply_Insulations < 1:
                        new_supply_Insulations = supply_value
                        if type(supply) == System.String:
                            doc.\
                                GetElement(el.GetTypeId()).\
                                LookupParameter(supply).\
                                Set(new_supply_Insulations)
                        else:
                            doc.\
                                GetElement(el.GetTypeId()).\
                                get_Parameter(supply).\
                                Set(new_supply_Insulations)

                # Запись единиц измерения
                if checkbox6:
                    unit_Insulations = doc.GetElement(el.GetTypeId()).get_Parameter(guidUnitParam).AsString()
                    if not unit_Insulations:
                        try:
                            doc.GetElement(el.GetTypeId()).get_Parameter(guidUnitParam).Set("м2")
                        except AttributeError:
                            try:
                                el.get_Parameter(guidUnitParam).Set("м2")
                            except:
                                output.print_md('У элемента **{} с id {}** не заполнен ед. изм!'.format(el.Name, output.linkify(el.Id)))

                # Запись номера для спецификации
                if checkbox3:
                    num_Insulations = doc.GetElement(el.GetTypeId()).get_Parameter(guidNumParam).AsDouble()
                    if num_Insulations < 1:
                        new_num_Insulations = 5
                        doc.GetElement(el.GetTypeId()).get_Parameter(guidNumParam).Set(new_num_Insulations)

        # Запись в параметр КП_О_Сортировка сумму параметров Семейство,
        # Типоразмер, КП_Размер текст, КП_И_Толщина стенки
        if checkbox7:
            for all_cat in [catLitsRange,
                            catList_quantity,
                            catListInsulation]:
                for cat in all_cat:
                    elList_all = FilteredElementCollector(doc).OfCategory(cat).WhereElementIsNotElementType().ToElements()
                    for el in elList_all:
                        try:
                            el_name_sort = el.Name
                            el_family_name_sort = doc.GetElement(el.GetTypeId()).FamilyName
                            el_size_sort = el.get_Parameter(guidSizeParam).AsString()
                            try:
                                try:
                                    el_fit_sort = round(el.get_Parameter(guidFitSizeParam).AsDouble() * 304.8, 1).ToString()
                                    sort_param = el_family_name_sort+'/'+el_name_sort+'/'+el_size_sort+'/'+el_fit_sort
                                except:
                                    sort_param = el_family_name_sort+'/'+el_name_sort+'/'+el_size_sort
                            except:
                                try:
                                    sort_param = el_family_name_sort+'/'+el_name_sort+'/'+el_size_sort
                                except:
                                    sort_param = el_family_name_sort+'/'+el_name_sort

                            el.get_Parameter(guidSortParam).Set(sort_param)
                        except AttributeError as attr:
                            output.print_md(
                                "Ошибка {} у элемента {}. Работа остановлена!".
                                format(attr.ToString(), output.linkify(el.Id)))
                            script.exit()
