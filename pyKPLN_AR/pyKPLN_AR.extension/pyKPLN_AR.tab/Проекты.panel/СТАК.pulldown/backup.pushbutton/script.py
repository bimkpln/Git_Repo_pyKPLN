# -*- coding: utf-8 -*-

__author__ = 'Gyulnara Fedoseeva'
__title__ = "BIM - Backup"
__doc__ = 'Резервное копирование моделей с RevitServer на диск Y, обновление исходных данных по разделам КР, ИОС'\

"""
Архитекурное бюро KPLN

"""

import os
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.Exceptions import InvalidOperationException
from Autodesk.Revit.UI import TaskDialog, TaskDialogResult

from pyrevit import forms
from pyrevit.forms import ProgressBar

uiapp = __revit__
doc = uiapp.ActiveUIDocument.Document
app = doc.Application

kor_list = ["1.1", "1.2", "2.1", "2.2", "3.1", "3.2", "4.1", "4.2"]
ios_list = [["7.1.ЭОМ", "ЭОМ"], ["7.2.ВК", "ВК"], ["7.4.ОВ", "ОВиК"], ["7.5.СС", "СС"]]

class ErrorBox:
    """Класс-контейнер для хранения значений ошибок при открытии файла и инструкций к тому, как эти ошибки закрыть"""

    def __init__(self, name, overrideResult):
        # Имя окна
        self.name = name
        # Кнопка, которую необходимо нажать (TaskDialogResult)
        self.overrideResult = overrideResult

errorBoxes = [
    ErrorBox('', TaskDialogResult.Close),
    ErrorBox('TaskDialog_Audit_Warning', TaskDialogResult.Yes)]

def dialog_box_handler(sender, e):
    global errorBoxes
    currentUIApp = sender
    dialogBoxShowingEventArgs = e
    dialogStrId = dialogBoxShowingEventArgs.DialogId
    if dialogBoxShowingEventArgs.Cancellable:
        dialogBoxShowingEventArgs.Cancel()
    elif len(dialogStrId) < 1:
        dialogBoxShowingEventArgs.OverrideResult(int(errorBoxes[0].overrideResult))
    else:
        for errorBox in errorBoxes:
            if (dialogStrId == errorBox.name):
                dialogBoxShowingEventArgs.OverrideResult(int(errorBox.overrideResult))
            else:
                print("Новый тип диалогового окна, для которого нет инструкции по обработке: " + dialogStrId)

def copy_files(source_path, new_path):
    # Открываем файл Revit
    options = OpenOptions()
    options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
    options.SetOpenWorksetsConfiguration(WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets))
    opened_doc = app.OpenDocumentFile(source_path, options)

    # Сохраненяем файлы
    save_options = SaveAsOptions()
    save_options.OverwriteExistingFile = True

    ws_options = WorksharingSaveAsOptions()
    ws_options.SaveAsCentral = True

    save_options.SetWorksharingOptions(ws_options)
    opened_doc.SaveAs(new_path, save_options)
    opened_doc.Close()

def copy_ar_from_revitserver_to_localfolder():
    # Список файлов АР в исходной папке
    ar_files = ["RSN://192.168.0.5/СТАК/Стадия_Р/АР/К1.1/СТАК_К1.1_РД_АР_АР1.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К1.1/СТАК_К1.1_РД_АР_АР2_Фасад.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К1.1/СТАК_К1.1_РД_АР_РазбФайл.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К1.2/СТАК_К1.2_РД_АР_АР1.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К1.2/СТАК_К1.2_РД_АР_АР2_Фасад.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К1.2/СТАК_К1.2_РД_АР_РазбФайл.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К2.1/СТАК_К2.1_РД_АР_АР1.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К2.1/СТАК_К2.1_РД_АР_АР2_Фасад.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К2.1/СТАК_К2.1_РД_АР_РазбФайл.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К2.2/СТАК_К2.2_РД_АР_АР1.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К2.2/СТАК_К2.2_РД_АР_АР2_Фасад.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К2.2/СТАК_К2.2_РД_АР_РазбФайл.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К3.1/СТАК_К3.1_РД_АР_АР1.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К3.1/СТАК_К3.1_РД_АР_АР2_Фасад.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К3.1/СТАК_К3.1_РД_АР_РазбФайл.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К3.2/СТАК_К3.2_РД_АР_АР1.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К3.2/СТАК_К3.2_РД_АР_АР2_Фасад.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К3.2/СТАК_К3.2_РД_АР_РазбФайл.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К4.1/СТАК_К4.1_РД_АР_АР1.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К4.1/СТАК_К4.1_РД_АР_АР2_Фасад.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К4.1/СТАК_К4.1_РД_АР_РазбФайл.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К4.2/СТАК_К4.2_РД_АР_АР1.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К4.2/СТАК_К4.2_РД_АР_АР2_Фасад.rvt",
            "RSN://192.168.0.5/СТАК/Стадия_Р/АР/К4.2/СТАК_К4.2_РД_АР_РазбФайл.rvt"]

    try:
        for f in ar_files:
            ar_file_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(f)

            ar_new_file_path = "Y:\\Жилые здания\\Старый_Акульшет\\10.Стадия_Р\\5.АР\\1.RVT\\RevitServer_Архив" + "\\" + f.split('/')[-1]
            ar_new_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(ar_new_file_path)

            copy_files(ar_file_path, ar_new_path)
    except InvalidOperationException as ex:
        TaskDialog.Show('Ошибка', '{}'.format(str(ex)))

def copy_kr_from_localfolder_to_revitserver():
    try:
        selected_kor = forms.SelectFromList.show(kor_list,
                                                 multiselect=True,
                                                 title="Выберите корпус:",
                                                 button_name="ОК")
        for n in selected_kor:
            kr_files = [f for f in os.listdir("Y:\\Жилые здания\\Старый_Акульшет\\10.Стадия_Р\\6.КР\\1.RVT\Корпус {}".format(n)) if f.endswith('.rvt')]
            for r in kr_files:
                # Исключения для файлом КР
                if not "НЕ " in r and not "СБМ" in r:
                    kr_file_path = ModelPathUtils.ConvertUserVisiblePathToModelPath("Y:\\Жилые здания\\Старый_Акульшет\\10.Стадия_Р\\6.КР\\1.RVT\Корпус {}".format(n) + "\\" + r)

                    kr_new_file_path = "RSN://192.168.0.5/СТАК/Стадия_Р/КР/Корпус {}/".format(n) + r
                    kr_new_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(kr_new_file_path)

                    copy_files(kr_file_path, kr_new_path)
    except InvalidOperationException as ex:
        TaskDialog.Show('Ошибка', '{}'.format(str(ex)))

def copy_ios_from_localfolder_to_revitserver():
    try:
        sub_list = []
        for i in ios_list:
            sub_list.append(i[1])
        selected_sub = forms.SelectFromList.show(sub_list,
                                                 multiselect=True,
                                                 title="Выберите корпус:",
                                                 button_name="ОК")
        selected_kor = forms.SelectFromList.show(kor_list,
                                                 multiselect=True,
                                                 title="Выберите корпус:",
                                                 button_name="ОК")
        ios_files = []
        for i in ios_list:
            for sub in selected_sub:
                if i[1] == sub:
                    ios_files = [f for f in os.listdir("Y:\\Жилые здания\\Старый_Акульшет\\10.Стадия_Р\\" + i[0] + "\\1.RVT") if f.endswith('.rvt')]
                    for r in ios_files:
                        for kor in selected_kor:
                            if kor in r:
                                ios_file_path = ModelPathUtils.ConvertUserVisiblePathToModelPath("Y:\\Жилые здания\\Старый_Акульшет\\10.Стадия_Р\\" + i[0] + "\\1.RVT\\" + r)

                                ios_new_file_path = "RSN://192.168.0.5/СТАК/Стадия_Р/{}/".format(i[1]) + r
                                ios_new_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(ios_new_file_path)

                                copy_files(ios_file_path, ios_new_path)
    except InvalidOperationException as ex:
        TaskDialog.Show('Ошибка', '{}'.format(str(ex)))

# Подписываемся на событие всплывающего окна
uiapp.DialogBoxShowing += dialog_box_handler

selected_options = forms.CommandSwitchWindow.show(
    ["АР", "КР", "ИОС", "ВСЕ РАЗДЕЛЫ"],
    message='Выберите раздел:')
if selected_options == "АР":
    # Копирование файлов АР с RevitServer на диск Y
    copy_ar_from_revitserver_to_localfolder()
elif selected_options == "КР":
    # Копирование файлов КР и ИОС с RevitServer на диск Y
    copy_kr_from_localfolder_to_revitserver()
elif selected_options == "ИОС":
    copy_ios_from_localfolder_to_revitserver()
elif selected_options == "ВСЕ РАЗДЕЛЫ":
    copy_ar_from_revitserver_to_localfolder()
    copy_kr_from_localfolder_to_revitserver()
    copy_ios_from_localfolder_to_revitserver()
else:
    pass

# Удаляем обработчик после завершения операции
uiapp.DialogBoxShowing -= dialog_box_handler