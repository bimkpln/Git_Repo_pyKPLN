# -*- coding: utf-8 -*-

__author__ = 'Gyulnara Fedoseeva'
__title__ = "test"
__doc__ = 'Резервное копирование моделей с RevitServer на диск Y, обновление исходных данных по разделам КР, ИОС'\

"""
Архитекурное бюро KPLN

"""
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.Exceptions import InvalidOperationException
from Autodesk.Revit.UI import TaskDialog, TaskDialogResult

#region Классы
class ErrorBox:
    """Класс-контейнер для хранения значений ошибок при открытии файла и инструкций к тому, как эти ошибки закрыть"""

    def __init__(self, name, overrideResult):
        self.name = name
        """Имя окна, но в Ревит это DialogId"""
        
        self.overrideResult = overrideResult
        """Кнопку, которую нужно нажать (TaskDialogResult)"""
#endregion

#region Функции
def dialog_box_handler(sender, e):
    global errorBoxes
    currentUIApp = sender
    dialogBoxShowingEventArgs = e
    dialogStrId = dialogBoxShowingEventArgs.DialogId
    # dialogMessage = dialogBoxShowingEventArgs.Message()
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



def copy_ar_from_revitserver_to_localfolder():
    # Список файлов АР в исходной папке
    ar_files = ["RSN://192.168.0.5/СТАК/Стадия_Р/АР/К2.1/СТАК_К2.1_РД_АР_АР1.rvt"]

    try:
        for f in ar_files:
            ar_file_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(f)

            ar_new_file_path = "Y:\\Жилые здания\\Старый_Акульшет\\10.Стадия_Р\\5.АР\\1.RVT\\RevitServer_Архив\\Новая папка" + "\\" + f.split('/')[-1]
            ar_new_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(ar_new_file_path)

            copy_files(ar_file_path, ar_new_path)
    except InvalidOperationException as ex:
        TaskDialog.Show('Ошибка', '{}'.format(str(ex)))


def copy_files(source_path, new_path):
    # Открываем файл Revit
    options = OpenOptions()
    options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
    options.SetOpenWorksetsConfiguration(WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets))
    try:
        opened_doc = app.OpenDocumentFile(source_path, options)

        # Сохраненяем файлы
        save_options = SaveAsOptions()
        save_options.OverwriteExistingFile = True

        ws_options = WorksharingSaveAsOptions()
        ws_options.SaveAsCentral = True

        save_options.SetWorksharingOptions(ws_options)
        # opened_doc.SaveAs(new_path, save_options)
    except Exception as e:
        print(e)
#endregion

uiapp = __revit__
doc = uiapp.ActiveUIDocument.Document
app = doc.Application
# Если что - есть большой список ошибок и их обработчиков в БД.
errorBoxes = [
    ErrorBox('', TaskDialogResult.Close),
    ErrorBox('TaskDialog_Audit_Warning', TaskDialogResult.Yes)]

# Подписываемся на событие всплывающего окна
uiapp.DialogBoxShowing += dialog_box_handler

copy_ar_from_revitserver_to_localfolder()

# Удаляем обработчик после завершения операции
uiapp.DialogBoxShowing -= dialog_box_handler

kor_list = ["1.1", "1.2", "2.1", "2.2", "3.1", "3.2", "4.1", "4.2"]
ios_list = [["7.1.ЭОМ", "ЭОМ"], ["7.2.ВК", "ВК"], ["7.4.ОВ", "ОВиК"], ["7.5.СС", "СС"]]