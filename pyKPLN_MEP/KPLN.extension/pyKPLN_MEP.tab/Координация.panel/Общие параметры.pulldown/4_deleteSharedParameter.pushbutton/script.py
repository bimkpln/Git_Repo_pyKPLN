# coding: utf-8

__title__ = "Удалить ОП в файлах семейства"
__author__ = 'Tima Kutsko'
__doc__ = "Удалить общие параметры (ОП) из файлов семейств"


from Autodesk.Revit.DB import Transaction, BuiltInParameterGroup,\
                              TransactionStatus, StorageType
from Autodesk.Revit.ApplicationServices import Application
from rpw import revit, ui
from pyrevit import script
from System.IO import Directory


# definitions
def main_process(del_param):
    global count
    for current_parameter in family_parameters:
        try:
            param_name = current_parameter.Definition.Name
        except AttributeError:
            continue
        if param_name == del_param:
            try:
                fdoc.Delete(current_parameter.Id)
            except Exception:
                if current_parameter.Definition.Name == param_name:
                    fdoc.FamilyManager.RemoveParameter(current_parameter)
                    break
            count += 1


# input form
components = [ui.forms.flexform.Label("Узел ввода", h_align="Center"),
              ui.forms.flexform.Label("Путь к файлу:"),
              ui.forms.flexform.TextBox("folder", MinWidth=350),
              ui.forms.flexform.
              Label("Имена параметров для удаления, через запятую:"),
              ui.forms.flexform.TextBox("del_param_input", MinWidth=350),
              ui.forms.flexform.Separator(),
              ui.forms.flexform.Button("Запуск")]
form = ui.forms.FlexForm("Params", components)
form.ShowDialog()
folder = form.values["folder"]
del_param_input = form.values["del_param_input"]

# main code
doc = revit.doc
app = doc.Application
count = 0
error_set = set()
output = script.get_output()
# create list of parameters
del_params = [opi for opi in del_param_input.split(',')]
rfas = Directory.GetFiles(folder, "*.rfa")
for rfa in rfas:
    rfa_name = str(rfa).split('/')[-1]
    rfa_copies = rfa_name.split('.')[-2]
    if not rfa_copies.isdigit():
        # variables
        fdoc = app.OpenDocumentFile(rfa)
        fman = fdoc.FamilyManager
        trans = Transaction(fdoc)
        trans.Start("Удалить ОП в семействе {}".format(rfa_name))
        # main part
        family_types = fdoc.FamilyManager.Types
        family_parameters = fdoc.FamilyManager.Parameters
        current_parameter = None
        for del_param in del_params:
            main_process(del_param)
        trans.Commit()
        fdoc.Close(True)


# output results
output.print_md("**Успешно для {} параметра/-ов**".format(str(count)))
if len(error_set) != 0:
    output.print_md("**Ошибки:**")
    for error in error_set:
        print(str(error))
