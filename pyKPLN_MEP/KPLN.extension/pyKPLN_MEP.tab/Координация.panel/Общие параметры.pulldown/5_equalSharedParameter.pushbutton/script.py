# coding: utf-8

__title__ = "Прировнять ОП в файлах семейства"
__author__ = 'Tima Kutsko'
__doc__ = '''Приравнивание значений общих параметров (ОП)
             из ФОП в файлах семейств'''


from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.ApplicationServices import Application
from rpw import revit, ui
from pyrevit import script
from System.IO import Directory


# input form
components = [ui.forms.flexform.Label("Узел ввода", h_align="Center"),
              ui.forms.flexform.Label("Путь к файлу:"),
              ui.forms.flexform.TextBox("folder", MinWidth=350),
              ui.forms.flexform.
              Label("Имена параметров, которые будут в формуле"),
              ui.forms.flexform.TextBox("formula_param_input", MinWidth=350),
              ui.forms.flexform.
              Label("Имена параметров, которые следует прировнять:"),
              ui.forms.flexform.TextBox("equal_param_input", MinWidth=350),
              ui.forms.flexform.Separator(),
              ui.forms.flexform.Button("Запуск", h_align="Center")]
form = ui.forms.FlexForm("Params", components)
form.ShowDialog()
folder = form.values["folder"]
formula_param_input = form.values["formula_param_input"]
equal_param_input = form.values["equal_param_input"]

# main code
doc = revit.doc
app = doc.Application
count = 0
error_set = set()
output = script.get_output()
# create list of parameters
formula_params = [fpi for fpi in formula_param_input.split(',')]
equal_params = [epi for epi in equal_param_input.split(',')]
rfas = Directory.GetFiles(folder, "*.rfa")
for rfa in rfas:
    rfa_name = str(rfa).split('/')[-1]
    rfa_copies = rfa_name.split('.')[-2]
    if not rfa_copies.isdigit():
        # variables
        fdoc = app.OpenDocumentFile(rfa)
        fman = fdoc.FamilyManager
        # main part
        family_types = fdoc.FamilyManager.Types
        family_parameters = fdoc.FamilyManager.Parameters
        formula_param = None
        equal_param = None
        trans = Transaction(fdoc)
        trans.Start("Прированять ОП в семействе {}".format(rfa_name))
        for formula_param_text, equal_param_text in \
                zip(formula_params, equal_params):
            for current_parameter in family_parameters:
                try:
                    param_name = current_parameter.Definition.Name
                except AttributeError:
                    continue
                if param_name == formula_param_text:
                    formula_param = current_parameter
                if param_name == equal_param_text:
                    equal_param = current_parameter
            if formula_param and equal_param:
                try:
                    fman.SetFormula(equal_param,
                                    "[" +
                                    str(formula_param.Definition.Name) +
                                    "]")
                except Exception as exc:
                    print(str(exc) + "for family {}".format(equal_param.
                                                            Definition.
                                                            Name))
                count += 1
        trans.Commit()
        fdoc.Close(True)


# output results
output.print_md("**Успешно для {} параметра/-ов**".format(str(count)))
if len(error_set) != 0:
    output.print_md("**Ошибки:**")
    for error in error_set:
        print(str(error))