# coding: utf-8

__title__ = "Заменить/добавить ОП в файлах семейства"
__author__ = 'Tima Kutsko'
__doc__ = "Замена/добавление общих параметров (ОП) из ФОП в файлах семейств"


from Autodesk.Revit.DB import Transaction, BuiltInParameterGroup,\
                              TransactionStatus, StorageType
from Autodesk.Revit.ApplicationServices import Application
from rpw import revit, ui
from pyrevit import script
from System.IO import Directory


# definitions
def GetsharedParam(paramName):
    deffile = app.OpenSharedParameterFile()
    for defgr in deffile.Groups:
        for exDef in defgr.Definitions:
            if exDef.Name == paramName:
                return exDef
    ui.forms.Alert("",
                   header="Не найден параметр " + paramName + ". Отменено!",
                   exit=True)
    return None


def creation_param(exDef, added=True, data_list=list()):
    global count
    global error_double_set
    global error_set
    try:
        added_parameter = fman.AddParameter(exDef, group, isInstance)
        if added:
            n = 0
            for current_type in family_types:
                fdoc.FamilyManager.CurrentType = current_type
                fdoc.FamilyManager.Set(added_parameter, data_list[n])
                n += 1
        count += 1
    except Exception as exc:
        if "is already in use" in str(exc):
            error_double_set.add(fdoc.Title)
            #print("Параметр в семействе {} уже есть!".format(fdoc.Title))
        else:
            error_set.add(fdoc.Title)
            #print("Не удалось добавить в {}".format(fdoc.Title))
        #trans.Commit()


def trans_close(tr):
    if tr.GetStatus() != TransactionStatus.Committed:
        tr.Commit()
        fdoc.Close(True)


def main_process(old_param):
    global flag
    global n_old_param
    for current_parameter in family_parameters:
        if flag:
            try:
                param_name = current_parameter.Definition.Name
            except AttributeError:
                continue
            if param_name == old_param:
                types_data_list = list()
                old_family_param = current_parameter
                for current_type in family_types:
                    if old_family_param.StorageType == StorageType.String:
                        old_parameter_data = current_type.\
                                            AsString(old_family_param)
                    elif old_family_param.StorageType == StorageType.Double:
                        old_parameter_data = current_type.\
                                            AsDouble(old_family_param)
                    elif old_family_param.StorageType == StorageType.Integer:
                        old_parameter_data = current_type.\
                                            AsInteger(old_family_param)
                    elif old_family_param.StorageType == StorageType.ElementId:
                        old_parameter_data = current_type.\
                                            AsElementId(old_family_param)
                    types_data_list.append(old_parameter_data)
                # deleting old parameter
                try:
                    fdoc.Delete(old_family_param.Id)
                except Exception:
                    for current_parameter in family_parameters:
                        try:
                            param_name = current_parameter.Definition.Name
                        except Exception:
                            continue
                        if param_name == old_family_param.Definition.Name:
                            fdoc.\
                                FamilyManager.\
                                RemoveParameter(current_parameter)
                            break
                try:
                    new_exDef = GetsharedParam(new_params[n_old_param])
                    creation_param(new_exDef, True, types_data_list)
                    #trans_close(trans)
                except Exception as exc:
                    print("Ошибка {} в семействе {}".format(exc, rfa_name))
            elif old_param in family_param_names:
                continue
            else:
                new_exDef = GetsharedParam(new_params[n_old_param])
                creation_param(new_exDef, False)
                flag = False


# input form
components = [ui.forms.flexform.Label("Узел ввода", h_align="Center"),
              ui.forms.flexform.Label("Путь к файлу:"),
              ui.forms.flexform.TextBox("folder", MinWidth=350),
              ui.forms.flexform.
              Label("Имена параметров для замены, через запятую:"),
              ui.forms.flexform.TextBox("old_param_input", MinWidth=350),
              ui.forms.flexform.
              Label("Имена новых параметров на замену, через запятую:"),
              ui.forms.flexform.TextBox("new_param_input", MinWidth=350),
              ui.forms.flexform.CheckBox('isInstance',
                                         'Экземпляр?',
                                         default=True),
              ui.forms.flexform.Separator(),
              ui.forms.flexform.Button("Запуск")]
form = ui.forms.FlexForm("Params", components)
form.ShowDialog()
folder = form.values["folder"]
old_param_input = form.values["old_param_input"]
new_param_input = form.values["new_param_input"]
isInstance = form.values["isInstance"]

# main code
doc = revit.doc
app = doc.Application
count = 0
error_double_set = set()
error_set = set()
output = script.get_output()
# create list of parameters
old_params = [opi for opi in old_param_input.split(',')]
new_params = [npi for npi in new_param_input.split(',')]
#new_exDef = GetsharedParam(new_param)
group = BuiltInParameterGroup.PG_TEXT
rfas = Directory.GetFiles(folder, "*.rfa")
for rfa in rfas:
    rfa_name = str(rfa).split('/')[-1]
    rfa_copies = rfa_name.split('.')[-2]
    if not rfa_copies.isdigit():
        # variables
        fdoc = app.OpenDocumentFile(rfa)
        fman = fdoc.FamilyManager
        trans = Transaction(fdoc)
        trans.Start("Заменить ОП в семействе {}".format(rfa_name))
        # main part
        family_types = fdoc.FamilyManager.Types
        family_parameters = fdoc.FamilyManager.Parameters
        old_family_param = None
        family_param_names = list()
        for current_parameter in family_parameters:
            try:
                param_name = current_parameter.Definition.Name
                family_param_names.append(param_name)
            except AttributeError:
                continue
        n_old_param = -1
        for old_param in old_params:
            flag = True
            n_old_param += 1
            main_process(old_param)
        trans.Commit()
        fdoc.Close(True)


# output results
output.print_md("**Успешно для {} параметра/-ов**".format(str(count)))
if len(error_double_set) != 0:
    output.print_md("**Параметры уже были:**")
    for error in error_double_set:
        print(str(error))
if len(error_set) != 0:
    output.print_md("**Ошибки:**")
    for error in error_set:
        print(str(error))
