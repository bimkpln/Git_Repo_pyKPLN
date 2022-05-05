# coding: utf-8

__title__ = "Удаление ОП в проекте/семействе"
__author__ = 'Tima Kutsko'
__doc__ = '''Поиск и удаление общих параметров (ОП) используемых
             в проекте/семействе'''


from Autodesk.Revit.DB import FilteredElementCollector, SharedParameterElement
from rpw import revit, ui, db
from pyrevit import script, forms
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger
import clr
clr.AddReference('RevitAPI')


# variables
doc = revit.doc
output = script.get_output()


# classes
class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value


# definitions
def create_check_boxes_by_name(elements):
    elements_options = [CheckBoxOption(e.Name, e)
                        for e in sorted(elements, key=lambda x: x.Name)]
    elements_checkboxes = forms.SelectFromList.show(elements_options,
                                                    multiselect=True,
                                                    title='''Выбери ОП,
                                                              чтобы удалить''',
                                                    width=500,
                                                    button_name='УДАЛИТЬ!')
    return elements_checkboxes


# main code
ui.forms.Alert('При удалении ОП - удалиться вся информация в нем без'
               ' предупреждения!',
               title='pyKPLN_Общие параметры',
               header='ВНИМАНИЕ!!!')


shared_parameter_list = sorted(FilteredElementCollector(doc).
                               OfClass(SharedParameterElement),
                               key=lambda n: n.Name)
parameter_checkbox = create_check_boxes_by_name(shared_parameter_list)
try:
    selected_parameters = [p.value
                           for p in parameter_checkbox if p.state is True]
except Exception:
    script.exit()


with db.Transaction('pyKPLN_Удалить общие параметры'):
    # getting info logger about user
    log_name = "Пространства и помещения_" + str(__title__)
    InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
    # main code
    isFamManager = False
    for parameter in selected_parameters:
        if not isFamManager:
            try:
                doc.Delete(parameter.Id)
            except Exception as exc:
                if "ElementId cannot be deleted" in str(exc):
                    isFamManager = True
                else:
                    print(exc)
                    break
        if isFamManager:
            family_parameters = doc.FamilyManager.Parameters
            for param in family_parameters:
                fManParamName = param.Definition.Name
                if fManParamName == parameter.Name:
                    output.print_md(
                        "Параметр с id {} удален,"
                        "**но он остался в файле!**".
                        format(parameter.Id))
                    doc.FamilyManager.RemoveParameter(param)
                    break

output.print_md('Удалено **{}** параметров'.format(len(selected_parameters)))
