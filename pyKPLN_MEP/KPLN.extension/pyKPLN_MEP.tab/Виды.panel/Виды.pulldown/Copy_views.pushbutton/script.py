# coding: utf8
"""
"""

__doc__ = "Копирование шаблонов видов из проекта в проект"
__title__ = "Копировать шаблоны проектов"

from datetime import datetime
from System.Collections.Generic import List
from pyrevit import script
from pyrevit import forms
from rpw import doc, uidoc, DB, UI, db, ui, revit
from Autodesk.Revit.DB import FilteredElementCollector, ViewFamilyType, CopyPasteOptions, ElementTransformUtils,\
    ElementId, Transform, Element, View
from System.Security.Principal import WindowsIdentity
from libKPLN.get_info_logger import InfoLogger


#classes
class CheckBoxOption:
    def __init__(self, name, value, default_state=True):
        self.name = name
        self.state = default_state
        self.value = value

    def __nonzero__(self):
        return self.state
    
    def __bool__(self):
        return self.state

opened_docs_dict = {document.Title: document for document in revit.docs}

# form
ComboBox = ui.forms.flexform.ComboBox
Label = ui.forms.flexform.Label
Button = ui.forms.flexform.Button
CheckBox = ui.forms.flexform.CheckBox
# doc selection
components = [Label("Выбери файл из которого будешь копировать:"),
              ComboBox("source", opened_docs_dict),
              CheckBox("templates", 'Шаблоны', default = True),
              CheckBox('schedules', 'Спецификации, чертежные виды', default = False),
              Button("Select")]

form = ui.forms.FlexForm("Select Source Document", components)
form.ShowDialog()

try:
    source_doc = form.values["source"]
    copy_templates_bool = form.values["templates"]
    copy_schedules_bool = form.values["schedules"]

    
    target_doc_options = [CheckBoxOption(doc.Title, doc)
                   for doc in sorted(revit.docs, key=lambda x: x.Title)
                   if doc.Title != source_doc.Title]
    target_doc_checkboxes = forms.SelectFromList.show(target_doc_options,
                                               multiselect = True,
                                               title='Выбери документ, куда копировать:',
                                               width=400,
                                               button_name='Выбрать')
    if target_doc_checkboxes:
        #getting info logger about user
        log_name = "Виды_Копировать шаблоны/спецификации"
        InfoLogger(WindowsIdentity.GetCurrent().Name, log_name)
        #main part of code
        selected_target_doc = [c.value for c in target_doc_checkboxes if c.state == True]

        source_schedules, source_templates = [], []
        if copy_schedules_bool:
            source_schedules = db.Collector(doc=source_doc, 
                                                of_category='OST_Schedules', 
                                                is_type=False).get_elements()
            source_schedules += db.Collector(doc=source_doc,
                                            of_category='OST_Views',
                                            is_type=False,
                                            where=lambda x: not x.IsTemplate 
                                            and x.ViewType == DB.ViewType.DraftingView).get_elements()
            
            source_schedules = sorted(source_schedules, key=lambda x: x.Name) 


            
        if copy_templates_bool:
            source_templates = db.Collector(doc=source_doc,
                                            of_category='OST_Views',
                                            is_type=False,
                                            where=lambda x: x.IsTemplate).get_elements()
            source_templates = sorted(source_templates, key=lambda x: x.Name)

        source_views = source_templates + source_schedules
        source_views_options = [CheckBoxOption(v.Name, v) for v in source_views]

        source_views_checkboxes = forms.SelectFromList.show(source_views_options,
                                                   multiselect = True,
                                                   title='Выбери элементы для копирования:',
                                                   width=400,
                                                   button_name='Выбрать')
        if source_views_checkboxes:
            selected_source_views = [c.value for c in source_views_checkboxes if c.state == True]
            source_view_ids = List[ElementId]([v.Id for v in selected_source_views])

            # Copy view templates
            copypasteoptions = CopyPasteOptions()
            for target_doc in selected_target_doc:
                # dupplicate checking:
                # target templates and schedules (if selected in CheckBox)
                target_schedules, target_templates = [], []
                if copy_schedules_bool:
                    target_schedules = db.Collector(doc=target_doc,
                                                    of_class='ViewSchedule',
                                                    is_type=False).elements
                if copy_templates_bool:
                    target_templates = db.Collector(doc=target_doc,
                                                    of_category='OST_Views',
                                                    is_type=False,
                                                    where=lambda x: x.IsTemplate).get_elements()

                target_views = target_schedules + target_templates
                source_view_names = [v.Name for v in selected_source_views]

                with db.Transaction(doc=target_doc,
                                    name="pyKPLN_Копировать виды из: {}".format(source_doc.Title)):

                    for v in target_views:
                        # rename old views
                        if v.Name in source_view_names:
                            v.Name = "{}_old_{}".format(v.Name,
                                                      datetime.now().strftime("%y%m%d_%H-%M-%S"))

                    ElementTransformUtils.CopyElements(source_doc,
                                                       source_view_ids,
                                                       target_doc,
                                                       Transform.Identity,
                                                       copypasteoptions)

except KeyError:
    script.exit()



