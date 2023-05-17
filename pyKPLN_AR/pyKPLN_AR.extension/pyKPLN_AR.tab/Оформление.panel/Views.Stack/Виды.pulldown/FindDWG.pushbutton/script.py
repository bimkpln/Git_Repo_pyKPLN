# -*- coding: utf-8 -*-

from pyrevit import revit, DB
from pyrevit import forms, script

doc = revit.doc

cad_all = DB.FilteredElementCollector(doc).OfClass(DB.ImportInstance).ToElements()

type_names = []

for c in cad_all:
    c_type = doc.GetElement(c.GetTypeId())
    t_name = DB.Element.Name.__get__(c_type)
    type_names.append(t_name)

set_names = list(set(type_names))

cad_ids, cad_types, cad_views = [], [], []

for sn in set_names:
    cad_ids_sub, cad_types_sub, cad_views_sub = [], [], []
    for cad, na in zip(cad_all, type_names):
        if na == sn:
            cad_ids_sub.append(cad.Id)
            if cad.IsLinked:
                cad_types_sub.append("Связь САПР: ")
            else:
                cad_types_sub.append("Импорт САПР: ")
            if cad.ViewSpecific:
                cad_views_sub.append(cad.OwnerViewId)
            else:
                cad_views_sub.append(cad.Document.ActiveView.Id)
                # cad_views_sub.append("")
    cad_ids.append(cad_ids_sub)
    cad_types.append(cad_types_sub)
    cad_views.append(cad_views_sub)

output = script.get_output()

output.print_md("# **Список всех DWG:**")
output.print_md("**Нажми на Id, чтобы выбрать объект, или на значок лупы, чтобы найти в проекте**")

for sn, ids, links, views in zip(set_names, cad_ids, cad_types, cad_views):
    output.print_html('\n<b><span style="background-color:Gainsboro;">Имя файла:</span></b><b><font color=#B22222> {}</b></font>'.format(sn))
    # output.print_html('\n<b><font color=#B22222><span style="background-color:Gainsboro;">{}</span></b></font>'.format(sn))
    # print('\n' + sn + ":")
    for i, l, v in zip(ids, links, views):
        if v == "":
            print(l + 'id {}'.format(output.linkify(i)) + " не удалось определить конкретный вид")
        else:
            print(l + 'id {}'.format(output.linkify(i)) + " размещен на виде: " + 'id {}'.format(output.linkify(v)))
    print('\n')