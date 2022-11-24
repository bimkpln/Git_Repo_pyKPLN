# -*- coding: utf-8 -*-
"""
Copy_Template

"""
__author__ = 'Igor Perfilyev - envato.perfilev@gmail.com'
__title__ = "Копировать\nтипы стен"
__doc__ = 'Копирование типоразмеров стен' \


"""
Архитекурное бюро KPLN

"""
import webbrowser
import math
from pyrevit.framework import clr
import re
from rpw import doc, uidoc, DB, UI, db, ui, revit
from pyrevit import script
from pyrevit import forms
from pyrevit import DB, UI
from pyrevit.revit import Transaction, selection
from System.Collections.Generic import *
import System
from System.Windows.Forms import *
from System.Drawing import *
import datetime
from rpw.ui.forms import Alert, CommandLink, TaskDialog

class CopyHandler(DB.IDuplicateTypeNamesHandler):
	def OnDuplicateTypeNamesFound(self, args):
		return DB.DuplicateTypeAction.UseDestinationTypes
class CreateWindow(Form):
	def __init__(self): 
		self.Name = "Копировать шаблоны"
		self.Text = "Копирование шаблонов"
		self.Size = Size(418, 608)
		self.Icon = Icon("X:\\BIM\\5_Scripts\\Git_Repo_pyKPLN\\pyKPLN_AR\\pyKPLN_AR.extension\\pyKPLN_AR.tab\\icon.ico")
		self.button_create = Button(Text = "Ок")
		self.ControlBox = True
		self.TopMost = True
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.MinimumSize = Size(418, 480)
		self.MaximumSize = Size(418, 480)
		self.FormBorderStyle = FormBorderStyle.FixedDialog
		self.CenterToScreen()

		self.lb1 = Label(Text = "Выберите проект с исходными типоразмерами:")
		self.lb1.Parent = self
		self.lb1.Size = Size(343, 15)
		self.lb1.Location = Point(30, 80)

		self.lb2 = Label(Text = "Типоразмеры стен из выбранного проекта будут\nскопированы в активный проект с заменой по наименованию.")
		self.lb2.Parent = self
		self.lb2.Size = Size(343, 40)
		self.lb2.Location = Point(30, 125)

		self.btn_confirm = Button(Text = "Далее")
		self.btn_confirm.Parent = self
		self.btn_confirm.Location = Point(30, 380)
		self.btn_confirm.Enabled = False
		self.btn_confirm.Click += self.OnNext

		self.active_doc = doc
		self.other_docs = self.GetDocs()
		self.target_doc = None

		self.cb = ComboBox()
		self.cb.Parent = self
		self.cb.Size = Size(343, 100)
		self.cb.Location = Point(30, 100)
		self.cb.DropDownStyle = ComboBoxStyle.DropDownList
		self.cb.SelectedIndexChanged += self.allow_next
		self.cb.Items.Add("Выберите файл")
		self.cb.Text = "Выберите файл"
		for docel in self.other_docs:
			self.cb.Items.Add(docel.Title)

		self.listbox = ListView()
		self.c_cb = ColumnHeader()
		self.c_cb.Text = ""
		self.c_cb.Width = -2
		self.c_cb.TextAlign = HorizontalAlignment.Center
		self.c_name = ColumnHeader()
		self.c_name.Text = "Имя"
		self.c_name.Width = -2
		self.c_name.TextAlign = HorizontalAlignment.Left

		self.SuspendLayout()
		self.listbox.Dock = DockStyle.Fill
		self.listbox.View = View.Details

		self.listbox.Parent = self
		self.listbox.Size = Size(400, 400)
		self.listbox.Location = Point(1, 1)
		self.listbox.FullRowSelect = True
		self.listbox.GridLines = True
		self.listbox.AllowColumnReorder = True
		self.listbox.Sorting = SortOrder.Ascending
		self.listbox.Columns.Add(self.c_cb)
		self.listbox.Columns.Add(self.c_name)
		self.listbox.LabelEdit = True
		self.listbox.CheckBoxes = True
		self.listbox.MultiSelect = True

		self.button_ok = Button(Text = "Ok")
		self.button_ok.Parent = self
		self.button_ok.Location = Point(10, 410)
		self.button_ok.Click += self.OnOk

		self.button_cancel = Button(Text = "Закрыть")
		self.button_cancel.Parent = self
		self.button_cancel.Location = Point(100, 410)
		self.button_cancel.Click += self.OnCancel

		self.types = []
		self.project_types = self.get_types(doc)

		self.listbox.Hide()
		self.button_ok.Hide()
		self.button_cancel.Hide()

	def get_checked_types(self):
		list = []
		for i in self.item:
			if i.Checked:
				si = i.SubItems
				type_name = si[1].Text
				for t in self.types:
					if type_name == t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString():
						list.append(t)
		return list

	def OnOk(self, sender, args):
		self.Hide()
		t = []
		t = self.get_checked_types()
		if len(t) != 0:
			self.Copy(self.active_doc, self.target_doc, t)
		self.Close()

	def OnNext(self, sender, args):
		self.cb.Hide()
		self.lb1.Hide()
		self.lb2.Hide()
		self.btn_confirm.Hide()
		self.listbox.Show()
		self.button_ok.Show()
		self.button_cancel.Show()
		for i in self.GetDocs():
			if i.Title == self.cb.Text:
				self.target_doc = i
		if self.target_doc != None:
			self.Text = "Выбор копируемых типоразмеров"
			self.update(self.active_doc, self.target_doc)
		else:
			self.Hide()
			forms.alert("Не найден документ!")
			self.Show()
			self.Close()

	def allow_next(self, sender, event):
		if self.cb.Text != "Выберите файл":
			self.btn_confirm.Enabled = True
		else:
			self.btn_confirm.Enabled = False

	def Exsist(self, type):
		try:
			for project_type in self.project_types:
				if type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == project_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString():
					return project_type
			return False
		except Exception as e:
			forms.alert("Exsist : {}".format(str(e)))
			return False

	def update(self, active_doc, target_doc):
		self.types = self.get_types(target_doc)
		self.listbox.Clear()
		self.listbox.Parent = self
		self.listbox.Size = Size(400, 400)
		self.listbox.Location = Point(1, 1)
		self.listbox.FullRowSelect = True
		self.listbox.GridLines = True
		self.listbox.AllowColumnReorder = True
		self.listbox.Sorting = SortOrder.Ascending
		self.listbox.Columns.Add(self.c_cb)
		self.listbox.Columns.Add(self.c_name)
		self.listbox.LabelEdit = True
		self.listbox.CheckBoxes = True
		self.listbox.MultiSelect = True
		self.item = []
		self.t = []
		for type in self.types:
			self.t.append(type)
			self.item.append(ListViewItem())
			self.item[len(self.item)-1].Text = ""
			self.item[len(self.item)-1].Checked = False
			self.item[len(self.item)-1].SubItems.Add(str(type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()))
			self.listbox.Items.Add(self.item[len(self.item)-1])	

	def get_types(self, document):
		list = []
		for i in DB.FilteredElementCollector(document).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsElementType().ToElements():
			try:
				if i.Kind == DB.WallKind.Basic:
					list.append(i)
			except:
				pass
		return list

	def OnCancel(self, sender, args):
		self.Close()

	def CopyMaterials(self, type, target_doc):
		try:
			options = DB.CopyPasteOptions()
			options.SetDuplicateTypeNamesHandler(CopyHandler())
			layers = type.GetCompoundStructure().GetLayers()
			for layer in layers:
				try:
					collection = List[DB.ElementId]()	
					collection.Add(layer.MaterialId)
					DB.ElementTransformUtils.CopyElements(target_doc, collection, doc, DB.Transform.Identity, options)
				except: pass
		except Exception as e:
			pass

	def GetDocs(self):
		try:
			list = []
			for o_doc in revit.docs:
				if self.active_doc.PathName != o_doc.PathName:
					list.append(o_doc)
			return list
		except:
			pass

	def GetAnalogMaterial(self, material):
		try:
			name = material.get_Parameter(DB.BuiltInParameter.MATERIAL_NAME).AsString()
			for i in DB.FilteredElementCollector(doc).OfClass(DB.Material):
				if i.get_Parameter(DB.BuiltInParameter.MATERIAL_NAME).AsString() == name:
					return i.Id
		except:
			pass

	def CopyParameters(self, type, fromtype):
		for j_1 in type.Parameters:
			for j_2 in fromtype.Parameters:
				try:
					if not j_1.IsReadOnly:
						if j_1.Definition.Name == j_2.Definition.Name and j_1.StorageType == j_2.StorageType:
							if j_1.StorageType == DB.StorageType.Integer:
								type.LookupParameter(j_1.Definition.Name).Set(fromtype.LookupParameter(j_1.Definition.Name).AsInteger())
							elif j_1.StorageType == DB.StorageType.Double:
								type.LookupParameter(j_1.Definition.Name).Set(fromtype.LookupParameter(j_1.Definition.Name).AsDouble())
							elif j_1.StorageType == DB.StorageType.String:
								type.LookupParameter(j_1.Definition.Name).Set(fromtype.LookupParameter(j_1.Definition.Name).AsString())
							elif j_1.StorageType == DB.StorageType.ElementId:
								type.LookupParameter(j_1.Definition.Name).Set(fromtype.LookupParameter(j_1.Definition.Name).AsElementId())
				except:
					pass


	def Copy(self, active_doc, target_doc, elements):
		for element in elements:
			collection = List[DB.ElementId]()	
			collection.Add(element.Id)
			options = DB.CopyPasteOptions()
			options.SetDuplicateTypeNamesHandler(CopyHandler())
			with db.Transaction(name = "copy tempates"):
				self.Hide()
				if self.Exsist(element) == False:
					try:
						DB.ElementTransformUtils.CopyElements(target_doc, collection, active_doc, DB.Transform.Identity, options)
						forms.alert("Тип «{}» успешно скопирован в проект!".format(str(element.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString())))
					except Exception as e:
						forms.alert("Тип «{}» невозможно копировать в проект! Операция отменена! {}".format(str(element.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()), str(e)))
				else:
					try:
						t = self.Exsist(element)
						self.CopyMaterials(element, target_doc)
						doc.Regenerate()
						layers_collection = List[DB.CompoundStructureLayer]()
						cs = element.GetCompoundStructure()
						pcs = t.GetCompoundStructure()
						for layer in cs.GetLayers():
							try:
								sl = DB.CompoundStructureLayer(layer.Width, layer.Function, self.GetAnalogMaterial(target_doc.GetElement(layer.MaterialId)))
							except:
								sl = DB.CompoundStructureLayer()
								sl.Width = layer.Width
								sl.Function = layer.Function
								try:
									sl.LayerCapFlag = layer.LayerCapFlag
								except: pass
								try:
									collection_deck = List[DB.ElementId]()	
									collection_deck.Add(layer.DeckProfileId)
									DB.ElementTransformUtils.CopyElements(target_doc, collection_deck, active_doc, DB.Transform.Identity, options)
									doc.Refresh()
									sl.DeckProfileId = layer.DeckProfileId
								except: pass
								try:
									sl.DeckEmbeddingType = layer.DeckEmbeddingType
								except: pass
							layers_collection.Add(sl)
						pcs.SetLayers(layers_collection)
						try:
							pcs.CutoffHeight = cs.CutoffHeight
						except:
							pass
						try:
							pcs.OpeningWrapping = cs.OpeningWrapping
						except:
							pass
						try:
							pcs.SampleHeight = cs.SampleHeight
						except:
							pass
						try:
							pcs.StructuralMaterialIndex = cs.StructuralMaterialIndex
						except:
							pass
						try:
							pcs.VariableLayerIndex = cs.VariableLayerIndex
						except:
							pass
						t.SetCompoundStructure(pcs)
						self.CopyParameters(t, element)
						forms.alert("Тип «{}» успешно заменен!".format(str(element.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString())))
					except Exception as e:
						forms.alert("Copy : {}".format(str(e)))
						forms.alert("Тип «{}» невозможно копировать в проект! Операция отменена! {}".format(str(element.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()), str(e)))



form = CreateWindow()
Application.Run(form)

