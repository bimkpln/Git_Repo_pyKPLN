# -*- coding: utf-8 -*-

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import *
from System.Drawing import *
from System import EventHandler
import webbrowser as wb

from pyrevit import script
from Autodesk.Revit import UI

SCRIPT_PATH = script.get_script_path()
UIAPP = __revit__
UIDOC = UIAPP.ActiveUIDocument
DOC = UIDOC.Document
ACTIVE_VIEW = DOC.ActiveView

for uiview in UIDOC.GetOpenUIViews():
    if uiview.ViewId == ACTIVE_VIEW.Id:
        uiview_ = uiview
        break
screen_size = uiview_.GetWindowRectangle()
SCREEN_SIZE_X = int(screen_size.Right - ((screen_size.Right-screen_size.Left)/2))
SCREEN_SIZE_Y = int(screen_size.Bottom - ((screen_size.Bottom-screen_size.Top)/2))

GRAPHICS_PATH = Drawing2D.GraphicsPath()

class KPLN_Button(Button):
    def __init__(self, form, image_off, image_on, action):
        '''
        '''
        self.form = form # Родительский класс формы
        self.image_off = image_off # Изображение кнопки без мыши
        self.image_on = image_on # Изображение кнопки с мышкой
        self.action = action # Ссылка

        #region ВИЗУАЛ КНОПКИ
        self.FlatAppearance.BorderSize = 0
        self.FlatStyle = FlatStyle.Flat
        self.Cursor = Cursors.Hand
        self.BackgroundImage = Image.FromFile(''.join([SCRIPT_PATH, "\\", self.image_off]))
        #endregion

        #region СОБЫТИЯ КНОПКИ
        self.Click += self.mouse_click
        self.MouseEnter += EventHandler(self.mouse_enter)
        self.MouseLeave += EventHandler(self.mouse_leave)
        #endregion

        self.form.Controls.Add(self)

    def mouse_click(self, sender, args):
        exec(self.action)

    def mouse_enter(self, sender, args):
        self.BackgroundImage = Image.FromFile(''.join([SCRIPT_PATH, "\\", self.image_on]))

    def mouse_leave(self, sender, args):
        self.BackgroundImage = Image.FromFile(''.join([SCRIPT_PATH, "\\", self.image_off]))

class Task_Dialog_KPLN_Form(Form):
    def __init__(self, text, site = None):
        '''
        '''
        self.text = text
        self.site = site
        
        #region ВИЗУАЛ ФОРМЫ
        self.StartPosition = FormStartPosition.Manual
        self.Location = Point(
            SCREEN_SIZE_X - 250,
            SCREEN_SIZE_Y - 150
        )
        self.Size = Size(500, 300)
        self.FormBorderStyle = FormBorderStyle.None
        self.BackgroundImage = Image.FromFile(''.join([SCRIPT_PATH, "\\", r"Images\Background.png"]))
        self.Opacity = 0.91
        #endregion

        self.InitComp()

    def InitComp(self):
        label = Label()
        label.Location = Point(36, 37)
        label.Size = Size(427, 164)
        label.Text = self.text
        label.TextAlign = ContentAlignment.MiddleCenter
        label.Font = Font("Century Gothic", 16, FontStyle.Regular)
        self.Controls.Add(label)

        if self.site != None:
            button_ok = KPLN_Button(self, r"Images\button_OK.png", r"Images\button_OK_White.png", "self.form.Close()")
            button_ok.Location = Point(150, 215)
            button_ok.Size = Size(125, 56)

            button_q = KPLN_Button(self, r"Images\button_Q.png", r"Images\button_Q_White.png", "wb.open_new(" + self.site + ")")
            button_q.Location = Point(295, 215)
            button_q.Size = Size(56, 56)
        else:
            button_ok = KPLN_Button(self, r"Images\button_OK.png", r"Images\button_OK_White.png", "self.form.Close()")
            button_ok.Location = Point(188, 215)
            button_ok.Size = Size(125, 56)

    def Run(self):
        Application.Run(self)

class Set_Of_Tools_KPLN_Form(Form):
    def __init__(self):
        self.StartPosition = FormStartPosition.Manual
        self.Location = Point(
            SCREEN_SIZE_X - 300,
            SCREEN_SIZE_Y - 300
        )
        self.Size = Size(600, 600)

        # Реализовать удержание клавиши после того, как будет нормально прописана функция Run
        self.Load += EventHandler(self.form_load)
        self.KeyPreview = True # Обязательно 
        self.KeyPress += KeyPressEventHandler(self.key_press)

        self.FormBorderStyle = FormBorderStyle.None
        self.AllowTransparency = True  
        self.BackColor = Color.AliceBlue 
        self.TransparencyKey = self.BackColor
        self.TopMost = True
        
        self.InitComp()

    def InitComp(self):
        # 0
        btn_0_command = "UI.RevitCommandId.LookupPostableCommandId(UI.PostableCommand.Duct)"
        btn_0 = KPLN_Button(self, r"Images\buttons_off.png", r"Images\buttons_on.png", "UIAPP.PostCommand(" + btn_0_command + ")\nself.form.Close()")
        btn_0.Location = Point(0, 0)
        btn_0.Size = Size(600, 600)

        GRAPHICS_PATH.Reset()
        GRAPHICS_PATH.AddPie(65, 65, 470, 470, 0, -45)
        btn_0.Region = Region(GRAPHICS_PATH)

        # 45
        btn_45_command = "UI.RevitCommandId.LookupPostableCommandId(UI.PostableCommand.AlignedDimension)"
        btn_45 = KPLN_Button(self, r"Images\buttons_off.png", r"Images\buttons_on.png", "UIAPP.PostCommand(" + btn_45_command + ")\nself.form.Close()")
        btn_45.Location = Point(0, 0)
        btn_45.Size = Size(600, 600)

        GRAPHICS_PATH.Reset()
        GRAPHICS_PATH.AddPie(65, 65, 470, 470, -45, -45)
        btn_45.Region = Region(GRAPHICS_PATH)

        # 90
        btn_90_command = "UI.RevitCommandId.LookupPostableCommandId(UI.PostableCommand.Align)"
        btn_90 = KPLN_Button(self, r"Images\buttons_off.png", r"Images\buttons_on.png", "UIAPP.PostCommand(" + btn_90_command + ")\nself.form.Close()")
        btn_90.Location = Point(0, 0)
        btn_90.Size = Size(600, 600)

        GRAPHICS_PATH.Reset()
        GRAPHICS_PATH.AddPie(65, 65, 470, 470, -90, -45)
        btn_90.Region = Region(GRAPHICS_PATH)

        # 135
        btn_135_command = "UI.RevitCommandId.LookupPostableCommandId(UI.PostableCommand.Move)"
        btn_135 = KPLN_Button(self, r"Images\buttons_off.png", r"Images\buttons_on.png", "UIAPP.PostCommand(" + btn_135_command + ")\nself.form.Close()")
        btn_135.Location = Point(0, 0)
        btn_135.Size = Size(600, 600)

        GRAPHICS_PATH.Reset()
        GRAPHICS_PATH.AddPie(65, 65, 470, 470, -135, -45)
        btn_135.Region = Region(GRAPHICS_PATH)

        # 180
        btn_180_command = "UI.RevitCommandId.LookupPostableCommandId(UI.PostableCommand.Section)"
        btn_180 = KPLN_Button(self, r"Images\buttons_off.png", r"Images\buttons_on.png", "UIAPP.PostCommand(" + btn_180_command + ")\nself.form.Close()")
        btn_180.Location = Point(0, 0)
        btn_180.Size = Size(600, 600)

        GRAPHICS_PATH.Reset()
        GRAPHICS_PATH.AddPie(65, 65, 470, 470, -180, -45)
        btn_180.Region = Region(GRAPHICS_PATH)

        # 225
        btn_225_command = "UI.RevitCommandId.LookupPostableCommandId(UI.PostableCommand.CreateSimilar)"
        btn_225 = KPLN_Button(self, r"Images\buttons_off.png", r"Images\buttons_on.png", "UIAPP.PostCommand(" + btn_225_command + ")\nself.form.Close()")
        btn_225.Location = Point(0, 0)
        btn_225.Size = Size(600, 600)

        GRAPHICS_PATH.Reset()
        GRAPHICS_PATH.AddPie(65, 65, 470, 470, -225, -45)
        btn_225.Region = Region(GRAPHICS_PATH)

        # 270
        btn_270_command = "UI.RevitCommandId.LookupPostableCommandId(UI.PostableCommand.CableTray)"
        btn_270 = KPLN_Button(self, r"Images\buttons_off.png", r"Images\buttons_on.png", "UIAPP.PostCommand(" + btn_270_command + ")\nself.form.Close()")
        btn_270.Location = Point(0, 0)
        btn_270.Size = Size(600, 600)

        GRAPHICS_PATH.Reset()
        GRAPHICS_PATH.AddPie(65, 65, 470, 470, -270, -45)
        btn_270.Region = Region(GRAPHICS_PATH)

        # 315
        btn_315_command = "UI.RevitCommandId.LookupPostableCommandId(UI.PostableCommand.Pipe)"
        btn_315 = KPLN_Button(self, r"Images\buttons_off.png", r"Images\buttons_on.png", "UIAPP.PostCommand(" + btn_315_command + ")\nself.form.Close()")
        btn_315.Location = Point(0, 0)
        btn_315.Size = Size(600, 600)

        GRAPHICS_PATH.Reset()
        GRAPHICS_PATH.AddPie(65, 65, 470, 470, -315, -45)
        btn_315.Region = Region(GRAPHICS_PATH)

    def key_press(self, sender, args):
        self.Close()

    def form_load(self, sender, args):
        Cursor.Position = Point(SCREEN_SIZE_X, SCREEN_SIZE_Y)

    def Run(self):
        '''Не запускать два раза
        '''
        try:
            Application.Run(self)
        except:
            pass


    

