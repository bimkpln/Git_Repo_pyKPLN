#encoding: utf-8 python
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import *
from System.Drawing import Size, Point, Image, ContentAlignment, Font, FontStyle, Color, SystemColors
from pyrevit import script
from System import EventHandler
import webbrowser as wb

class KPLN_Alarm_Form(Form):
    def __init__(self, label_text):
        self.label_text = label_text
        self.script_path = script.get_script_path()

        self.IMAGE_LOGO = image_logo = Image.FromFile("".join([self.script_path, "\\", r"Images\Background.png"]))
        self.IMAGE_OK = Image.FromFile("".join([self.script_path, "\\", r"Images\button_OK.png"]))
        self.IMAGE_OK_WHITE = Image.FromFile("".join([self.script_path, "\\", r"Images\button_OK_White.png"]))
        self.IMAGE_Q = Image.FromFile("".join([self.script_path, "\\", r"Images\button_Q.png"]))
        self.IMAGE_Q_WHITE = Image.FromFile("".join([self.script_path, "\\", r"Images\button_Q_White.png"]))

        self.COLOR = SystemColors.GradientInactiveCaption
        self.Opacity = 0.95

        self.InitComponent()

    def InitComponent(self):
        #
        #   ОСНОВНАЯ ФОРМА
        #
        self.StartPosition = FormStartPosition.CenterScreen
        self.Size = Size(500, 300)
        self.FormBorderStyle = FormBorderStyle.None
        image_logo = self.IMAGE_LOGO
        self.BackgroundImage = image_logo
        self.BackColor = self.COLOR
        #
        #   КНОПКА ОК
        #
        self.button_OK = Button()
        self.button_OK.Location = Point(150, 215)
        self.button_OK.Size = Size(125, 56)
        self.button_OK.FlatAppearance.BorderSize = 0
        self.button_OK.FlatStyle = FlatStyle.Flat
        self.button_OK.BackgroundImage = self.IMAGE_OK
        self.button_OK.Cursor = Cursors.Hand
        self.button_OK.Click += self.click_OK
        self.button_OK.MouseEnter += EventHandler(self.change_OK_hover)
        self.button_OK.MouseLeave += EventHandler(self.change_OK_leave)
        #
        #   КНОПКА ?
        #
        self.button_Q = Button()
        self.button_Q.Location = Point(295, 215)
        self.button_Q.Size = Size(56, 56)
        self.button_Q.FlatAppearance.BorderSize = 0
        self.button_Q.FlatStyle = FlatStyle.Flat
        self.button_Q.BackgroundImage = self.IMAGE_Q
        self.button_Q.Cursor = Cursors.Hand
        self.button_Q.Click += self.click_Q
        self.button_Q.MouseEnter += EventHandler(self.change_Q_hover)
        self.button_Q.MouseLeave += EventHandler(self.change_Q_leave)
        #
        #   ЯРЛЫК ТЕКСТА
        #
        self.label = Label()
        self.label.Location = Point(36, 37)
        self.label.Size = Size(427, 164)
        self.label.Text = self.label_text
        self.label.TextAlign = ContentAlignment.MiddleCenter
        self.label.Font = Font("Century Gothic", 16, FontStyle.Regular)
        self.hidden_button = Button()
        self.hidden_button.Location = Point(490,0)
        self.hidden_button.Size = Size(10,10)
        self.hidden_button.BackColor = Color.Black
        self.hidden_button.FlatAppearance.BorderSize = 0
        self.hidden_button.FlatStyle = FlatStyle.Flat
        self.hidden_button.Click += self.go
        #
        #controls
        #
        self.Controls.Add(self.button_OK)
        self.Controls.Add(self.button_Q)
        self.Controls.Add(self.label)
        self.Controls.Add(self.hidden_button)

    def click_OK(self, sender, args):
        self.Close()

    def click_Q(self, sender, args):
        wb.open_new("http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=1100")

    def change_OK_hover(self, sender, args):
        self.button_OK.BackgroundImage = self.IMAGE_OK_WHITE

    def change_OK_leave(self, sender, args):
        self.button_OK.BackgroundImage = self.IMAGE_OK

    def change_Q_hover(self, sender, args):
        self.button_Q.BackgroundImage = self.IMAGE_Q_WHITE

    def change_Q_leave(self, sender, args):
        self.button_Q.BackgroundImage = self.IMAGE_Q

    def go(self, sender, args):
        wb.open_new("https://www.youtube.com/watch?v=dQw4w9WgXcQ")

class KPLN_Alarm:
    def __init__(self, text):
        Application.Run(KPLN_Alarm_Form(text))