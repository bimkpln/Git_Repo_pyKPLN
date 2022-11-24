# -*- coding: utf-8 -*-
from System.Windows.Forms import *
from System.Drawing import Point, SystemColors, Size, Color, Font, FontStyle
import webbrowser as wb

class InputWindow(Form):
    '''
    '''
    def __init__(self, levels_names):

        #region ПЕРЕМЕННЫЕ КЛАССА
        self.level_name = levels_names
        self.selected_text = None
        self.selected_start = None
        self.dis_operation = None
        self.from_level = None
        #endregion

        #region ОСНОВНАЯ ФОРМА
        self.BackColor = SystemColors.GradientInactiveCaption
        self.FormBorderStyle = FormBorderStyle.None
        self.StartPosition = FormStartPosition.CenterScreen
        self.ClientSize = Size(420, 200)
        self.TopMost = True
        #endregion

        #region ОКНО С ТЕКСТОМ
        self.textBox1 = TextBox()
        self.textBox1.Text = self.level_name
        self.textBox1.Location = Point(10, 100)
        self.textBox1.Size = Size(400, 22)
        self.textBox1.BorderStyle = BorderStyle.None
        self.textBox1.TextAlign = HorizontalAlignment.Center
        self.textBox1.Font = Font("Bahnschrift SemiBold Condensed", 16, FontStyle.Regular)
        self.textBox1.TextChanged += self.textBox1_TextChanged_1
        #endregion

        #region ОСНОВНАЯ КНОПКА
        self.button1 = Button()
        self.button1.Location = Point(10, 140)
        self.button1.Size = Size(400, 50)
        self.button1.BackColor = Color.Silver
        self.button1.Text = "Выделите часть текста, где обозначен уровень и нажмите сюда"
        self.button1.Font = Font("Bahnschrift SemiBold Condensed", 12, FontStyle.Regular)
        self.button1.FlatStyle = FlatStyle.Flat
        self.button1.Click += self.button1_Click
        self.button1.BackColor = SystemColors.ActiveCaptionText
        self.button1.ForeColor = SystemColors.Control
        #endregion

        #region КНОПКА ЗАКРЫТИЯ ФОРМЫ
        self.button2 = Button()
        self.button2.Location = Point(382, 0)
        self.button2.Size = Size(38, 22)
        self.button2.BackColor = Color.Silver
        self.button2.Text = "Х"
        self.button2.Font = Font("Bahnschrift SemiBold Condensed", 12, FontStyle.Regular)
        self.button2.FlatStyle = FlatStyle.Flat
        self.button2.BackColor = SystemColors.ActiveCaptionText
        self.button2.ForeColor = SystemColors.Control
        self.button2.Click += self.button2_Click
        #endregion

        #region КНОПКА СПРАВКИ
        self.button3 = Button()
        self.button3.Location = Point(345, 0)
        self.button3.Size = Size(32, 22)
        self.button3.BackColor = Color.Silver
        self.button3.Text = "?"
        self.button3.Font = Font("Bahnschrift SemiBold Condensed", 12, FontStyle.Regular)
        self.button3.FlatStyle = FlatStyle.Flat
        self.button3.Click += self.button3_Click
        self.button3.BackColor = SystemColors.ActiveCaptionText
        self.button3.ForeColor = SystemColors.Control
        #endregion

        #region ДОБАВЛЕНИЕ ОПЦИОНАЛЬНОСТИ
        self.checkBox1 = CheckBox()
        self.checkBox1.Location = Point(10, 70)
        self.checkBox1.Size = Size(300, 20)
        self.checkBox1.Text = "Записать из параметра уровня"
        self.checkBox1.CheckedChanged += self.checkBox1_CheckedChanged
        self.checkBox1.Font = Font("Bahnschrift SemiBold Condensed", 10, FontStyle.Regular)
        #region

        #region ОПИСАНИЕ
        self.label1 = Label()
        self.label1.Location = Point(0,0)
        self.label1.Size = Size(200, 100)
        self.label1.Text = "KPLN. Заполнение отметок уровня"
        self.label1.Font = Font("Bahnschrift SemiBold Condensed", 14, FontStyle.Regular)
        #endregion

        #region ПОДСКАЗКА
        #endregion

        #region ДОБАВЛЕНИЕ КОНТРОЛОВ
        self.Controls.Add(self.textBox1)
        self.Controls.Add(self.button1)
        self.Controls.Add(self.button2)
        self.Controls.Add(self.button3)
        self.Controls.Add(self.checkBox1)
        self.Controls.Add(self.label1)
        #endregion

    def button1_Click(self, sender, args):
        self.selected_text = str(self.textBox1.SelectedText)
        self.selected_start = str(self.textBox1.SelectionStart)
        self.Close()

    def button2_Click(self, sender, args):
        self.dis_operation = True
        self.Close()

    def button3_Click(self, sender, args):
        wb.open("http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=1098")

    def checkBox1_CheckedChanged(self, sender, args):
        if self.checkBox1.CheckState:
            self.textBox1.Visible = False
            self.from_level = True
        else:
            self.textBox1.Visible = True
            self.from_level = False
    
    def textBox1_TextChanged_1(self, sender, args):
        self.textBox1.Text = self.level_name