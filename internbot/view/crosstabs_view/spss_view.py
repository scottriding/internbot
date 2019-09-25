## outside modules
import kivy
kivy.require('1.11.1')

from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.image import Image
from kivy.graphics import Color, Rectangle
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.core.text import LabelBase
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView, FileChooserIconView
import webbrowser
import os


class SPSSView(BoxLayout):

    def __init__(self, **kwargs):
        super(SPSSView, self).__init__(**kwargs)

        self.__controller = None

        self.spss_screen = self.create_spss_screen()

        self.add_widget(self.spss_screen)

    def create_spss_screen(self):
        spss_screen = BoxLayout(orientation='vertical')

        back_arrow = "<"
        self.__back_btn = Button(text=back_arrow, size_hint=(.1, .05))
        self.__back_btn.font_name = "Y2"

        spss_screen.add_widget(self.__back_btn)

        double_back_arrow = "<<"
        self.__double_back_btn = Button(text=double_back_arrow, size_hint=(.1, .05))
        self.__double_back_btn.font_name = "Y2"

        spss_screen.add_widget(self.__double_back_btn)

        help_btn = Button(text='Help', size_hint=(.1, .05), on_press = self.spss_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (.9529, 0.6, .2392, 1)

        spss_screen.add_widget(help_btn)

        button_layout = BoxLayout()
        var_btn = Button(text='Create variable script', size_hint=(.5, .1))
        var_btn.font_name = "Y2"
        
        table_btn = Button(text='Create table script', size_hint=(.5, .1))
        table_btn.font_name = "Y2"

        format_btn = Button(text='Format tables', size_hint=(.5, .1))
        format_btn.font_name = "Y2"

        button_layout.add_widget(var_btn)
        button_layout.add_widget(table_btn)
        button_layout.add_widget(format_btn)

        spss_screen.add_widget(button_layout)

        return spss_screen

    @property
    def back_button(self):
        return self.__back_btn

    @property
    def double_back_button(self):
        return self.__double_back_btn

    @property
    def controller(self):
        return self.__controller

    @controller.setter
    def controller(self, controller):
        self.__controller = controller

    def spss_help(self, instance):
        help_text = "SPSS is statistical software located on the old office PC.\n\n"
        help_text += "Before building a report in SPSS, you will need the \n"
        help_text += "survey file (.qsf) of the project to produce a variable script\n"
        help_text += "in internbot (Create variable script). And the column variables \n"  
        help_text += "or banners for creating table script in internbot\n" 
        help_text += "(Create table script). Both files will be used in generating\n" 
        help_text += "several table files in a folder that internbot will use to create\n"
        help_text += "the final report (Build report).\n\n"
        help_text += "[ref=click][color=F3993D]Click here for SPSS files and further instructions[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family = "Y2"

        popup = Popup(title='SPSS Help',
        content=label,
        size_hint=(.6, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
