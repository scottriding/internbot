from view.crosstabs_view.format_report_view import FormatReportView
from view.crosstabs_view.toc_view import TOCView

## outside modules
import kivy
kivy.require('1.11.1')

from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.label import Label
import webbrowser
import os


class QResearchView(BoxLayout):

    def __init__(self, **kwargs):
        super(QResearchView, self).__init__(**kwargs)

        self.__controller = None

        self.qresearch_screen = self.create_qresearch_screen()

        self.toc_screen = TOCView()
        self.report_screen = FormatReportView()

        self.add_widget(self.qresearch_screen)

    def create_qresearch_screen(self):
        qresearch_screen = BoxLayout(orientation='vertical')

        back_arrow = "<"
        self.__back_btn = Button(text=back_arrow, size_hint=(.1, .05))
        self.__back_btn.font_name = "Y2"

        qresearch_screen.add_widget(self.__back_btn)

        double_back_arrow = "<<"
        self.__double_back_btn = Button(text=double_back_arrow, size_hint=(.1, .05))
        self.__double_back_btn.font_name = "Y2"

        qresearch_screen.add_widget(self.__double_back_btn)

        help_btn = Button(text='Help', size_hint=(.1, .05), on_press = self.qresearch_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (.9529, 0.6, .2392, 1)

        qresearch_screen.add_widget(help_btn)

        button_layout = BoxLayout()
        toc_btn = Button(text='Create table of contents', size_hint=(.5, .1), on_press = self.build_toc)
        toc_btn.font_name = "Y2"
        
        format_btn = Button(text='Format report', size_hint=(.5, .1), on_press = self.format_report)
        format_btn.font_name = "Y2"

        button_layout.add_widget(toc_btn)
        button_layout.add_widget(format_btn)

        qresearch_screen.add_widget(button_layout)

        return qresearch_screen

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

    def build_toc(self, instance):
        self.toc_screen.run(self.__controller)

    def format_report(self, instance):
        self.report_screen.run(self.__controller)

    def qresearch_help(self, instance):
        help_text = "QResearch is statistical software located on the new office PC.\n\n"
        help_text += "Before building a report in QResearch, you will need the \n"
        help_text += "survey file (.qsf) of the project to produce a table of contents\n"
        help_text += "in internbot (Create table of contents) and the column variables \n"  
        help_text += "or banners for building the report in the QResearch software. \n\n"
        help_text += "Once the QResearch unformatted report is generated in QResearch \n"
        help_text += "internbot will format it to Qualtrics/Y2 standards (Format report).\n\n"
        help_text += "[ref=click][color=F3993D]Click here for QResearch files and further instructions[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/luixf0mkcfxlfjc/AAC3ncnz0dAdUdQjdJXTeCEsa?dl=0")


        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family = "Y2"

        popup = Popup(title='QResearch Help',
        content=label,
        size_hint=(.6, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
