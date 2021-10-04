## crosstabs modules
from view.crosstabs_view.qresearch_view import QResearchView
from view.crosstabs_view.amazon_view import AmazonView

## outside modules
import kivy
kivy.require('1.11.1')

from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.core.text import LabelBase
from kivy.uix.relativelayout import RelativeLayout
import webbrowser
import os


class CrosstabsView(BoxLayout):

    def __init__(self, **kwargs):
        super(CrosstabsView, self).__init__(**kwargs)

        self.__controller = None

        self.crosstabs_screen = self.create_crosstabs_screen()
        self.qresearch_screen = QResearchView()
        self.qresearch_screen.back_button.bind(on_press=self.qresearch_to_xtabs)

        self.amazon_screen = AmazonView()

        self.__qresearch_to_main_btn = self.qresearch_screen.double_back_button
        self.qresearch_screen.double_back_button.bind(on_press=self.qresearch_to_xtabs)

        self.add_widget(self.crosstabs_screen)

    def create_crosstabs_screen(self):
        crosstabs_screen = BoxLayout(orientation='vertical')

        back_arrow = "<"
        self.__xtabs_to_main_btn = Button(text=back_arrow, size_hint=(.1, .05))

        crosstabs_screen.add_widget(self.__xtabs_to_main_btn)

        help_btn = Button(text='Help', size_hint=(.1, .05), on_press = self.xtabs_help)
        help_btn.background_normal = ''
        help_btn.background_color = (.9529, 0.6, .2392, 1)

        crosstabs_screen.add_widget(help_btn)

        button_layout = BoxLayout()

        qresearch_btn = Button(text='QResearch', size_hint=(.1, .1), on_press = self.xtabs_to_qresearch)

        amaz_btn = Button(text='Amazon', size_hint=(.1, .1), on_press = self.xtabs_to_amaz)
        
        button_layout.add_widget(amaz_btn)
        button_layout.add_widget(qresearch_btn)

        crosstabs_screen.add_widget(button_layout)

        return crosstabs_screen

    def xtabs_to_qresearch(self, instance):
        self.remove_widget(self.crosstabs_screen)
        self.add_widget(self.qresearch_screen)

    def xtabs_to_amaz(self, instance):
        self.amazon_screen.run(self.__controller)

    def qresearch_to_xtabs(self, instance):
        self.remove_widget(self.qresearch_screen)
        self.add_widget(self.crosstabs_screen)

    @property
    def xtabs_to_main_btn(self):
        return self.__xtabs_to_main_btn

    @property
    def qresearch_to_main_btn(self):
        return self.__qresearch_to_main_btn

    @property
    def controller(self):
        return self.__controller

    @controller.setter
    def controller(self, controller):
        self.__controller = controller
        self.qresearch_screen.controller = self.__controller

    def xtabs_help(self, instance):
        help_text = "\n \nCrosstabs are a report that \"crosses\" selected dataset \n" 
        help_text += "variables against the rest of the dataset variables \n"
        help_text += "visualized in a table form with statistically significant \n"
        help_text += "values highlighted.\n\n"
        help_text += "QResearch and SPSS are different crosstab generating \n"
        help_text += "software that require different automation from internbot.\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of crosstab reports[/color][/ref]"
        help_text += "\n\n\n\n\n\n"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/2b67i3hbj1teg5j/AACWbEIwGtqq5hK894ItmmGpa?dl=0")
            
        help_content = RelativeLayout()
        content_cancel = Button(text='confirm', 
						pos_hint={'center_x': 0.5, 'center_y': 0.15}, 
        				size_hint=(.25, .2),
        				background_normal='',
        				background_color=(0, 0.4, 1, 1))
        help_content.add_widget(content_cancel)
        help_label = (Label(text=help_text, markup=True))
        help_label.bind(on_ref_press=examples_link)
        help_content.add_widget(help_label)
		
        popup = Popup(title='Crosstabs Help',
        content=help_content,
        auto_dismiss=False,
        size_hint=(.6, .6), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        

        popup.open()
        content_cancel.bind(on_press=popup.dismiss)