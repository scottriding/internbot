## internbot view modules
from view import crosstabs_view
from view import topline_view
from view import rnc_view

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
from kivy.core.audio import SoundLoader
import webbrowser
import os


class View(App):

    def build(self):
        KIVY_FONTS = [
        {
            "name": "Y2",
            "fn_regular": "resources/fonts/GothamBook.otf"
        }
        ]
        for font in KIVY_FONTS:
            LabelBase.register(**font)
        
        self.root = self.create_screens()
        self.title = "Internbot"
        self.root.bind(size=self._update_rect, pos=self._update_rect)
        self.root.bind(on_close=self.play_close)

        with self.root.canvas.before:
            self.rect = Rectangle(size=self.root.size, pos=self.root.pos, source='resources/images/DWTWNSLC.png')

        self.open_sound = SoundLoader.load('resources/sounds/open.mp3')
        self.close_sound = SoundLoader.load('resources/sounds/close.mp3')

        self.open_sound.play()

        self.__controller = None

        return self.root

    @property
    def controller(self):
        return self.__controller

    @controller.setter
    def controller(self, controller):
        self.__controller = controller

    def _update_rect(self, instance, value):
        self.rect.pos = instance.pos
        self.rect.size = instance.size

    def play_close(self, instance):
        self.close_sound.play()

    def create_screens(self):
        layered_menu = BoxLayout()

        self.main_screen = self.create_main_screen()

        self.crosstabs_screen = crosstabs_view.CrosstabsView()
        self.crosstabs_screen.xtabs_to_main_btn.bind(on_press=self.xtabs_to_main)
        self.crosstabs_screen.qresearch_to_main_btn.bind(on_press=self.xtabs_to_main)
        self.crosstabs_screen.controller = self.__controller

        self.topline_screen = topline_view.ToplineView()
        self.topline_screen.back_button.bind(on_press=self.top_to_main)
        self.topline_screen.controller = self.__controller

        self.rnc_screen = rnc_view.RNCView()
        self.rnc_screen.back_button.bind(on_press=self.rnc_to_main)
        self.rnc_screen.controller = self.__controller

        layered_menu.add_widget(self.main_screen)
        return layered_menu

    def create_main_screen(self):
        main_screen = BoxLayout()

        ## buttons
        button_layout = BoxLayout(orientation='vertical')

        xtabs_btn = Button(text="Crosstab Reports", size_hint=(.5,.3), on_press = self.main_to_xtabs)
        xtabs_btn.font_name = "Y2"

        top_btn = Button(text="Topline Reports", size_hint=(.5,.3), on_press = self.main_to_top)
        top_btn.font_name = "Y2"
 
        rnc_btn = Button(text="RNC Reports", size_hint=(.5,.3), on_press = self.main_to_rnc)
        rnc_btn.font_name = "Y2"
        
        help_btn = Button(text="Help", size_hint=(.5,.15), on_press = self.main_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (0.07, 0.306, 0.651, 1)

        button_layout.add_widget(xtabs_btn)
        button_layout.add_widget(top_btn)
        button_layout.add_widget(rnc_btn)
        button_layout.add_widget(help_btn)

        main_screen.add_widget(button_layout)

        # y2 logo
        screen_image = Image()
        screen_image.source = "resources/images/y2_white_logo.png"

        main_screen.add_widget(screen_image)

        return main_screen

    def main_to_xtabs(self, instance):
        self.root.remove_widget(self.main_screen)
        self.root.add_widget(self.crosstabs_screen)
        self.rect.source = 'resources/images/DWTWNSLC3.jpg'

    def xtabs_to_main(self, instance):
        self.root.remove_widget(self.crosstabs_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'resources/images/DWTWNSLC.png'

    def main_to_top(self, instance):
        self.root.remove_widget(self.main_screen)
        self.root.add_widget(self.topline_screen)
        self.rect.source = 'resources/images/DWTWNSLC2.jpg'

    def top_to_main(self, instance):
        self.root.remove_widget(self.topline_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'resources/images/DWTWNSLC.png'

    def main_to_rnc(self, instance):
        self.root.remove_widget(self.main_screen)
        self.root.add_widget(self.rnc_screen)
        self.rect.source = 'resources/images/DWTWNSLC1.jpg'

    def rnc_to_main(self, instance):
        self.root.remove_widget(self.rnc_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'resources/images/DWTWNSLC.png'

    def main_help(self, instance):
        help_text = "Internbot is Y2 Analytics' report automation tool that \n" 
        help_text += "manages \"quantity\" in a report so that the analytical \n"
        help_text += "team can focus on the \"quality\" of a report.\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of reports[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/bbcle8f9nifk4bo/AAA-JGnsx1XnhLaD_5Z9oDZna?dl=0")


        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family = "Y2"

        popup = Popup(title='Internbot Help',
        content=label,
        size_hint=(.6, .4), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
