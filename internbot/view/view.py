## internbot view modules
from view import crosstabs_view
from view import topline_view
from view import rnc_view
from view import gui_tools

## outside modules
import kivy
kivy.require('1.11.1')

from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.image import Image
from kivy.graphics import Color, Rectangle
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.core.audio import SoundLoader
from kivy.core.window import Window
import webbrowser
import os
import time

class View(App):

    def build(self):
        self.root = self.create_screens()
        
        self.title = "Internbot - 1.3.1"

        self.root.bind(size=self._update_rect, pos=self._update_rect)
        
        with self.root.canvas.before:
            self.rect = Rectangle(size=self.root.size, pos=self.root.pos, source='resources/images/DWTWNSLC.png')

        self.icon = 'resources/images/y2.icns'

        self.open_sound = SoundLoader.load('resources/sounds/open.mp3')
        self.close_sound = SoundLoader.load('resources/sounds/close.mp3')

        self.open_sound.play()
        Window.bind(on_request_close=self.play_close)

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

    def play_close(self, *args):
        self.close_sound.play()

        # delay so that the sound can actually play before application closes
        time.sleep(0.5)

        self.stop()

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

        xtabs_btn = gui_tools.HoverButton(text="Crosstab Reports", size_hint=(.4,.2), on_press = self.main_to_xtabs)

        top_btn = gui_tools.HoverButton(text="Topline Reports", size_hint=(.4,.2), on_press = self.main_to_top)
 
        rnc_btn = gui_tools.HoverButton(text="RNC Reports", size_hint=(.4,.2), on_press = self.main_to_rnc)
        
        help_btn = gui_tools.HoverButton(text="Help", size_hint=(.4,.15), on_press = self.main_help)
        help_btn.background_normal = ''
        help_btn.background_color = (3/255, 169/255, 244/255, 1)

        button_layout.add_widget(Label(text=" ", size_hint=(.4, .01)))
        button_layout.add_widget(xtabs_btn)
        button_layout.add_widget(Label(text=" ", size_hint=(.4, .01)))
        button_layout.add_widget(top_btn)
        button_layout.add_widget(Label(text=" ", size_hint=(.4, .01)))
        button_layout.add_widget(rnc_btn)
        button_layout.add_widget(Label(text=" ", size_hint=(.4, .01)))
        button_layout.add_widget(help_btn)
        button_layout.add_widget(Label(text=" ", size_hint=(.4, .01)))

        main_screen.add_widget(BoxLayout(size_hint=(.01, 1)))
        main_screen.add_widget(button_layout)

        # y2 logo
        screen_image = Image()
        screen_image.source = "resources/images/y2_white_logo.png"

        main_screen.add_widget(screen_image)

        return main_screen

    def main_to_xtabs(self, instance):
        Window.set_system_cursor('arrow')
        self.root.remove_widget(self.main_screen)
        self.root.add_widget(self.crosstabs_screen)
        self.rect.source = 'resources/images/CANYON.jpg'

    def xtabs_to_main(self, instance):
        Window.set_system_cursor('hand')
        self.root.remove_widget(self.crosstabs_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'resources/images/DWTWNSLC.png'

    def main_to_top(self, instance):
        Window.set_system_cursor('arrow')
        self.root.remove_widget(self.main_screen)
        self.root.add_widget(self.topline_screen)
        self.rect.source = 'resources/images/TEMPLE.jpg'

    def top_to_main(self, instance):
        Window.set_system_cursor('hand')
        self.root.remove_widget(self.topline_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'resources/images/DWTWNSLC.png'

    def main_to_rnc(self, instance):
        Window.set_system_cursor('arrow')
        self.root.remove_widget(self.main_screen)
        self.root.add_widget(self.rnc_screen)
        self.rect.source = 'resources/images/CAPITOL.jpg'

    def rnc_to_main(self, instance):
        Window.set_system_cursor('hand')
        self.root.remove_widget(self.rnc_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'resources/images/DWTWNSLC.png'

    def main_help(self, instance):
        Window.set_system_cursor('arrow')
        help_content = BoxLayout(orientation="vertical")
        desc_text = "Internbot is Y2 Analytics' report automation tool that \n" 
        desc_text += "manages \"quantity\" in a report so that the analytical \n"
        desc_text += "team can focus on the \"quality\" of a report."

        desc_label = (Label(text=desc_text, markup=True))

        help_content.add_widget(desc_label)

        help_text = "[ref=click][color=03a9f4][u]Click here for examples of reports[/u][/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/ia0c3b2lvq8kb68/AAAV4ELm-FPW1giqs4RZdRMra?dl=0")
        
        help_label = gui_tools.HoverLink(text=help_text, markup=True, valign= "top")
        help_label.bind(on_ref_press=examples_link)
        help_label.size_hint=(.5,.3)
        help_label.pos_hint={'center_x': 0.5, 'center_y': 0.2}

        help_content.add_widget(help_label)

        empty_label = Label(text=" ")
        empty_label.size_hint=(1,.3)
        help_content.add_widget(empty_label)

        confirm_btn = gui_tools.HoverButton(text='confirm',
						pos_hint={'center_x': 0.5, 'center_y': 0.15},
        				size_hint=(.2, .3)) 
              
        help_content.add_widget(confirm_btn)

        popup = Popup(title='Internbot Help',
        content=help_content,
        auto_dismiss=False,
        separator_color=[243/255.,153/255.,61/255.,1.],
        size_hint=(.6, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        def fix_mouse(instance):
            Window.set_system_cursor('arrow')
            popup.dismiss()

        confirm_btn.bind(on_press=fix_mouse) 
        popup.open()
        
