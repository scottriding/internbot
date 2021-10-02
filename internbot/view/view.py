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
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.image import Image
from kivy.graphics import Color, Rectangle
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.core.text import LabelBase
from kivy.core.audio import SoundLoader
from kivy.uix.relativelayout import RelativeLayout
from kivy.core.window import Window
from kivy.properties import BooleanProperty, ObjectProperty
from kivy.factory import Factory
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
        
        self.title = "Internbot - 1.3.1"

        self.root.bind(size=self._update_rect, pos=self._update_rect)
        self.root.bind(on_close=self.play_close)

        with self.root.canvas.before:
            self.rect = Rectangle(size=self.root.size, pos=self.root.pos, source='resources/images/DWTWNSLC.png')

        self.open_sound = SoundLoader.load('resources/sounds/open.mp3')
        self.close_sound = SoundLoader.load('resources/sounds/close.mp3')
        self.icon = 'resources/images/y2.icns'

        self.open_sound.play()

        self.__controller = None

        Window.bind(on_request_close=self.play_close)

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
        self.goodbye_message()
        return True

    def goodbye_message(self):
        bye_content = BoxLayout(orientation="vertical")
        desc_text = "Thanks for using internbot!" 

        desc_label = (Label(text=desc_text, markup=True))
        desc_label.font_family = "Y2"

        bye_content.add_widget(desc_label)

        empty_label = Label(text=" ")
        empty_label.size_hint=(1,.3)
        bye_content.add_widget(empty_label)

        cool_btn = Button(text='cool',
						pos_hint={'center_x': 0.5, 'center_y': 0.15},
        				size_hint=(.3, .5)) 
              
        bye_content.add_widget(cool_btn)

        popup = Popup(title='Bye',
        content=bye_content,
        auto_dismiss=False,
        separator_color=[243/255.,153/255.,61/255.,1.],
        size_hint=(.3, .3), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        cool_btn.bind(on_press=self.stop) 
        popup.open()

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

        xtabs_btn = HoverButton(text="Crosstab Reports", size_hint=(.5,.2), on_press = self.main_to_xtabs)
        xtabs_btn.font_name = "Y2"

        top_btn = HoverButton(text="Topline Reports", size_hint=(.5,.2), on_press = self.main_to_top)
        top_btn.font_name = "Y2"
 
        rnc_btn = HoverButton(text="RNC Reports", size_hint=(.5,.2), on_press = self.main_to_rnc)
        rnc_btn.font_name = "Y2"
        
        help_btn = HoverButton(text="Help", size_hint=(.5,.15), on_press = self.main_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (3/255, 169/255, 244/255, 1)

        button_layout.add_widget(Label(text=" ", size_hint=(.5, .01)))
        button_layout.add_widget(xtabs_btn)
        button_layout.add_widget(Label(text=" ", size_hint=(.5, .01)))
        button_layout.add_widget(top_btn)
        button_layout.add_widget(Label(text=" ", size_hint=(.5, .01)))
        button_layout.add_widget(rnc_btn)
        button_layout.add_widget(Label(text=" ", size_hint=(.5, .01)))
        button_layout.add_widget(help_btn)
        button_layout.add_widget(Label(text=" ", size_hint=(.5, .01)))

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
        desc_label.font_family = "Y2"

        help_content.add_widget(desc_label)

        help_text = "[ref=click][color=03a9f4][u]Click here for examples of reports[/u][/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/ia0c3b2lvq8kb68/AAAV4ELm-FPW1giqs4RZdRMra?dl=0")
        
        help_label = HoverLink(text=help_text, markup=True, valign= "top")
        help_label.bind(on_ref_press=examples_link)
        help_label.font_family = "Y2"
        help_label.size_hint=(.5,.3)
        help_label.pos_hint={'center_x': 0.5, 'center_y': 0.2}

        help_content.add_widget(help_label)

        empty_label = Label(text=" ")
        empty_label.size_hint=(1,.3)
        help_content.add_widget(empty_label)

        confirm_btn = HoverButton(text='confirm',
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

"""Hoverable Behaviour (changing when the mouse is on the widget by O. Poyen.
License: LGPL
"""
__author__ = 'Olivier POYEN @ https://gist.github.com/opqopq/15c707dc4cffc2b6455f' 

class HoverBehavior(object):
    """Hover behavior.
    :Events:
        `on_enter`
            Fired when mouse enter the bbox of the widget.
        `on_leave`
            Fired when the mouse exit the widget 
    """

    hovered = BooleanProperty(False)
    border_point= ObjectProperty(None)
    '''Contains the last relevant point received by the Hoverable. This can
    be used in `on_enter` or `on_leave` in order to know where was dispatched the event.
    '''

    def __init__(self, **kwargs):
        self.register_event_type('on_enter')
        self.register_event_type('on_leave')
        Window.bind(mouse_pos=self.on_mouse_pos)
        super(HoverBehavior, self).__init__(**kwargs)

    def on_mouse_pos(self, *args):
        if not self.get_root_window():
            return # do proceed if I'm not displayed <=> If have no parent
        pos = args[1]
        #Next line to_widget allow to compensate for relative layout
        inside = self.collide_point(*self.to_widget(*pos))
        if self.hovered == inside:
            #We have already done what was needed
            return
        self.border_point = pos
        self.hovered = inside
        if inside:
            self.dispatch('on_enter')
        else:
            self.dispatch('on_leave')

    def on_enter(self):
        pass

    def on_leave(self):
        pass

class HoverLink(Label, HoverBehavior):
    def on_enter(self, *args):
        Window.set_system_cursor('hand')

    def on_leave(self, *args):
        Window.set_system_cursor('arrow')

class HoverButton(Button, HoverBehavior):
    def on_enter(self, *args):
        Window.set_system_cursor('hand')

    def on_leave(self, *args):
        Window.set_system_cursor('arrow')

Factory.register('HoverBehavior', HoverBehavior)
Factory.register('HoverLink', HoverLink)
Factory.register('HoverButton', HoverButton)
        
