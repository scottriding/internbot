from view.topline_view.appendix_view import AppendixView
from view.topline_view.document_view import DocumentView
from view.topline_view.powerpoint_view import PowerpointView

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


class ToplineView(BoxLayout):

    def __init__(self, **kwargs):
        super(ToplineView, self).__init__(**kwargs)

        self.topline_screen = self.create_topline_screen()

        self.appendix_screen = AppendixView()
        self.document_screen = DocumentView()
        self.powerppt_screen = PowerpointView()

        self.__controller = None

        self.add_widget(self.topline_screen)

    def create_topline_screen(self):
        topline_screen = BoxLayout(orientation='vertical')

        back_arrow = "<"
        self.__back_btn = Button(text=back_arrow, size_hint=(.1, .05))

        topline_screen.add_widget(self.__back_btn)

        help_btn = Button(text='Help', size_hint=(.1, .05), on_press = self.top_help)
        help_btn.background_normal = ''
        help_btn.background_color = (.9529, 0.6, .2392, 1)

        topline_screen.add_widget(help_btn)

        button_layout = BoxLayout()

        docx_btn = Button(text='Appendix', size_hint=(.1, .1), on_press = self.build_appendix)

        pptx_btn = Button(text='Document', size_hint=(.1, .1), on_press = self.build_doc)

        app_btn = Button(text='Powerpoint', size_hint=(.1, .1), on_press = self.build_ppt)

        button_layout.add_widget(docx_btn)
        button_layout.add_widget(pptx_btn)
        button_layout.add_widget(app_btn)

        topline_screen.add_widget(button_layout)

        return topline_screen

    @property
    def back_button(self):
        return self.__back_btn

    @property
    def controller(self):
        return self.__controller

    @controller.setter
    def controller(self, controller):
        self.__controller = controller

    def build_appendix(self, instance):
        self.appendix_screen.run(self.__controller)

    def build_doc(self, instance):
        self.document_screen.run(self.__controller)

    def build_ppt(self, instance):
        self.powerppt_screen.run(self.__controller)

    def top_help(self, instance):
        help_text = "Topline reports are deliverables that give high level results\n"
        help_text += "for a survey project.\n\n"
        help_text += "An Appendix report shows the results of text entries or\n"
        help_text += "open-ended questions.\n\n"
        help_text += "Document and Powerpoint show topline frequencies verbatim\n"
        help_text += "in tables or charts, respectively.\n\n"
        help_text += "[ref=click][color=F3993D]Click here for Topline report examples[/color][/ref]"
        help_text += "\n\n\n\n"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/c653vu8jxx0d42u/AADlcD8HDL2J4HxXqDVT3bAVa?dl=0")
        
        help_content = RelativeLayout()
        content_cancel = Button(text='confirm', 
						pos_hint={'center_x': 0.5, 'center_y': 0.15}, 
        				size_hint=(.3, .2),
        				background_normal='',
        				background_color=(0, 0.4, 1, 1))
        help_content.add_widget(content_cancel)
        help_label = (Label(text=help_text, markup=True))
        help_label.bind(on_ref_press=examples_link)
        help_content.add_widget(help_label)
		
        popup = Popup(title='Topline Help',
        content=help_content,
        auto_dismiss=False,
        size_hint=(.6, .6), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        

        popup.open()
        content_cancel.bind(on_press=popup.dismiss)
