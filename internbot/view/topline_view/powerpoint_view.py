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


class PowerpointView(BoxLayout):

    def __init__(self, **kwargs):
        super(PowerpointView, self).__init__(**kwargs)

        self.open_survey_prompt = self.create_survey_prompt()
        self.open_survey_dialog = self.create_survey_dialog()
        self.open_template_prompt = self.create_template_prompt()
        self.open_template_dialog = self.create_open_template_dialog()
        self.open_freq_prompt = self.create_open_freq_prompt()
        self.open_freq_dialog = self.create_open_freq_dialog()
        self.save_file_prompt = self.create_save_file_prompt()
        self.save_file_dialog = self.create_save_file_dialog()

    def create_survey_prompt(self):
        label = Label(text="Choose a survey (.qsf) file")
        label.font_family= "Y2"

        popup = Popup(title="Select survey file",
        content=label,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5},
        on_dismiss=self.open_survey_prompt_to_dialog)

        return popup

    def create_survey_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                path, ext = os.path.splitext(filepath)
                if ext != ".qsf":
                    self.error_message("Please pick a survey (.qsf) file")
                else:
                    self.open_survey_dialog_to_template_prompt()
            except IndexError:
                self.error_message("Please pick a survey (.qsf) file")

        filechooser = FileChooserListView()
        filechooser.path = os.path.expanduser("~")
        filechooser.bind(on_selection=lambda x: filechooser.selection)

        open_btn = Button(text='open', size_hint=(.2,.1), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        open_btn.bind(on_release=lambda x: open_file(filechooser.path, filechooser.selection))

        container.add_widget(filechooser)
        container.add_widget(open_btn)
        chooser.add_widget(container)

        file_chooser = Popup(title='Open file',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return file_chooser 

    def create_template_prompt(self):
        label = Label(text="Choose a powerpoint template (.pptx) file")
        label.font_family= "Y2"

        popup = Popup(title="Select template file",
        content=label,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5},
        on_dismiss=self.open_template_prompt_to_dialog)

        return popup

    def create_open_template_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                path, ext = os.path.splitext(filepath)
                if ext != ".pptx":
                    self.error_message("Please pick a powerpoint template (.pptx) file")
                else:
                    self.open_template_dialog_to_freq_prompt()
            except IndexError:
                self.error_message("Please pick a powerpoint template (.pptx) file")

        filechooser = FileChooserListView()
        filechooser.path = os.path.expanduser("~")
        filechooser.bind(on_selection=lambda x: filechooser.selection)

        open_btn = Button(text='open', size_hint=(.2,.1), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        open_btn.bind(on_release=lambda x: open_file(filechooser.path, filechooser.selection))

        container.add_widget(filechooser)
        container.add_widget(open_btn)
        chooser.add_widget(container)

        file_chooser = Popup(title='Open file',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return file_chooser 

    def create_open_freq_prompt(self):
        label = Label(text="Choose a frequencies (.csv) file")
        label.font_family= "Y2"

        popup = Popup(title="Select frequencies file",
        content=label,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5},
        on_dismiss=self.open_freq_prompt_to_dialog)

        return popup

    def create_open_freq_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                path, ext = os.path.splitext(filepath)
                if ext != ".csv":
                    self.error_message("Please pick a frequencies (.csv) file")
                else:
                    self.open_freq_dialog_to_save_prompt()
            except IndexError:
                self.error_message("Please pick a frequencies (.csv) file")

        filechooser = FileChooserListView()
        filechooser.path = os.path.expanduser("~")
        filechooser.bind(on_selection=lambda x: filechooser.selection)

        open_btn = Button(text='open', size_hint=(.2,.1), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        open_btn.bind(on_release=lambda x: open_file(filechooser.path, filechooser.selection))

        container.add_widget(filechooser)
        container.add_widget(open_btn)
        chooser.add_widget(container)

        file_chooser = Popup(title='Open file',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return file_chooser 

    def create_save_file_prompt(self):
        label = Label(text="Choose a file location and name for topline appendix report")
        label.font_family= "Y2"

        popup = Popup(title="Select save file location",
        content=label,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5},
        on_dismiss=self.save_file_prompt_to_dialog)

        return popup

    def create_save_file_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        filechooser = FileChooserIconView()
        filechooser.path = os.path.expanduser("~")

        container.add_widget(filechooser)

        def save_file(path, filename):
            self.finish()

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)

        file_name = TextInput(text="File name.pptx")
        button_layout.add_widget(file_name)

        save_btn = Button(text='save', size_hint=(.2,1))
        save_btn.bind(on_release=lambda x: save_file(filechooser.path, file_name.text))

        button_layout.add_widget(save_btn)
        container.add_widget(button_layout)
        chooser.add_widget(container)

        file_chooser = Popup(title='Save report',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return file_chooser

    def run(self):
        self.open_survey_prompt.open()

    def open_survey_prompt_to_dialog(self, instance):
        self.open_survey_dialog.open()

    def open_survey_dialog_to_template_prompt(self):
        self.open_survey_dialog.dismiss()
        self.open_template_prompt.open()

    def open_template_prompt_to_dialog(self, instance):
        self.open_template_dialog.open()

    def open_template_dialog_to_freq_prompt(self):
        self.open_template_dialog.dismiss()
        self.open_freq_prompt.open()

    def open_freq_prompt_to_dialog(self, instance):
        self.open_freq_dialog.open()

    def open_freq_dialog_to_save_prompt(self):
        self.open_freq_dialog.dismiss()
        self.save_file_prompt.open()

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_dialog.open()

    def finish(self):
        self.save_file_dialog.dismiss()

    def error_message(self, error):
        label = Label(text=error)
        label.font_family= "Y2"

        popup = Popup(title="Something Went Wrong",
        content=label,
        size_hint=(.5, .8), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
