## outside modules
import kivy
kivy.require('1.11.1')

from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView

import os
import io
from contextlib import redirect_stdout
import webbrowser


class TOCView(BoxLayout):

    def __init__(self, **kwargs):
        super(TOCView, self).__init__(**kwargs)

        self.open_file_path = ''
        self.navigated_path = os.path.expanduser("~")
        self.save_file_path = ''

        self.open_file_prompt = self.create_open_file_prompt()
        self.open_file_dialog = self.create_open_file_dialog()
        self.save_file_prompt = self.create_save_file_prompt()
        self.save_file_dialog = self.create_save_file_dialog()

    def create_open_file_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose Qualtrics survey file (.qsf) for project\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of generated table of contents[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/lyhon1rrgjv88xi/AABh9gVKfON5zVMUYKQL6pzwa?dl=0")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)

        popup_layout.add_widget(label)

        save_btn = Button(text='>', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.open_file_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select survey file",
        content=popup_layout,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return popup

    def create_open_file_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                self.__open_filename = filepath
                self.open_file_dialog_to_prompt()
            except IndexError:
                self.error_message("Please pick a survey (.qsf) file")

        filechooser = FileChooserListView()
        filechooser.path = os.path.expanduser("~")
        filechooser.bind(on_selection=lambda x: filechooser.selection)
        filechooser.filters = ["*.qsf"]

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
        popup_layout = BoxLayout(orientation='vertical')
        label = Label(text="Choose a file location and name for QResearch table of contents file.")

        popup_layout.add_widget(label)

        save_btn = Button(text='>', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.save_file_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select save file location",
        content=popup_layout,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return popup

    def create_save_file_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        filechooser = FileChooserListView()
        filechooser.path = self.navigated_path
        filechooser.show_hidden = False

        container.add_widget(filechooser)

        def save_file(path, filename):
            filepath = os.path.join(path, filename)
            path, ext = os.path.splitext(filepath)
            if ext != ".xlsx":
                filepath += ".xlsx"
            self.__save_filename = filepath
            self.finish()

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)
        file_name = TextInput(text="File name.xlsx")
        button_layout.add_widget(file_name)

        save_btn = Button(text='save', size_hint=(.2,1))
        save_btn.bind(on_release=lambda x: save_file(filechooser.path, file_name.text))

        button_layout.add_widget(save_btn)
        container.add_widget(button_layout)
        chooser.add_widget(container)

        file_chooser = Popup(title='Save TOC file',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return file_chooser

    def run(self, controller):
        self.__controller = controller
        self.open_file_prompt.open()

    def open_file_prompt_to_dialog(self, instance):
        self.open_file_prompt.dismiss()
        self.open_file_dialog.open()

    def open_file_dialog_to_prompt(self):
        self.open_file_dialog.dismiss()
        self.save_file_prompt.open()

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_prompt.dismiss()
        try:
            self.__survey = self.__controller.build_survey(self.__open_filename)
            self.save_file_dialog.open()
        except:
            self.error_message("Issue parsing survey file.")

    def finish(self):
        try:
            self.__controller.build_toc_report(self.__survey, self.__save_filename)
            self.save_file_dialog.dismiss()
        except:
            self.error_message("Error creating report.")

    def error_message(self, error):
        label = Label(text=error)

        popup = Popup(title="Something Went Wrong",
        content=label,
        size_hint=(.5, .8), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()

