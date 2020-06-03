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
from kivy.uix.filechooser import FileChooserListView
from kivy.core.window import Window, Keyboard

import webbrowser
import os


class IssueTrendedView(BoxLayout):

    def __init__(self, **kwargs):
        super(IssueTrendedView, self).__init__(**kwargs)

        self.open_file_prompt = self.create_open_file_prompt()
        self.open_file_dialog = self.create_open_file_dialog()
        self.report_descriptions = self.create_report_descriptions()
        self.save_file_prompt = self.create_save_file_prompt()
        self.save_file_dialog = self.create_save_file_dialog()

        self.open_filepath = ""
        self.round = 0
        self.save_filepath = ""
        
    def create_report_descriptions(self):
        inputter = BoxLayout(orientation='vertical')

        text = "Input the round number of report"

        label = Label(text=text)
        label.font_family = "Y2"

        inputter.add_widget(label)

        def inputted_round(round):
            try:
                self.report_desc_to_open_prompt(int(round))
            except:
                self.error_message("Error reading round number")

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)

        text_input = TextInput(text="Number of rounds:")

        button_layout.add_widget(text_input)

        enter_btn = Button(text="Enter", size_hint=(.3, 1),
        on_press=lambda x: inputted_round(text_input.text))

        button_layout.add_widget(enter_btn)

        inputter.add_widget(button_layout)

        round_inputter = Popup(title='Enter round numbers',
        content=inputter,
        size_hint=(.5, .6), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return round_inputter

    def create_open_file_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose issue trended report data (.csv) file \n\n"
        help_text += "[ref=click][color=F3993D]Click here for issue trended data examples[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/bcd2q8h848t7ha5/AABxfI7X026RIRqKbSzY1EhLa?dl=0")

        report_label = Label(text=help_text, markup=True)
        report_label.bind(on_ref_press=examples_link)
        report_label.font_family = "Y2"

        popup_layout.add_widget(report_label)

        save_btn = Button(text='>', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.open_file_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select issue trended data file",
        content=popup_layout,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return popup

    def create_open_file_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                self.open_filepath = filepath
                self.open_file_dialog_to_save_prompt()
            except IndexError:
                self.error_message("Please pick a issue trended report data (.csv) file")

        filechooser = FileChooserListView()
        filechooser.path = os.path.expanduser("~")
        filechooser.bind(on_selection=lambda x: filechooser.selection)
        filechooser.filters = ["*.csv"]

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
        label = Label(text="Choose a file location and name for issue trended report")
        label.font_family= "Y2"

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
        filechooser.path = os.path.expanduser("~")

        container.add_widget(filechooser)

        def save_file(path, filename):
            filepath = os.path.join(path, filename)
            path, ext = os.path.splitext(filepath)
            if ext != ".xlsx":
                filepath += ".xlsx"
            self.save_filepath = filepath
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

        file_chooser = Popup(title='Save report',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return file_chooser

    def run(self, controller):
        self.__controller = controller
        self.report_descriptions.open()

    def report_desc_to_open_prompt(self, round):
        self.round = round
        self.report_descriptions.dismiss()
        self.open_file_prompt.open()

    def open_file_prompt_to_dialog(self, instance):
        self.open_file_prompt.dismiss()
        self.open_file_dialog.open()

    def open_file_dialog_to_save_prompt(self):
        try:
            self.__controller.build_issues_model(self.open_filepath, self.round)
            self.open_file_dialog.dismiss()
            self.save_file_prompt.open()
        except KeyError as key_error:
            string = "Misspelled or missing column (%s):\n %s" % (type(key_error), str(key_error))
            self.error_message(string)
        except Exception as inst:
            string = "Error (%s):\n %s" % (type(inst), str(inst))
            self.error_message(string)

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_prompt.dismiss()
        self.save_file_dialog.open()

    def finish(self):
        try:
            self.__controller.build_issues_report(self.save_filepath)
            self.save_file_dialog.dismiss()
        except Exception as inst:
            self.error_message("Error in creating report: " + str(inst.args))

    def error_message(self, error):
        label = Label(text=error)
        label.font_family= "Y2"

        popup = Popup(title="Something Went Wrong",
        content=label,
        size_hint=(.5, .8), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
