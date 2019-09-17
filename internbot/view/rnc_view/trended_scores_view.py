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


class TrendedScoresView(BoxLayout):

    def __init__(self, **kwargs):
        super(TrendedScoresView, self).__init__(**kwargs)

        self.open_file_prompt = self.create_open_file_prompt()
        self.open_file_dialog = self.create_open_file_dialog()
        self.report_descriptions = self.create_report_descriptions()
        self.save_folder_prompt = self.create_save_folder_prompt()
        self.save_folder_dialog = self.create_save_folder_dialog()

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
                self.error_message("Error reading round number.")

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
        help_text = "Choose trended score reports (TSR) data (.csv) file \n\n"
        help_text += "[ref=click][color=F3993D]Click here for TSR data examples[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/9xxpm4kirspdtoj/AADOK8dXTKuHt8sn5t8KD7Sja?dl=0")

        report_label = Label(text=help_text, markup=True)
        report_label.bind(on_ref_press=examples_link)
        report_label.font_family = "Y2"

        popup = Popup(title="Select TSR data file",
        content=report_label,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5},
        on_dismiss=self.open_file_prompt_to_dialog)

        return popup

    def create_open_file_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                path, ext = os.path.splitext(filepath)
                if ext != ".csv":
                    self.error_message("Please pick a trended score reports data (.csv) file")
                else:
                    self.open_filepath = filepath
                    self.open_file_dialog_to_save_prompt()
            except IndexError:
                self.error_message("Please pick a trended score reports data (.csv) file")

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

    def create_save_folder_prompt(self):
        label = Label(text="Choose a folder and name for several trended score reports")
        label.font_family= "Y2"

        popup = Popup(title="Select location for multiple reports",
        content=label,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5},
        on_dismiss=self.save_folder_prompt_to_dialog)

        return popup

    def create_save_folder_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        folderchooser = FileChooserIconView()
        folderchooser.path = os.path.expanduser("~")

        container.add_widget(folderchooser)

        def save_directory(path, foldername):
            save_filepath = os.path.join(path, foldername)
            if os.path.isdir(save_filepath):
                self.save_filepath = save_filepath
            elif os.path.isfile(save_filepath):
                self.save_filepath = os.path.splitext(save_filepath)[0]
                os.mkdir(self.save_filepath)
            else:
                os.mkdir(save_filepath)
                self.save_filepath = save_filepath
            self.finish()

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)

        folder_name = TextInput(text="New folder name")
        button_layout.add_widget(folder_name)

        save_btn = Button(text='save', size_hint=(.2,1))
        save_btn.bind(on_release=lambda x: save_directory(folderchooser.path, folder_name.text))

        button_layout.add_widget(save_btn)
        container.add_widget(button_layout)
        chooser.add_widget(container)

        directory_chooser = Popup(title='Save reports',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return directory_chooser

    def run(self):
        self.report_descriptions.open()

    def report_desc_to_open_prompt(self, round):
        self.report_descriptions.dismiss()
        self.round = round
        self.open_file_prompt.open()

    def open_file_prompt_to_dialog(self, instance):
        self.open_file_dialog.open()

    def open_file_dialog_to_save_prompt(self):
        self.open_file_dialog.dismiss()
        try:
            #self.__generator.build_report(self.open_filepath, self.round)
            self.save_folder_prompt.open()
        except Exception:
            self.error_message("Error reading data file.")
        
    def save_folder_prompt_to_dialog(self, instance):
        self.save_folder_dialog.open()

    def finish(self):
        self.save_folder_dialog.dismiss()
#         try:
#             self.__generator.generate_trended_scores(self.save_filepath)
#         except Exception as inst:
#             self.error_message("Error in creating report: " + str(inst.args))

    def error_message(self, error):
        label = Label(text=error)
        label.font_family= "Y2"

        popup = Popup(title="Something Went Wrong",
        content=label,
        size_hint=(.5, .8), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
