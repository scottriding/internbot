import kivy
kivy.require('1.11.1')

from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView, FileChooserIconView
from kivy.uix.scrollview import ScrollView
from kivy.properties import StringProperty
import webbrowser
import os
import io
from contextlib import redirect_stdout


class FormatReportView(BoxLayout):

    def __init__(self, **kwargs):
        super(FormatReportView, self).__init__(**kwargs)

        self.open_file_prompt = self.create_open_file_prompt()
        self.open_file_dialog = self.create_open_file_dialog()
        self.format_selector = self.create_format_selector()
        self.save_file_prompt = self.create_save_file_prompt()
        self.save_file_dialog = self.create_save_file_dialog()

        self.open_file_path = ''
        self.image_dir = 'resources/images/'
        self.__is_qualtrics = True
        self.save_file_path = ''

    def create_open_file_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose unformatted QResearch (.xlsx) crosstab report\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of unformatted reports[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/zwiaf534dfsgtlf/AABmd51ihZVairDCoRZuPbHRa?dl=0")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family= "Y2"

        popup_layout.add_widget(label)

        save_btn = Button(text='>', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.open_file_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select crosstab file",
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
                self.open_file_dialog_to_selector()
            except IndexError:
                self.error_message("Please pick a QResearch report (.xlsx) file")

        filechooser = FileChooserListView()
        filechooser.path = os.path.expanduser("~")
        filechooser.bind(on_selection=lambda x: filechooser.selection)
        filechooser.filters = ["*.xlsx"]

        open_btn = Button(text='open', size_hint=(.2,.1), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        open_btn.bind(on_release=lambda x: open_file(filechooser.path, filechooser.selection))

        container.add_widget(filechooser)
        container.add_widget(open_btn)
        chooser.add_widget(container)

        file_chooser = Popup(title='Open file',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return file_chooser 

    def create_format_selector(self):
        chooser = BoxLayout(orientation='vertical')

        text = "Choose from the following format options."
        label = Label(text=text)
        label.font_family = "Y2"

        chooser.add_widget(label)

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)
        qualtrics_btn = Button(text="Qualtrics", on_press=self.is_qualtrics)

        y2_btn = Button(text="Y2 Analytics", on_press=self.is_y2)

        button_layout.add_widget(qualtrics_btn)
        button_layout.add_widget(y2_btn)

        chooser.add_widget(button_layout)

        format_chooser = Popup(title='Choose format',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return format_chooser 

    def create_save_file_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        label = Label(text="Choose a file location and name for QResearch crosstabs report")
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

        filechooser = FileChooserIconView()
        filechooser.path = os.path.expanduser("~")

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

        file_chooser = Popup(title='Save report',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return file_chooser

    def run(self, controller):
        self.__controller = controller
        self.open_file_prompt.open()

    def open_file_prompt_to_dialog(self, instance):
        self.open_file_prompt.dismiss()
        self.open_file_dialog.open()

    def open_file_dialog_to_selector(self):
        self.open_file_dialog.dismiss()
        self.format_selector.open()

    def is_qualtrics(self, instance):
        self.format_selector.dismiss()
        try:
            self.__controller.build_qresearch_report(self.__open_filename, self.__is_qualtrics)
            self.save_file_prompt.open()
        except:
            self.error_message("Issue formatting report.")
        
    def is_y2(self, instance):
        self.__is_qualtrics = False
        self.format_selector.dismiss()
        try:
            self.__controller.build_qresearch_report(self.__open_filename, self.__is_qualtrics)
            self.save_file_prompt.open()
        except:
            self.error_message("Issue formatting report.")

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_prompt.dismiss()
        self.save_file_dialog.open()

    def finish(self):
        try:
            self.__controller.save_qresearch_report(self.__save_filename)
            self.save_file_dialog.dismiss()
        except:
            self.error_message("Issue saving formatted report.")

    def error_message(self, error):
        label = Label(text=error)
        label.font_family= "Y2"

        popup = Popup(title="Something Went Wrong",
        content=label,
        size_hint=(.5, .8), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
