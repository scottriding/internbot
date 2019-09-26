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


class AmazonView(BoxLayout):

    def __init__(self, **kwargs):
        super(AmazonView, self).__init__(**kwargs)

        self.open_file_prompt = self.create_open_file_prompt()
        self.open_file_dialog = self.create_open_file_dialog()
        self.open_toc_prompt = self.create_toc_prompt()
        self.open_toc_dialog = self.create_toc_dialog()
        self.trended_selector = self.create_trended_selector()
        self.save_file_prompt = self.create_save_file_prompt()
        self.save_file_dialog = self.create_save_file_dialog()

    def create_open_file_prompt(self):
        label = Label(text="Choose unformatted Amazon SPSS (.xlsx) crosstab report")
        label.font_family= "Y2"

        popup = Popup(title="Select crosstab file",
        content=label,
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
                if ext != ".xlsx":
                    self.error_message("Please pick an Amazon report (.xlsx) file")
                else:
                    self.__open_filename = filepath
                    self.open_file_dialog_to_toc_prompt()
            except IndexError:
                self.error_message("Please pick an Amazon report (.xlsx) file")

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

    def create_toc_prompt(self):
        label = Label(text="Choose table of contents (.csv) crosstab file")
        label.font_family= "Y2"

        popup = Popup(title="Select table of contents file",
        content=label,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5},
        on_dismiss=self.toc_prompt_to_dialog)

        return popup

    def create_toc_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                path, ext = os.path.splitext(filepath)
                if ext != ".csv":
                    self.error_message("Please pick table of contents (.csv) file")
                else:
                    self.__toc_filename = filepath
                    self.toc_dialog_to_trended_selector()
            except IndexError:
                self.error_message("Please pick a table of contents (.csv) file")

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

    def create_trended_selector(self):
        chooser = BoxLayout(orientation='vertical')

        text = "Does this report have grouped or trended banners?"
        label = Label(text=text)
        label.font_family = "Y2"

        chooser.add_widget(label)

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)
        yes_btn = Button(text="Yes", on_press=self.is_trended)

        no_btn = Button(text="No", on_press=self.is_basic)

        button_layout.add_widget(yes_btn)
        button_layout.add_widget(no_btn)

        chooser.add_widget(button_layout)

        trended_chooser = Popup(title='Trended crosstabs',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return trended_chooser

    def create_save_file_prompt(self):
        label = Label(text="Choose a file location and name for Amazon crosstabs report")
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
            filepath = os.path.join(path, filename)
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
        self.__is_trended = False
        self.__controller = controller
        self.open_file_prompt.open()

    def open_file_prompt_to_dialog(self, instance):
        self.open_file_dialog.open()

    def open_file_dialog_to_toc_prompt(self):
        self.open_file_dialog.dismiss()
        self.open_toc_prompt.open()

    def toc_prompt_to_dialog(self, instance):
        self.open_toc_dialog.open()

    def toc_dialog_to_trended_selector(self):
        self.open_toc_dialog.dismiss()
        self.trended_selector.open()

    def is_trended(self, instance):
        self.__is_trended = True
        self.trended_selector.dismiss()
        self.save_file_prompt.open()

    def is_basic(self, instance):
        self.trended_selector.dismiss()
        self.save_file_prompt.open()

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_dialog.open()

    def finish(self):
        self.save_file_dialog.dismiss()
        try:
            workbook = self.__controller.rename(self.__open_filename, self.__toc_filename)
            self.__controller.highlight(workbook, self.__save_filename, self.__is_trended)
        except:
            self.error_message("Issue formatting report.")

    def error_message(self, error):
        label = Label(text=error)
        label.font_family= "Y2"

        popup = Popup(title="Something Went Wrong",
        content=label,
        size_hint=(.5, .8), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
