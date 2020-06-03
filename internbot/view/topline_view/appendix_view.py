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
from kivy.uix.filechooser import FileChooserListView
import webbrowser
import os

class AppendixView(BoxLayout):

    def __init__(self, **kwargs):
        super(AppendixView, self).__init__(**kwargs)

        self.__is_doc_report = True
        self.__is_qualtrics = False

        self.open_file_prompt = self.create_open_file_prompt()
        self.open_file_dialog = self.create_open_file_dialog()
        self.format_selector = self.create_format_selector()
        self.other_template_dialog = self.create_other_template_dialog()
        self.save_file_prompt = self.create_save_file_prompt()

    def create_open_file_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose labelled appendix verbatims (.csv) file\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of verbatim files[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/tmg33zeh71sb71k/AAB5tpanqADX96yB3VL5yLw_a?dl=0")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family= "Y2"

        popup_layout.add_widget(label)

        save_btn = Button(text='>', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.open_file_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select appendix file",
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
                self.open_file_dialog_to_format_selector()
            except IndexError:
                self.error_message("Please pick an appendix (.csv) file")

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

    def create_format_selector(self):
        chooser = BoxLayout(orientation='vertical')

        text = "Choose from the following format options."
        label = Label(text=text)
        label.font_family = "Y2"

        chooser.add_widget(label)

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)

        policy_btn = Button(text="Utah Policy", on_press=self.is_policy)
        y2_btn = Button(text="Y2 Analytics", on_press=self.is_y2)
        oth_btn = Button(text="Other", on_press=self.is_other)

        button_layout.add_widget(policy_btn)
        button_layout.add_widget(y2_btn)
        button_layout.add_widget(oth_btn)

        chooser.add_widget(button_layout)

        format_chooser = Popup(title='Choose format',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return format_chooser 

    def create_other_template_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                self.__template_file_path = filepath
                self.other_template_dialog_to_save()
            except IndexError:
                self.error_message("Please select a template document (.docx) file")

        filechooser = FileChooserListView()
        filechooser.path = os.path.expanduser("~")
        filechooser.bind(on_selection=lambda x: filechooser.selection)
        filechooser.filters = ["*.docx"]

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
        label = Label(text="Choose a file location and name for topline appendix report")
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
            if ext != ".xlsx" and ext != ".docx":
                if self.__is_doc_report:
                    filepath += ".docx"
                else:
                    filepath += ".xlsx"
            self.__save_filename = filepath
            self.finish()

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)

        if self.__is_doc_report:
            file_default = "File name.docx"
        else:
            file_default = "File name.xlsx"

        file_name = TextInput(text=file_default)
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
        self.__template_file_path = None
        self.__controller = controller
        self.open_file_prompt.open()

    def open_file_prompt_to_dialog(self, instance):
        self.open_file_prompt.dismiss()
        self.open_file_dialog.open()

    def open_file_dialog_to_format_selector(self):
        self.open_file_dialog.dismiss()
        self.format_selector.open()

    def is_y2(self, instance):
        self.__template_name = "Y2"
        self.format_selector.dismiss()
            
        self.save_file_prompt.open()

    def is_policy(self, instance):
        self.__template_name = "UT_POLICY"
        self.format_selector.dismiss()

        self.save_file_prompt.open()

    def is_other(self, instance):
        self.__template_name = "OTHER"
        self.format_selector.dismiss()

        self.other_template_dialog.open()

    def other_template_dialog_to_save(self):
        self.other_template_dialog.dismiss()
        self.save_file_prompt.open()

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_prompt.dismiss()
        self.save_file_dialog = self.create_save_file_dialog()
        self.save_file_dialog.open()

    def finish(self):
        self.save_file_dialog.dismiss()
        try:
            questions = self.__controller.build_appendix_model(self.__open_filename)
            self.__controller.build_appendix_report(questions, self.__save_filename, self.__template_name, self.__template_file_path)
        except KeyError as key_error:
            string = "Misspelled or missing column (%s):\n %s" % (type(key_error), str(key_error))
            self.error_message(string)
        except Exception as inst:
            string = "Error (%s):\n %s" % (type(inst), str(inst))
            self.error_message(string)
        
    def error_message(self, error):
        label = Label(text=error)
        label.font_family= "Y2"

        popup = Popup(title="Something Went Wrong",
        content=label,
        size_hint=(.5, .8), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
