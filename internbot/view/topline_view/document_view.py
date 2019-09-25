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

class DocumentView(BoxLayout):

    def __init__(self, **kwargs):
        super(DocumentView, self).__init__(**kwargs)

        self.__controller = None

        self.is_qsf = True
        self.__group_names = []
        self.__survey = None

        self.open_survey_prompt = self.create_open_survey_prompt()
        self.open_survey_dialog = self.create_open_survey_dialog()
        self.trended_selector = self.create_trended_selector()
        self.trended_count = self.create_trended_count()
        self.open_freq_prompt = self.create_open_freq_prompt()
        self.open_freq_dialog = self.create_open_freq_dialog()
        self.save_file_prompt = self.create_save_file_prompt()
        self.save_file_dialog = self.create_save_file_dialog()

    def create_open_survey_prompt(self):
        label = Label(text="Choose survey (.csv or .qsf) file")
        label.font_family= "Y2"

        popup = Popup(title="Select survey file",
        content=label,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5},
        on_dismiss=self.open_survey_prompt_to_dialog)

        return popup

    def create_open_survey_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                path, ext = os.path.splitext(filepath)
                if ext == ".csv":
                    self.is_qsf = False
                    self.__open_filename = filepath
                    self.open_survey_dialog_to_trended_selector()
                elif ext == ".qsf":
                    self.__open_filename = filepath
                    self.open_survey_dialog_to_trended_selector()
                else:
                    self.error_message("Please pick a survey (.csv or .qsf) file")
            except IndexError:
                self.error_message("Please pick a survey (.csv or .qsf) file")

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

        text = "Does this report have grouped or trended frequencies?"
        label = Label(text=text)
        label.font_family = "Y2"

        chooser.add_widget(label)

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)
        yes_btn = Button(text="Yes", on_press=self.trended_selector_to_count)

        no_btn = Button(text="No", on_press=self.trended_selector_to_freqs)

        button_layout.add_widget(yes_btn)
        button_layout.add_widget(no_btn)

        chooser.add_widget(button_layout)

        trended_chooser = Popup(title='Trended frequencies',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return trended_chooser

    def create_trended_count(self):
        inputter = BoxLayout(orientation='vertical')

        text = "Input the number of groups or rounds to report."

        label = Label(text=text)
        label.font_family = "Y2"

        inputter.add_widget(label)

        def inputted_groups(groups):
            self.create_trended_labels(int(groups))

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)

        text_input = TextInput(text="Number of groupings")

        button_layout.add_widget(text_input)

        enter_btn = Button(text="Enter", size_hint=(.1, 1),
        on_press=lambda x: inputted_groups(text_input.text))

        button_layout.add_widget(enter_btn)

        inputter.add_widget(button_layout)

        group_inputter = Popup(title='Enter group count',
        content=inputter,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return group_inputter

    def create_trended_labels(self, number_of_forms):
        self.trended_count.dismiss()
        self.trended_labels = Popup(size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        entry_layout = BoxLayout(orientation="vertical")
        dict = {}
        for i in range(0, number_of_forms):
            text = "Group #%s" % str(i+1)
            new_entry = TextInput(text=text, size_hint=(1,.1))
            entry_layout.add_widget(new_entry)
            dict[text] = new_entry

        def grab_labels():
            groups = []
            for i in range(0, number_of_forms):
                text = "Group #%s" % str(i+1)
                groups.append(dict.get(text).text)
            self.__group_names = groups
            self.trended_labels_to_freqs()

        enter_btn = Button(text="Enter", size_hint=(.2, .1), pos_hint={'center_x': 0.5, 'center_y': 0.5}, 
        on_press=lambda x: grab_labels())

        entry_layout.add_widget(enter_btn)

        self.trended_labels.content = entry_layout
        self.trended_labels.title = "Enter group names"
        self.trended_labels.open()

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
                    self.__open_filename = filepath
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
        label = Label(text="Choose a file location and name for topline document report")
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
            self.__save_filename = os.path.join(path, filename)
            self.finish()

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)

        file_name = TextInput(text="File name.docx")
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
        self.is_qsf = True
        self.__group_names = []
        self.__survey = None
        self.__controller = controller
        self.open_survey_prompt.open()

    def open_survey_prompt_to_dialog(self, instance):
        self.open_survey_dialog.open()

    def open_survey_dialog_to_trended_selector(self):
        self.open_survey_dialog.dismiss()
        
        if self.is_qsf:
            #self.__survey = self.__controller.build_survey(self.__open_filename)
            try:
                self.__survey = self.__controller.build_survey(self.__open_filename)
                self.trended_selector.open()
            except:
                self.error_message("Issue parsing .qsf file.")
        else:
            self.trended_selector.open()

    def trended_selector_to_count(self, instance):
        self.trended_selector.dismiss()
        self.trended_count.open()

    def trended_selector_to_freqs(self, instance):
        self.trended_selector.dismiss()
        if self.is_qsf:
            self.open_freq_prompt.open()
        else:
            self.save_file_prompt.open()

    def trended_labels_to_freqs(self):
        self.trended_labels.dismiss()
        if self.is_qsf:
            self.open_freq_prompt.open()
        else:
            self.save_file_prompt.open()

    def open_freq_prompt_to_dialog(self, instance):
        self.open_freq_dialog.open()

    def open_freq_dialog_to_save_prompt(self):
        self.open_freq_dialog.dismiss()
        self.save_file_prompt.open()

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_dialog.open()

    def finish(self):
        self.save_file_dialog.dismiss()
        try:
            self.__controller.build_document_model(self.__open_filename, self.__group_names, self.__survey)
        except:
            self.error_message("Issue parsing frequency file")

        try:
            self.__controller.build_document_report(self.__save_filename)
        except:
            self.error_message("Issue building topline report")

    def error_message(self, error):
        label = Label(text=error)
        label.font_family= "Y2"

        popup = Popup(title="Something Went Wrong",
        content=label,
        size_hint=(.5, .8), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
