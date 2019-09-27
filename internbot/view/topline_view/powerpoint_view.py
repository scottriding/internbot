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
        self.trended_selector = self.create_trended_selector()
        self.trended_count = self.create_trended_count()
        self.open_freq_prompt = self.create_open_freq_prompt()
        self.open_freq_dialog = self.create_open_freq_dialog()
        self.open_template_prompt = self.create_template_prompt()
        self.open_template_dialog = self.create_open_template_dialog()
        self.save_file_prompt = self.create_save_file_prompt()
        self.save_file_dialog = self.create_save_file_dialog()

    def create_survey_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose survey (.csv or .qsf) file\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of survey files[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/qrk41o9k3dc761d/AAAr3L7bk2GTOJEQBfU9m5F-a?dl=0")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family= "Y2"

        popup_layout.add_widget(label)

        save_btn = Button(text='>', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.open_survey_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select survey file",
        content=popup_layout,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return popup

    def create_survey_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                self.__open_filename = filepath
                self.open_survey_dialog_to_trended_selector()
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

    def create_trended_selector(self):
        chooser = BoxLayout(orientation='vertical')

        help_text = "Does this report have grouped or trended frequencies?\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of trended/not-trended reports[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/1qespwi66d6o8cp/AACML0z5Poii3XZFy4VEm1e2a?dl=0")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family= "Y2"

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
            try:
                count = int(groups)
                if count < 2:
                    self.error_message("2 or more items make a trended report")
                else:
                    self.create_trended_labels(count)
            except:
                self.error_message("Issue parsing grouping count")

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

        if number_of_forms < 5:
            self.trended_labels.size_hint = (.9, .4)
        else:
            self.trended_labels.size_hint = (.9, .7)

        pop_up_layout = BoxLayout(orientation="vertical")
        entry_layout = BoxLayout(orientation="vertical")
        dict = {}
        for i in range(0, number_of_forms):
            text = "Group #%s" % str(i+1)
            new_entry = TextInput(text=text)
            new_entry.size_hint = (1, .2)
            new_entry.write_tab = False
            entry_layout.add_widget(new_entry)
            dict[text] = new_entry

        def grab_labels():
            groups = []
            for i in range(0, number_of_forms):
                text = "Group #%s" % str(i+1)
                groups.append(dict.get(text).text)
            self.__group_names = groups
            self.trended_labels_to_freqs()

        enter_btn = Button(text="Enter", size_hint=(.2, .2), pos_hint={'center_x': 0.5, 'center_y': 0.5}, 
        on_press=lambda x: grab_labels())

        pop_up_layout.add_widget(entry_layout)
        pop_up_layout.add_widget(enter_btn)

        self.trended_labels.content = pop_up_layout
        self.trended_labels.title = "Enter group names"
        self.trended_labels.open()

    def create_open_freq_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose a frequencies (.csv) file\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of frequency files[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/5jxxjp9fd3djfj5/AAAXE1qgqw3Jk2kefD4cElvIa?dl=0")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family= "Y2"

        popup_layout.add_widget(label)

        save_btn = Button(text='>', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.open_freq_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select frequencies file",
        content=popup_layout,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return popup

    def create_open_freq_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                self.__freq_filename = filepath
                self.open_freq_dialog_to_open_template_prompt()
            except IndexError:
                self.error_message("Please pick a frequencies (.csv) file")

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

    def create_template_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose a powerpoint template (.pptx) file\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of template files[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/748h81mypofblv2/AADDz5cvC9O37s8g1URWsQqSa?dl=0")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family= "Y2"

        popup_layout.add_widget(label)

        save_btn = Button(text='>', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.open_template_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select template file",
        content=popup_layout,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return popup

    def create_open_template_dialog(self):
        chooser = BoxLayout()
        container = BoxLayout(orientation='vertical')

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                self.__template_filename = filepath
                self.open_template_dialog_to_save_prompt()
            except IndexError:
                self.error_message("Please pick a powerpoint template (.pptx) file")

        filechooser = FileChooserListView()
        filechooser.path = os.path.expanduser("~")
        filechooser.bind(on_selection=lambda x: filechooser.selection)
        filechooser.filters = ["*.pptx"]

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
        label = Label(text="Choose a file location and name for topline powerpoint report")
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
            if ext != ".pptx":
                filepath += ".pptx"
            self.__save_filename = filepath
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

    def run(self, controller):
        self.__group_names = []
        self.__controller = controller
        self.open_survey_prompt.open()
        
    def open_survey_prompt_to_dialog(self, instance):
        self.open_survey_prompt.dismiss()
        self.open_survey_dialog.open()

    def open_survey_dialog_to_trended_selector(self):
        self.open_survey_dialog.dismiss()
        self.__survey = self.__controller.build_survey(self.__open_filename)
        self.trended_selector.open()

    def trended_selector_to_count(self, instance):
        self.trended_selector.dismiss()
        self.trended_count.open()

    def trended_selector_to_freqs(self, instance):
        self.trended_selector.dismiss()
        self.open_freq_prompt.open()

    def trended_labels_to_freqs(self):
        self.trended_labels.dismiss()
        self.open_freq_prompt.open()

    def open_freq_prompt_to_dialog(self, instance):
        self.open_freq_prompt.dismiss()
        self.open_freq_dialog.open()

    def open_freq_dialog_to_open_template_prompt(self):
        self.open_freq_dialog.dismiss()
        self.__controller.build_powerpoint_model(self.__freq_filename, self.__group_names, self.__survey)
        self.open_template_prompt.open()

    def open_template_prompt_to_dialog(self, instance):
        self.open_template_prompt.dismiss()
        self.open_template_dialog.open()

    def open_template_dialog_to_save_prompt(self):
        self.open_template_dialog.dismiss()
        self.save_file_prompt.open()

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_prompt.dismiss()
        self.save_file_dialog.open()

    def finish(self):
        self.save_file_dialog.dismiss()
        self.__controller.build_powerpoint_report(self.__template_filename, self.__save_filename)

    def error_message(self, error):
        label = Label(text=error)
        label.font_family= "Y2"

        popup = Popup(title="Something Went Wrong",
        content=label,
        size_hint=(.5, .8), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
