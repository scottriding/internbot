import kivy
kivy.require('1.11.1')

from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView
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
        self.save_file_path = ''

    def create_open_file_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose unformatted QResearch (.xlsx) crosstab report\n\n"
        help_text += "[ref=click][color=F3993D][u]Click here for examples of unformatted reports[/u][/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/zwiaf534dfsgtlf/AABmd51ihZVairDCoRZuPbHRa?dl=0")

        popup = Popup(title="",
        separator_height = 0,
        content=popup_layout,
        size_hint=(.7, .5), 
        pos_hint={'center_x': 0.5, 'center_y': 0.5})

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)

        popup_layout.add_widget(label)

        next_btn = Button(text='Next', size_hint=(.2,.2))
        next_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        next_btn.bind(on_release=self.open_file_prompt_to_dialog)

        popup_layout.add_widget(next_btn)

        return popup

    def create_open_file_dialog(self):
        chooser_layout = BoxLayout(orientation='vertical')
        container = BoxLayout(orientation='vertical')

        file_chooser = Popup(title='',
        separator_height = 0,
        content=chooser_layout,
        size_hint=(.9, .7), 
        pos_hint={'center_x': 0.5, 'center_y': 0.5})

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                self.__open_filename = filepath
                self.open_file_dialog_to_selector()
            except IndexError:
                self.error_message("Please pick a QResearch report (.xlsx) file")

        chooser_view = FileChooserListView()
        chooser_view.path = os.path.expanduser("~")
        chooser_view.bind(on_selection=lambda x: chooser_view.selection)
        chooser_view.filters = ["*.xlsx"]

        open_btn = Button(text='open', size_hint=(.2,.1), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        open_btn.bind(on_release=lambda x: open_file(chooser_view.path, chooser_view.selection))

        container.add_widget(chooser_view)
        container.add_widget(open_btn)
        chooser_layout.add_widget(container)

        return file_chooser 

    def create_format_selector(self):
        chooser = BoxLayout(orientation='vertical')

        popup = Popup(title="",
        separator_height = 0,
        content=chooser,
        size_hint=(.9, .7), 
        pos_hint={'center_x': 0.5, 'center_y': 0.5})

        text = "Choose from the following format options."
        label = Label(text=text)
        label.font_family = "Y2"

        chooser.add_widget(label)

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)

        y2_btn = Button(text="Y2 Analytics", on_press=self.is_y2)
        qualtrics_btn = Button(text="Qualtrics", on_press=self.is_qualtrics)
        wa_btn = Button(text="WhatsApp", on_press=self.is_whatsapp)
        fb_btn = Button(text="Facebook", on_press=self.is_fb)

        button_layout.add_widget(y2_btn)
        button_layout.add_widget(qualtrics_btn)
        button_layout.add_widget(wa_btn)
        button_layout.add_widget(fb_btn)

        chooser.add_widget(button_layout)

        return popup 

    def create_save_file_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        label = Label(text="Choose a file location and name for QResearch crosstabs report")

        popup = Popup(title="",
        separator_height = 0,
        content=popup_layout,
        size_hint=(.7, .5), 
        pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup_layout.add_widget(label)

        save_btn = Button(text='Next', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.save_file_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        return popup

    def create_save_file_dialog(self):
        chooser_layout = BoxLayout(orientation='vertical')
        container = BoxLayout(orientation='vertical')

        file_chooser = Popup(title='',
        separator_height = 0,
        content=chooser_layout,
        size_hint=(.9, .7), 
        pos_hint={'center_x': 0.5, 'center_y': 0.5})

        menu_layout = BoxLayout(orientation='vertical')
        menu_layout.size_hint = (.1, .2)

        chooser_view = FileChooserListView()
        chooser_view.path = os.path.expanduser("~")

        container.add_widget(chooser_view)

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
        save_btn.bind(on_release=lambda x: save_file(chooser_view.path, file_name.text))

        button_layout.add_widget(save_btn)
        container.add_widget(button_layout)
        chooser_layout.add_widget(container)

        return file_chooser

    def run(self, controller):
        self.__controller = controller
        self.open_file_prompt.open()

    def open_file_prompt_to_dialog(self, instance):
        self.open_file_prompt.dismiss()
        self.open_file_dialog.open()

    def open_dialog_to_prompt(self, instance):
        self.open_file_dialog.dismiss()
        self.open_file_prompt.open()

    def open_file_dialog_to_selector(self):
        self.open_file_dialog.dismiss()
        self.format_selector.open()

    def selector_to_open_file_dialog(self, instance):
        self.format_selector.dismiss()
        self.open_file_dialog.open()

    def is_qualtrics(self, instance):
        self.format_selector.dismiss()
        try:
            self.__controller.build_qresearch_report(self.__open_filename, "QUALTRICS")
            self.save_file_prompt.open()
        except Exception as inst:
            string = "Issue formatting report:/n %s" % str(inst)
            self.error_message(string)

    def is_y2(self, instance):
        self.format_selector.dismiss()
        try:
            self.__controller.build_qresearch_report(self.__open_filename, "Y2")
            self.save_file_prompt.open()
        except Exception as inst:
            string = "Issue formatting report:/n %s" % str(inst)
            self.error_message(string)

    def is_whatsapp(self, instance):
        self.format_selector.dismiss()
        try:
            self.__controller.build_qresearch_report(self.__open_filename, "WHATSAPP")
            self.save_file_prompt.open()
        except Exception as inst:
            string = "Issue formatting report:/n %s" % str(inst)
            self.error_message(string)

    def is_fb(self, instance):
        self.format_selector.dismiss()
        try:
            self.__controller.build_qresearch_report(self.__open_filename, "FACEBOOK")
            self.save_file_prompt.open()
        except Exception as inst:
            string = "Issue formatting report:/n %s" % str(inst)
            self.error_message(string)

    def save_file_prompt_to_selector(self, instance):
        self.save_file_prompt.dismiss()
        self.format_selector.open()

    def save_dialog_to_prompt(self, instance):
        self.save_file_dialog.dismiss()
        self.save_file_prompt.open()

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_prompt.dismiss()
        self.save_file_dialog.open()

    def finish(self):
        self.__controller.save_qresearch_report(self.__save_filename)
        self.save_file_dialog.dismiss()

    def error_message(self, error):
        label = Label(text=error)

        popup = Popup(title="Something Went Wrong",
        content=label,
        size_hint=(.5, .8), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        popup.open()
