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
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose an unformatted Amazon SPSS (.xlsx) crosstab report\n\n"
        help_text += "[ref=click][color=F3993D][u]Click here for examples of unformatted reports[/u][/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/1mu0eogzluyy8s1/AACXSWbtpTP5nuXEANE3ud0qa?dl=0")

        popup = Popup(title="",
        separator_height = 0,
        content=popup_layout,
        size_hint=(.7, .5), 
        pos_hint={'center_x': 0.5, 'center_y': 0.5},
        auto_dismiss=False)

        close_btn = Button(text='x', size_hint=(.08,.2))
        close_btn.bind(on_release=popup.dismiss)

        popup_layout.add_widget(close_btn)

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family= "Y2"

        popup_layout.add_widget(label)

        next_btn = Button(text='Next', size_hint=(.2,.2), pos_hint={'center_x': 0.5, 'center_y': 0.5})
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
        pos_hint={'center_x': 0.5, 'center_y': 0.5},
        auto_dismiss=False)

        button_layout = BoxLayout(orientation='vertical')
        button_layout.size_hint = (.1, .2)

        close_btn = Button(text='x')
        close_btn.bind(on_release=file_chooser.dismiss)

        back_btn = Button(text='<')
        back_btn.bind(on_release=self.open_dialog_to_prompt)

        button_layout.add_widget(close_btn)
        button_layout.add_widget(back_btn)

        chooser_layout.add_widget(button_layout)

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                path, ext = os.path.splitext(filepath)
                self.__open_filename = filepath
                self.open_file_dialog_to_toc_prompt()
            except IndexError:
                self.error_message("Please pick an Amazon report (.xlsx) file")

        chooser_view = FileChooserListView()
        chooser_view.bind(on_selection=lambda x: chooser_view.selection)
        chooser_view.path = os.path.expanduser("~")
        chooser_view.filters = ["*.xlsx"]

        open_btn = Button(text='open', size_hint=(.2,.1), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        open_btn.bind(on_release=lambda x: open_file(chooser_view.path, chooser_view.selection))

        container.add_widget(chooser_view)
        container.add_widget(open_btn)
        chooser_layout.add_widget(container)

        return file_chooser 

    def create_toc_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose table of contents (.csv) crosstab file\n\n"
        help_text += "[ref=click][color=F3993D][u]Click here for examples of table of content files[/u][/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/md12sy6blwc5rzo/AACNcJMpFxKhBstbevxlAsZja?dl=0")

        popup = Popup(title="",
        separator_height = 0,
        content=popup_layout,
        size_hint=(.7, .5), 
        pos_hint={'center_x': 0.5, 'center_y': 0.5},
        auto_dismiss=False)

        button_layout = BoxLayout(orientation='vertical')
        button_layout.size_hint = (.15, .3)

        close_btn = Button(text='x')
        close_btn.bind(on_release=popup.dismiss)

        back_btn = Button(text='<')
        back_btn.bind(on_release=self.toc_prompt_to_open_dialog)

        button_layout.add_widget(close_btn)
        button_layout.add_widget(back_btn)

        popup_layout.add_widget(button_layout)

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)

        popup_layout.add_widget(label)

        save_btn = Button(text='Next', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.toc_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        return popup

    def create_toc_dialog(self):
        chooser_layout = BoxLayout(orientation='vertical')
        container = BoxLayout(orientation='vertical')

        file_chooser = Popup(title='',
        separator_height = 0,
        content=chooser_layout,
        size_hint=(.9, .7), 
        pos_hint={'center_x': 0.5, 'center_y': 0.5},
        auto_dismiss=False)

        button_layout = BoxLayout(orientation='vertical')
        button_layout.size_hint = (.1, .2)

        close_btn = Button(text='x')
        close_btn.bind(on_release=file_chooser.dismiss)

        back_btn = Button(text='<')
        back_btn.bind(on_release=self.toc_dialog_to_prompt)

        button_layout.add_widget(close_btn)
        button_layout.add_widget(back_btn)

        chooser_layout.add_widget(button_layout)

        def open_file(path, filename):
            try:
                filepath = os.path.join(path, filename[0])
                self.__toc_filename = filepath
                self.toc_dialog_to_trended_selector()
            except IndexError:
                self.error_message("Please pick a table of contents (.csv) file")

        chooser_view = FileChooserListView()
        chooser_view.path = os.path.expanduser("~")
        chooser_view.filters = ["*.csv"]
        chooser_view.bind(on_selection=lambda x: chooser_view.selection)

        open_btn = Button(text='open', size_hint=(.2,.1), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        open_btn.bind(on_release=lambda x: open_file(chooser_view.path, chooser_view.selection))

        container.add_widget(chooser_view)
        container.add_widget(open_btn)
        chooser_layout.add_widget(container)

        return file_chooser 

    def create_trended_selector(self):
        chooser = BoxLayout(orientation='vertical')

        help_text = "Does this report have grouped or trended banners?\n\n"
        help_text += "[ref=click][color=F3993D][u]Click here for examples of trended banners[/u][/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/5vnjdqh6rji5w78/AAC0T3o-UgPNEfWalmP8PrIYa?dl=0")

        popup = Popup(title="",
        separator_height = 0,
        content=chooser,
        size_hint=(.9, .7), 
        pos_hint={'center_x': 0.5, 'center_y': 0.5},
        auto_dismiss=False)

        button_layout = BoxLayout(orientation='vertical')
        button_layout.size_hint = (.1, .2)

        close_btn = Button(text='x')
        close_btn.bind(on_release=popup.dismiss)

        back_btn = Button(text='<')
        back_btn.bind(on_release=self.trended_selector_to_toc_dialog)

        button_layout.add_widget(close_btn)
        button_layout.add_widget(back_btn)

        chooser.add_widget(button_layout)

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)

        chooser.add_widget(label)

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)
        yes_btn = Button(text="Yes", on_press=self.is_trended)

        no_btn = Button(text="No", on_press=self.is_basic)

        button_layout.add_widget(yes_btn)
        button_layout.add_widget(no_btn)

        chooser.add_widget(button_layout)

        return popup

    def create_save_file_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        label = Label(text="Choose a file location and name for Amazon crosstabs report")
        label.font_family= "Y2"

        popup = Popup(title="",
        separator_height = 0,
        content=popup_layout,
        size_hint=(.7, .5), 
        pos_hint={'center_x': 0.5, 'center_y': 0.5},
        auto_dismiss=False)

        button_layout = BoxLayout(orientation='vertical')
        button_layout.size_hint = (.15, .3)

        close_btn = Button(text='x')
        close_btn.bind(on_release=popup.dismiss)

        back_btn = Button(text='<')
        back_btn.bind(on_release=self.save_file_prompt_to_selector)

        button_layout.add_widget(close_btn)
        button_layout.add_widget(back_btn)
        popup_layout.add_widget(button_layout)

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
        pos_hint={'center_x': 0.5, 'center_y': 0.5},
        auto_dismiss=False)

        menu_layout = BoxLayout(orientation='vertical')
        menu_layout.size_hint = (.1, .2)

        close_btn = Button(text='x')
        close_btn.bind(on_release=file_chooser.dismiss)

        back_btn = Button(text='<')
        back_btn.bind(on_release=self.save_dialog_to_prompt)

        menu_layout.add_widget(close_btn)
        menu_layout.add_widget(back_btn)

        chooser_layout.add_widget(menu_layout)

        chooser_view = FileChooserIconView()
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
        self.__is_trended = False
        self.__controller = controller
        self.open_file_prompt.open()

    def open_file_prompt_to_dialog(self, instance):
        self.open_file_prompt.dismiss()
        self.open_file_dialog.open()

    def open_dialog_to_prompt(self, instance):
        self.open_file_dialog.dismiss()
        self.open_file_prompt.open()

    def open_file_dialog_to_toc_prompt(self):
        self.open_file_dialog.dismiss()
        self.open_toc_prompt.open()

    def toc_prompt_to_open_dialog(self, instance):
        self.open_toc_prompt.dismiss()
        self.open_file_dialog.open()

    def toc_prompt_to_dialog(self, instance):
        self.open_toc_prompt.dismiss()
        self.open_toc_dialog.open()

    def toc_dialog_to_prompt(self, instance):
        self.open_toc_dialog.dismiss()
        self.open_toc_prompt.open()

    def toc_dialog_to_trended_selector(self):
        self.open_toc_dialog.dismiss()
        self.trended_selector.open()

    def trended_selector_to_toc_dialog(self, instance):
        self.trended_selector.dismiss()
        self.open_toc_dialog.open()

    def toc_dialog_to_selector(self, instance):
        self.open_toc_dialog.dismiss()
        self.trended_selector.open()

    def is_trended(self, instance):
        self.__is_trended = True
        self.trended_selector.dismiss()
        self.save_file_prompt.open()

    def is_basic(self, instance):
        self.__is_trended = False
        self.trended_selector.dismiss()
        self.save_file_prompt.open()

    def save_file_prompt_to_selector(self, instance):
        self.save_file_prompt.dismiss()
        self.trended_selector.open()

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_prompt.dismiss()
        self.save_file_dialog.open()

    def save_dialog_to_prompt(self, instance):
        self.save_file_dialog.dismiss()
        self.save_file_prompt.open()

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
