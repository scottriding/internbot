from view import gui_tools

## outside modules
from plyer import filechooser as fc
from plyer import notification

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
from kivy.uix.spinner import Spinner
from kivy.uix.scrollview import ScrollView
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.core.window import Window
from kivy.uix.gridlayout import GridLayout
from kivy.properties import ObjectProperty
import webbrowser
import os
from datetime import date, time, datetime
from collections import OrderedDict

class DocumentView(BoxLayout):

    def __init__(self, **kwargs):
        super(DocumentView, self).__init__(**kwargs)

        self.__controller = None

        self.is_qsf = True
        self.__group_names = []
        self.__survey = None

        self.open_survey_prompt = self.create_open_survey_prompt()
        self.trended_selector = self.create_trended_selector()
        self.trended_count = self.create_trended_count()
        self.open_freq_prompt = self.create_open_freq_prompt()
        self.format_selector = self.create_format_selector()
        self.save_file_prompt = self.create_save_file_prompt()

    def create_open_survey_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose survey (.csv or .qsf) file\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of survey files[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/qrk41o9k3dc761d/AAAr3L7bk2GTOJEQBfU9m5F-a?dl=0")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)

        popup_layout.add_widget(label)

        save_btn = Button(text='Next', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.open_survey_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select survey file",
        content=popup_layout,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return popup

    def create_trended_selector(self):
        chooser = BoxLayout(orientation='vertical')

        help_text = "Does this report have grouped or trended frequencies?\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of trended reports[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/tfr7twrf5cmt7md/AADQx960X8E4BaR2Nr0aD_yca?dl=0")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)

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
                    self.error_message("2 or more groups make up a trended report")
                else:
                    self.create_trended_labels(count)
            except:
                self.error_message("Issue parsing group count")

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
        
    def create_matching_checker(self, survey, qsf_block, freq_block):
    	self.matching_checker = Popup(title='Check qsf to frequency file match', size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})	
    	
    	pop_up_layout = BoxLayout(orientation="vertical")

    	scrollable_layout = GridLayout(cols=1, row_force_default=True, row_default_height=44, spacing=10, size_hint_y=None)
    	scrollable_layout.bind(minimum_height=scrollable_layout.setter('height'))
    	
    	row = BoxLayout()
    	row.add_widget(Label(text="QSF"))
    	row.add_widget(Label(text="Automatched Freq"))
    	row.add_widget(Label(text="Self Reassigned"))
    	
    	scrollable_layout.add_widget(row)

    	change_log = []
    	def rematch_value(spinner, text):
    		root_name = str(spinner.parent.children[2].text)

    		old_question = qsf_block.find_question_by_name(root_name)
    		new_question = freq_block.find_question_by_name(str(text))

    		old_question.responses = new_question.responses

    		log = "QSF: %s\tAutomatched Freq: %s\tSelf Reassigned: %s\n" % (root_name, spinner.parent.children[1].text, str(text))
    		change_log.append(log)
            
    	def submit():
            self.__log_text += "Rematched frequences: \n"
            for text in change_log:
                self.__log_text += text
            self.finish(qsf_block)
    	
    	freq_questions = []
    	for question in freq_block.questions:
    		freq_questions.append(question.name)

    	for question in qsf_block.questions:
    		# on the left is the qsf question name
    		# matched in the middle
    		# on the right is the matched frequency name
    		# button
            print(question.assigned)
            if question.parent != "CompositeQuestion":
                row = BoxLayout()
                r1_values = []
                r2_values = []
                for response in question.responses:
                    text = "%s - %s" % (response.value, response.label)
                    r1_values.append(text)
                    r2_values.append(response.assigned)
			
                qsf_spin = Spinner(text=question.name,
    								values=r1_values,
    								pos_hint={'center_x': .5, 'center_y': .5})
    		
                row.add_widget(qsf_spin)

                match_spin = Spinner(text=str(question.assigned),
    								values=r2_values,
    								pos_hint={'center_x': .5, 'center_y': .5})
    		
                row.add_widget(match_spin)
    			
                csv_spin = Spinner(text="rematch",
    								values=freq_questions,
    								pos_hint={'center_x': .5, 'center_y': .5})
                csv_spin.bind(text=rematch_value)
    		
                row.add_widget(csv_spin)
                scrollable_layout.add_widget(row)
            else:
                for subquestion in question.questions:
                    row = BoxLayout()
                    r1_values = []
                    r2_values = []
                    for response in subquestion.responses:
                        text = "%s - %s" % (response.value, response.label)
                        r1_values.append(text)
                        r2_values.append(response.assigned)
			
                    qsf_spin = Spinner(text=subquestion.name,
    									values=r1_values,
    									pos_hint={'center_x': .5, 'center_y': .5})
    		
                    row.add_widget(qsf_spin)
                    match_spin = Spinner(text=subquestion.assigned,
    									values=r2_values,
    									pos_hint={'center_x': .5, 'center_y': .5})
                    row.add_widget(match_spin)
    				
                    csv_spin = Spinner(text="rematch",
    									values=freq_questions,
    									pos_hint={'center_x': .5, 'center_y': .5})
                    csv_spin.bind(text=rematch_value)
    		
                    row.add_widget(csv_spin)
                    scrollable_layout.add_widget(row)
    	
    	scrollable = ScrollView(size=(self.matching_checker.width, self.matching_checker.height))
    	scrollable.add_widget(scrollable_layout)
    	
    	enter_btn = Button(text="Enter", size_hint=(.2, .2), pos_hint={'center_x': 0.5, 'center_y': 0.5}, 
        on_press=lambda x: submit())
    	
    	pop_up_layout.add_widget(scrollable)
    	pop_up_layout.add_widget(enter_btn)
    	self.matching_checker.content = pop_up_layout

    	self.matching_checker.open()

    def create_open_freq_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        help_text = "Choose a frequencies (.csv) file\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of frequency files[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/5jxxjp9fd3djfj5/AAAXE1qgqw3Jk2kefD4cElvIa?dl=0")

        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)

        popup_layout.add_widget(label)

        save_btn = Button(text='>', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.open_freq_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select frequencies file",
        content=popup_layout,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return popup

    def create_format_selector(self):
        chooser = BoxLayout(orientation='vertical')

        text = "Choose from the following format options."
        label = Label(text=text)
        label.font_family = "Y2"

        chooser.add_widget(label)

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)

        y2_btn = Button(text="Y2 Analytics", on_press=self.is_y2)
        oth_btn = Button(text="Other", on_press=self.is_other)

        button_layout.add_widget(y2_btn)
        button_layout.add_widget(oth_btn)

        chooser.add_widget(button_layout)

        format_chooser = Popup(title='Choose format',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return format_chooser 

    def create_save_file_prompt(self):
        popup_layout = BoxLayout(orientation='vertical')
        label = Label(text="Choose a file location and name for topline document report")

        popup_layout.add_widget(label)

        save_btn = Button(text='>', size_hint=(.2,.2))
        save_btn.pos_hint={'center_x': 0.5, 'center_y': 0.5}
        save_btn.bind(on_release=self.save_file_prompt_to_dialog)

        popup_layout.add_widget(save_btn)

        popup = Popup(title="Select save file location",
        content=popup_layout,
        size_hint=(.7, .5), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        return popup

    def run(self, controller):
        self.__template_name = ""
        self.__group_names = []
        self.__survey = None
        self.__template_file_path = None
        self.__survey_path = None
        self.__freq_path = None
        self.__controller = controller
        
        today = date.today()
        now = datetime.now()
        current_time = time(now.hour, now.minute, now.second)
        self.__log_text = "%s %s\n" % (today, current_time)
        self.__log_text += "Running topline document report...\n"
        
        self.open_survey_prompt.open()

    def open_survey_prompt_to_dialog(self, instance):
        self.open_survey_prompt.dismiss()
        survey_file = fc.open_file(title="Pick a survey file", 
                             filters=[("Comma-separated Values", "*.csv"), ("Qualtrics survey file", "*.qsf")])
        if survey_file is not None:
            path, ext = os.path.splitext(survey_file[0])
            if ext == ".csv":
                self.is_qsf = False
                self.__survey_path = survey_file[0]
                self.__log_text += "No QSF selected\n"
                self.__log_text += "Inputted survey file = %s\n" % (self.__survey_path)
                self.open_survey_dialog_to_trended_selector()
            elif ext == ".qsf":
                self.is_qsf = True
                self.__survey_path = survey_file[0]
                self.__log_text += "Inputted survey file = %s\n" % (self.__survey_path)
                self.open_survey_dialog_to_trended_selector()

    def open_survey_dialog_to_trended_selector(self):
        if self.is_qsf:
            try:
                self.__survey = self.__controller.build_survey(self.__survey_path)
                self.__log_text += "QSF successfully parsed by internbot.\n"
                self.trended_selector.open()
            except Exception as inst:
                string = "Error (%s):\n %s" % (type(inst), str(inst))
                self.__log_text += string
                self.error_message("Issue parsing .qsf file.")
        else:
            self.trended_selector.open()

    def trended_selector_to_count(self, instance):
        self.__log_text += "Trended report selected.\n"
        self.trended_selector.dismiss()
        self.trended_count.open()

    def trended_selector_to_freqs(self, instance):
        self.trended_selector.dismiss()
        if self.is_qsf:
            self.open_freq_prompt.open()
        else:
            self.format_selector.open()

    def trended_labels_to_freqs(self):
        self.__log_text += "Groups inputted: %s\n" % (self.__group_names)
        self.trended_labels.dismiss()
        if self.is_qsf:
            self.open_freq_prompt.open()
        else:
            self.format_selector.open()

    def open_freq_prompt_to_dialog(self, instance):
        self.open_freq_prompt.dismiss()
        freq_file = fc.open_file(title="Pick a frequency file", 
                             filters=[("Comma-separated Values", "*.csv")])
        if freq_file is not None:
            self.__freq_path = freq_file[0]
            self.__log_text += "Inputted frequency file = %s\n" % (self.__freq_path)
            self.format_selector.open()

    def is_y2(self, instance):
        self.__template_name = "Y2"
        self.__log_text += "Template selected: Y2\n"
        self.format_selector.dismiss()
        self.save_file_prompt.open()

    def is_other(self, instance):
        self.__template_name = "OTHER"
        self.__log_text += "Template selected: OTHER\n"
        self.format_selector.dismiss()

        template_file = fc.open_file(title="Pick a template file", 
                             filters=[("Word Document Template", "*.dotx"), ("Word Document", "*.docx")])

        if template_file is not None:
            self.__template_file_path = template_file[0]
            self.__log_text += "Inputted template file: %s\n" % (self.__template_file_path)
            self.save_file_prompt.open()

    def save_file_prompt_to_dialog(self, instance):
        self.save_file_prompt.dismiss()

        save_path = fc.save_file(title="Save report", 
                             filters=[("Word Document", "*.docx")])

        if save_path is not None:
            self.__save_filename = save_path[0]
            self.__log_text += "Inputted save filepath: %s\n" % (self.__save_filename)
            self.grab_questions()
			
    def grab_questions(self):
        if self.is_qsf:
            try:
                qsf_questions = self.__controller.build_document_model(self.__freq_path, self.__group_names, self.__survey)
                freq_questions = self.__controller.build_document_model(self.__freq_path, self.__group_names, None)
                self.__log_text += "Frequency file successfully read by internbot.\n"
                self.create_matching_checker(self.__survey, qsf_questions, freq_questions)
            except KeyError as key_error:
                string = "Misspelled or missing column (%s):\n %s" % (type(key_error), str(key_error))
                self.error_message(string)
            except Exception as inst:
                string = "Error (%s):\n %s" % (type(inst), str(inst))
                self.error_message(string)  
        else:
            try:
                questions = self.__controller.build_document_model(self.__survey_path, self.__group_names, None)
                self.__log_text += "Frequency file successfully read by internbot.\n"
                self.finish(questions)
            except KeyError as key_error:
                string = "Misspelled or missing column (%s):\n %s" % (type(key_error), str(key_error))
                self.error_message(string)
            except Exception as inst:
                string = "Error (%s):\n %s" % (type(inst), str(inst))
                self.error_message(string)

    def finish(self, questions):
        if self.is_qsf:
            self.matching_checker.dismiss()
        try:
            self.__log_text += "Building document..."
            self.__controller.build_document_report(questions, self.__template_name, self.__save_filename, self.__template_file_path)
            self.__log_text += "Success."

            today = date.today()
            now = datetime.now()
            current_time = time(now.hour, now.minute, now.second)
            log_filename = "internbot log %s %s.txt" % (today, current_time)
            log_dir = os.path.dirname(self.__save_filename)
            log_path = os.path.join(log_dir, log_filename)
            with open(log_path, 'w') as f:
                f.write(self.__log_text)
        except KeyError as key_error:
            string = "Misspelled or missing column (%s):\n %s" % (type(key_error), str(key_error))
            self.error_message(string)
        except Exception as inst:
            string = "Error (%s):\n %s" % (type(inst), str(inst))
            self.error_message(string)

    def error_message(self, error):
    	self.create_error_popup(error)

    def print_log(self, instance):
        self.error_chooser.dismiss()
        save_path = fc.save_file(title="Save log", 
                             filters=[("Text Document", "*.txt")])
                             
        if save_path is not None:
            with open(save_path[0], 'w') as f:
                f.write(self.__log_text)

    def dismiss_error(self, instance):
        self.error_chooser.dismiss()
        
    def create_error_popup(self, error):
        chooser = BoxLayout(orientation='vertical')

        label = Label(text=error)

        chooser.add_widget(label)

        button_layout = BoxLayout()
        button_layout.size_hint = (1, .1)
        yes_btn = Button(text="Print log", on_press=self.print_log)

        no_btn = Button(text="Ignore", on_press=self.dismiss_error)

        button_layout.add_widget(yes_btn)
        button_layout.add_widget(no_btn)

        chooser.add_widget(button_layout)

        self.error_chooser = Popup(title='Something went wrong',
        content=chooser,
        size_hint=(.9, .7 ), pos_hint={'center_x': 0.5, 'center_y': 0.5})

        self.error_chooser.open()
