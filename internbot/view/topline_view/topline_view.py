from view.topline_view import frequency_matcher

import os
import fnmatch

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

class ToplineView(ttk.Frame):

	def __init__(self, parent, controller):
		super().__init__(parent)

		self.controller = controller
		self.parent = parent

		# create widgets		
		# input frequencies file section
		self.freq_label = ttk.Label(self, text='Path to frequencies file:', justify=tk.RIGHT, anchor="e")
		self.freq_label.grid(sticky = tk.E, row=1, column=0)

		self.freq_var = tk.StringVar()
		self.freq_entry = ttk.Entry(self, textvariable=self.freq_var, width=30)
		self.freq_entry.grid(row=1, column=1, sticky=tk.NSEW)

		self.freq_button = ttk.Button(self, text='Open', command=self.freq_button_clicked)
		self.freq_button.grid(row=1, column=2, padx=10)

		self.freq_message_label = ttk.Label(self, text='', foreground='red')
		self.freq_message_label.grid(row=1, column=3, sticky=tk.W)

		# input output file section
		self.output_label = ttk.Label(self, text='Save as:', justify=tk.RIGHT, anchor="e")
		self.output_label.grid(sticky = tk.E, row=2, column=0)
		
		self.output_var = tk.StringVar()
		self.output_entry = ttk.Entry(self, textvariable=self.output_var, width=30)
		self.output_entry.grid(row=2, column=1, sticky=tk.NSEW)
		
		self.output_button = ttk.Button(self, text='Save as', command=self.output_button_clicked)
		self.output_button.grid(row=2, column=2, padx=10)

		self.output_message_label = ttk.Label(self, text='', foreground='red')
		self.output_message_label.grid(row=2, column=3, sticky=tk.W)
		
		# input survey file section
		self.survey_label = ttk.Label(self, text='Path to survey file (OPTIONAL):', justify=tk.RIGHT, anchor="e")
		self.survey_label.grid(sticky = tk.E, row=3, column=0)
		
		self.survey_var = tk.StringVar()
		self.survey_entry = ttk.Entry(self, textvariable=self.survey_var, width=30)
		self.survey_entry.grid(row=3, column=1, sticky=tk.NSEW)
		
		self.survey_button = ttk.Button(self, text='Open', command=self.survey_button_clicked)
		self.survey_button.grid(row=3, column=2, padx=10)

		self.survey_message_label = ttk.Label(self, text='', foreground='red')
		self.survey_message_label.grid(row=3, column=3, sticky=tk.W)
		
		# input groups section
		self.group_label = ttk.Label(self, text='Group names (OPTIONAL):', justify=tk.RIGHT, anchor="e")
		self.group_label.grid(sticky = tk.E, row=4, column=0)

		self.group_var = tk.StringVar()
		self.group_entry = ttk.Entry(self, textvariable=self.group_var, width=30)
		self.group_entry.grid(row=4, column=1, sticky=tk.NSEW)


		self.group_message_label = ttk.Label(self, text='', foreground='red')
		self.group_message_label.grid(row=4, column=2, sticky=tk.W)

		# submit		
		self.submit_button = ttk.Button(self, text='Submit', command=self.submit_button_clicked)
		self.submit_button.grid(row=6, column=2, pady=15, padx=10)

		self.submit_message = ttk.Label(self, text='', foreground='blue')
		self.submit_message.grid(row=6, column=1, sticky=tk.W)

	def show_matcher(self, csv_questions, qsf_questions, groups):
		# create a popup-window
		self.matcher_window = tk.Toplevel()
		
		# create frequency matcher view and place it in window
		matcher = frequency_matcher.Matcher(self.matcher_window, csv_questions, qsf_questions, groups, self.controller)
		matcher.pack(side="top", fill="both", expand=True)

	def hide_matcher(self):
		self.matcher_window.destroy()

	def show_loading_success(self, message):
		self.update_idletasks()
		self.submit_message['foreground'] = 'blue'
		self.submit_message['text'] = message
		self.submit_message.after(1000, self.hide_submit_message)

	def hide_submit_message(self):
		self.submit_message['text'] = ''
		self.parent.destroy()

	def show_loading_error(self, message):
		self.update_idletasks()
		self.submit_message['foreground'] = 'red'
		self.submit_message['text'] = message
		
	def freq_button_clicked(self):
		"""
		Handle frequency button click event
		:return:
		"""
		filename = filedialog.askopenfilename(title='Open a frequencies file', initialdir='~/', filetypes=[('Comma-Separated File', '*.csv')])
		if filename:
			self.freq_entry['state'] = "normal"
			self.freq_entry.delete(0,tk.END)
			self.freq_entry.insert(0,filename)
			self.check_freq()

	def show_freq_error(self, message):
		"""
		Show an error message
		:param message:
		:return:
		"""
		self.update_idletasks()
		self.freq_message_label['text'] = message
		self.freq_message_label['foreground'] = 'red'

	
	def hide_freq_error(self):
		"""
		hide an error message
		:param message:
		:return:
		"""
		self.freq_message_label['text'] = ''

	def output_button_clicked(self):
		"""
		Handle output button click event
		:return:
		"""
		filename = filedialog.asksaveasfilename(title='Save report', initialdir='~/', filetypes=[('Microsoft Word Document', '*.docx')])
		if filename:
			self.output_entry['state'] = "normal"
			self.output_entry.delete(0,tk.END)
			self.output_entry.insert(0,filename)
			self.check_output()

	def show_output_error(self, message):
		"""
		Show an error message
		:param message:
		:return:
		"""
		self.update_idletasks()
		self.output_message_label['text'] = message
		self.output_message_label['foreground'] = 'red'

	
	def hide_output_error(self):
		"""
		hide an error message
		:param message:
		:return:
		"""
		self.output_message_label['text'] = ''

	def survey_button_clicked(self):
		"""
		Handle survey button click event
		:return:
		"""
		filename = filedialog.askopenfilename(title='Open a survey file', initialdir='~/', filetypes=[('Qualtrics Survey File', '*.qsf')])
		if filename:
			self.survey_entry['state'] = "normal"
			self.survey_entry.delete(0, tk.END)
			self.survey_entry.insert(0, filename)
			self.check_survey()

	def show_survey_error(self, message):
		"""
		Show an error message
		:param message:
		:return:
		"""
		self.update_idletasks()
		self.survey_message_label['text'] = message
		self.survey_message_label['foreground'] = 'red'

	
	def hide_survey_error(self):
		"""
		hide an error message
		:param message:
		:return:
		"""
		self.survey_message_label['text'] = ''

	def show_group_feedback(self, message):
		"""
		Show an group error message
		:param message:
		:return:
		"""
		self.update_idletasks()
		self.group_message_label['text'] = message
		self.group_message_label['foreground'] = 'green'

	
	def hide_group_feedback(self):
		"""
		hide a group error message
		:param message:
		:return:
		"""
		self.group_message_label['text'] = ''
		
	def check_freq(self):
		# check if frequency file path exists
		if not os.path.exists(self.freq_entry.get()):
			self.show_freq_error("This path does not exist.")
		else:
			self.hide_freq_error()
			
			# check the extension
			if not fnmatch.fnmatch(os.path.join(self.freq_entry.get()), "*.csv"):
				self.show_freq_error("Comma-separated files (.csv) only.")
			else:
				self.hide_freq_error()
				return True
		
		return False

	def check_output(self):
		# check if directory exists
		if not os.path.isdir(os.path.dirname(os.path.join(self.output_var.get()))):
			self.show_output_error("This path does not exist.")
		else:
			self.hide_output_error()
			
			# check the extension
			if not fnmatch.fnmatch(os.path.join(self.output_var.get()), "*.docx"):
				self.show_output_error("Microsoft Word documents (.docx) only.")
			else:
				self.hide_output_error
				return True
		
		return False
	
	def check_survey(self):
		# this input is optional - so check if there's a path
		if self.survey_var.get() != "":
			# a path was submitted - check that it exists
			if not os.path.exists(os.path.join(self.survey_var.get())):
				self.show_survey_error("This path does not exist.")
				return False
			else:
				self.hide_survey_error()
				
				# check the extension
				if not fnmatch.fnmatch(os.path.join(self.survey_var.get()), "*.qsf"):
					self.show_survey_error("Qualtrics survey files (.qsf) only.")
					return False
				else:
					self.hide_survey_error()
					return True
		else:
			self.hide_survey_error()

	def check_group(self):
		# this input is optional - so check if there are names
		if self.group_var.get() != "":
			str_input = self.group_var.get().split(",")
			str_input = [i.strip() for i in str_input]
			message = "Groups: "
			for group in str_input:
				message += str(group + " ")
			self.show_group_feedback(message)
			return(str_input)
		else:
			self.hide_group_feedback()
	
	def submit_button_clicked(self):
		"""
		Handle submit button click event
		:return:
		"""
		# check all entry variables for valid paths/exts
		# we do global check on variables so all input errors can be shown at once
		self.check_freq()
		self.check_output()
		self.check_survey()
		self.check_group()

		# ask if we're ready to be submitted
		# here we stack checks based on priority for report
		
		# for topline, frequency and output are required fields
		# survey/groups are optional   	
		if self.check_freq():
			if self.check_output():
				frequency_path = os.path.join(self.freq_var.get())
				output_path = os.path.join(self.output_var.get())
				# this is ready to go to controller at this point
				# check for the optional inputs
				
				if self.check_survey() is False:
					# bad input for survey file
					pass
				else:
					if self.survey_var.get() != "":
						survey_path = os.path.join(self.survey_var.get())
					else:
						survey_path = None
					groups = self.check_group()
					if groups is None:
						groups = []

					# ask controller to run files					
					self.controller.run_topline(frequency_path, output_path, survey_path, groups)
