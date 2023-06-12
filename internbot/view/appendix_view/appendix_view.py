import os
import fnmatch

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

class AppendixView(ttk.Frame):
	def __init__(self, parent, controller):
		super().__init__(parent)

		# set the controller
		self.controller = controller
		self.parent = parent
		
		# create widgets
		# verbatims
		self.verbatims_label = ttk.Label(self, text='Path to verbatims file:', justify=tk.RIGHT, anchor="e")
		self.verbatims_label.grid(sticky = tk.E, row=1, column=0)
		
		self.verbatims_var = tk.StringVar()
		self.verbatims_entry = ttk.Entry(self, textvariable=self.verbatims_var, width=30)
		self.verbatims_entry.grid(row=1, column=1, sticky=tk.NSEW)
		
		self.verbatims_button = ttk.Button(self, text='Open', command=self.verbatims_button_clicked)
		self.verbatims_button.grid(row=1, column=2, padx=10)
		
		self.verbatims_message = ttk.Label(self, text='', foreground='red')
		self.verbatims_message.grid(row=1, column=3, sticky=tk.W)
		
		# output
		self.output_label = ttk.Label(self, text='Save as:', justify=tk.RIGHT, anchor="e")
		self.output_label.grid(sticky = tk.E, row=2, column=0)
		
		self.output_var = tk.StringVar()
		self.output_entry = ttk.Entry(self, textvariable=self.output_var, width=30)
		self.output_entry.grid(row=2, column=1, sticky=tk.NSEW)
		
		self.output_button = ttk.Button(self, text='Save as', command=self.output_button_clicked)
		self.output_button.grid(row=2, column=2, padx=10)
		
		self.output_message = ttk.Label(self, text='', foreground='red')
		self.output_message.grid(row=2, column=3, sticky=tk.W)

		# submit		
		self.submit_button = ttk.Button(self, text='Submit', command=self.submit_button_clicked)
		self.submit_button.grid(row=3, column=2, pady=15, padx=10)
		
		self.submit_message = ttk.Label(self, text='', foreground='blue')
		self.submit_message.grid(row=3, column=1, sticky=tk.W)

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


	def verbatims_button_clicked(self):
		"""
		Handle verbatims button click event
		:return:
		"""
		filename = filedialog.askopenfilename(title='Open a verbatims file', initialdir='~/', filetypes=[('Comma-Separated File', '*.csv')])
		if filename:
			self.verbatims_entry['state'] = "normal"
			self.verbatims_entry.delete(0,tk.END)
			self.verbatims_entry.insert(0,filename)
			self.check_verbatims()

	def show_verbatims_error(self, message):
		"""
		Show an error message
		:param message:
		:return:
		"""
		self.update_idletasks()
		self.verbatims_message['text'] = message
		self.verbatims_message['foreground'] = 'red'

	
	def hide_verbatims_error(self):
		"""
		hide an error message
		:param message:
		:return:
		"""
		self.verbatims_message['text'] = ''

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
		self.output_message['text'] = message
		self.output_message['foreground'] = 'red'

	
	def hide_output_error(self):
		"""
		hide an error message
		:param message:
		:return:
		"""
		self.output_message['text'] = ''

	def check_verbatims(self):
		# check if verbatims file path exists
		if not os.path.exists(self.verbatims_entry.get()):
			self.show_verbatims_error("This path does not exist.")
		else:
			self.hide_verbatims_error()
			
			# check the extension
			if not fnmatch.fnmatch(os.path.join(self.verbatims_entry.get()), "*.csv"):
				self.show_verbatims_error("Comma-separated files (.csv) only.")
			else:
				self.hide_verbatims_error()
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

	def submit_button_clicked(self):
		"""
		Handle submit button click event
		:return:
		"""
		# check all entry variables for valid paths/exts
		# we do global check on variables so all input errors can be shown at once
		self.check_verbatims()
		self.check_output()
		
		# ask if we're ready to be submitted
		# here we stack checks based on priority for report
		
		# for appendix, all input fields need to be filled out for report to run
		if self.check_verbatims():
			if self.check_output():
				 verbatims_path = os.path.join(self.verbatims_entry.get())
				 output_path = os.path.join(self.output_entry.get())
				 # ask controller to run files
				 self.controller.run_appendix(verbatims_path, output_path)





