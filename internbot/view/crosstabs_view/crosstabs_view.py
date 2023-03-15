import os
import fnmatch

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

class CrosstabsView(ttk.Frame):
	def __init__(self, parent, controller):
		super().__init__(parent)

		self.controller = controller

		# create widgets
		# unformatted report
		self.cross_label = ttk.Label(self, text='Path to unformatted report:', justify=tk.RIGHT, anchor="e")
		self.cross_label.grid(sticky = tk.E, row=1, column=0)
		
		self.cross_var = tk.StringVar()
		self.cross_entry = ttk.Entry(self, textvariable=self.cross_var, width=30)
		self.cross_entry.grid(row=1, column=1, sticky=tk.NSEW)
		
		self.cross_button = ttk.Button(self, text='Open', command=self.cross_button_clicked)
		self.cross_button.grid(row=1, column=2, padx=10)
		
		self.cross_message = ttk.Label(self, text='', foreground='red')
		self.cross_message.grid(row=1, column=3, sticky=tk.W)
		
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
		
		# formatting selection
		self.format_label = ttk.Label(self, text='Formatting template:', justify=tk.RIGHT, anchor="e")
		self.format_label.grid(sticky = tk.E, row=3, column=0)
		
		self.format_menu = tk.StringVar()
		self.format_menu.set("Pick a format:")
		
		formats = ["Y2", "Qualtrics", "Facebook"]
		
		self.format_dropdown = tk.OptionMenu(self, self.format_menu, *formats, command=self.format_menu_clicked)
		self.format_dropdown.grid(row=3, column=1)

		self.format_message = ttk.Label(self, text='', foreground='red')
		self.format_message.grid(row=3, column=3, sticky=tk.W)

		# submit		
		self.submit_button = ttk.Button(self, text='Submit', command=self.submit_button_clicked)
		self.submit_button.grid(row=4, column=2, pady=15, padx=10)

		self.submit_message = ttk.Label(self, text='', foreground='blue')
		self.submit_message.grid(row=4, column=1, sticky=tk.W)

	def show_loading_success(self, message):
		self.submit_message['foreground'] = 'blue'
		self.submit_message['text'] = message
		self.submit_message.after(3000, self.hide_submit_message)

	def hide_submit_message(self):
		self.submit_message['text'] = ''

	def show_loading_error(self, message):
		self.submit_message['foreground'] = 'red'
		self.submit_message['text'] = message

	def cross_button_clicked(self):
		"""
		Handle verbatims button click event
		:return:
		"""
		self.hide_submit_message()
		filename = filedialog.askopenfilename(title='Open unformatted report', initialdir='~/', filetypes=[('Microsoft Excel', '*.xlsx')])
		if filename:
			self.cross_entry['state'] = "normal"
			self.cross_entry.delete(0,tk.END)
			self.cross_entry.insert(0,filename)
			self.check_cross()

	def show_cross_error(self, message):
		"""
		Show an error message
		:param message:
		:return:
		"""
		self.cross_message['text'] = message
		self.cross_message['foreground'] = 'red'

	
	def hide_cross_error(self):
		"""
		hide an error message
		:param message:
		:return:
		"""
		self.cross_message['text'] = ''

	def output_button_clicked(self):
		"""
		Handle output button click event
		:return:
		"""
		self.hide_submit_message()
		filename = filedialog.asksaveasfilename(title='Save report', initialdir='~/', filetypes=[('Microsoft Excel', '*.xlsx')])
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
		self.output_message['text'] = message
		self.output_message['foreground'] = 'red'

	
	def hide_output_error(self):
		"""
		hide an error message
		:param message:
		:return:
		"""
		self.output_message['text'] = ''

	def format_menu_clicked(self, event):
		self.hide_submit_message()
		# if clicked at all - user forced to pick dropdown item
		# if error message on screen, remove
		self.format_message['text'] = ''
	
	def show_format_error(self, message):
		self.format_message['text'] = message
		self.format_message['foreground'] = 'red'
		
	def hide_format_error(self):
		self.format_message['text'] = ''

	def check_cross(self):
		# check if verbatims file path exists
		if not os.path.exists(self.cross_entry.get()):
			self.show_cross_error("This path does not exist.")
		else:
			self.hide_cross_error()
			
			# check the extension
			if not fnmatch.fnmatch(os.path.join(self.cross_entry.get()), "*.xlsx"):
				self.show_cross_error("Microsoft Excel (.xlsx) file only.")
			else:
				self.hide_cross_error()
				return True
		
		return False

	def check_output(self):
		# check if directory exists
		if not os.path.isdir(os.path.dirname(os.path.join(self.output_var.get()))):
			self.show_output_error("This path does not exist.")
		else:
			self.hide_output_error()
			
			# check the extension
			if not fnmatch.fnmatch(os.path.join(self.output_var.get()), "*.xlsx"):
				self.show_output_error("Microsoft Excel (.xlsx) file only.")
			else:
				self.hide_output_error
				return True
		
		return False

	def check_format(self):
		# check if format was selected
		if self.format_menu.get() == "Pick a format:":
			self.show_format_error("Pick a format template")
		else:
			self.hide_format_error()
			return True
		
		return False

	def submit_button_clicked(self):
		"""
		Handle submit button click event
		:return:
		"""
		self.hide_submit_message()
		
		# check all entry variables for valid paths/exts
		# we do global check on variables so all input errors can be shown at once
		self.check_cross()
		self.check_output()
		self.check_format()
		
		# ask if we're ready to be submitted
		# here we stack checks based on priority for report
		
		# for crosstabs, all input fields need to be filled out for report to run
		if self.check_cross():
			if self.check_output():
				if self.check_format():
					cross_path = os.path.join(self.cross_var.get())
					output_path = os.path.join(self.output_var.get())
					template = self.format_menu.get()
					
					# ask controller to run files
					self.controller.run_format_crosstabs(cross_path, output_path, template)




