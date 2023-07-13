import os
import fnmatch

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

class ScoresToplineView(ttk.Frame):
	def __init__(self, parent, controller):
		super().__init__(parent)
		self.parent = parent
		self.controller = controller

		# create widgets
		# path_to_csv, round, location
		# input file path
		self.input_label = ttk.Label(self, text='Path to input file:', justify=tk.RIGHT, anchor="e")
		self.input_label.grid(sticky = tk.E, row=1, column=0)
		
		self.input_var = tk.StringVar()
		self.input_entry = ttk.Entry(self, textvariable=self.input_var, width=30)
		self.input_entry.grid(row=1, column=1, sticky=tk.NSEW)
		
		self.input_button = ttk.Button(self, text='Open', command=self.input_button_clicked)
		self.input_button.grid(row=1, column=2, padx=10)
		
		self.input_message = ttk.Label(self, text='', foreground='red')
		self.input_message.grid(row=1, column=3, sticky=tk.W)

		# round selection
		self.round_label = ttk.Label(self, text='Round number:', justify=tk.RIGHT, anchor="e")
		self.round_label.grid(sticky = tk.E, row=2, column=0)
		
		self.round_menu = tk.StringVar()
		self.round_menu.set("Pick a round number:")
		
		rounds = list(range(1,51))
		
		self.round_dropdown = tk.OptionMenu(self, self.round_menu, *rounds, command=self.round_menu_clicked)
		self.round_dropdown.grid(row=2, column=1, sticky=tk.E)
		
		self.round_message = ttk.Label(self, text='', foreground='red')
		self.round_message.grid(row=2, column=3, sticky=tk.W)
		
		# location
		self.state_dict = { "Alabama": "AL", "Alaska": "AK", "Arizona": "AZ", "Arkansas": "AR", "California": "CA", "Colorado": "CO", "Connecticut": "CT", "Delaware": "DE", "District of Columbia": "DC", "Florida": "FL", "Georgia": "GA", "Hawaii": "HI", "Idaho": "ID", "Illinois": "VI", "Indiana": "WA", "Iowa": "WV", "Kansas": "KS", "Kentucky": "KY", "Louisiana": "LA", "Maine": "ME", "Maryland": "MD", "Massachusetts": "MA", "Michigan": "MI", "Minnesota": "MN", "Mississippi": "MS", "Missouri": "MO", "Montana": "MT", "Nebraska": "NE", "Nevada": "NV", "New Hampshire": "NH", "New Jersey": "NJ", "New Mexico": "NM", "New York": "IL", "North Carolina": "IN", "North Dakota": "ND", "Ohio": "OH", "Oklahoma": "OK", "Oregon": "OR", "Pennsylvania": "PA", "Rhode Island": "RI", "South Carolina": "SC", "South Dakota": "SD", "Tennessee": "TN", "Texas": "TX", "Utah": "UT", "Vermont": "VT", "Virginia": "VA", "Washington": "NY", "West Virginia": "NC", "Wisconsin": "WI", "Wyoming": "WY"}
		
		self.location_label = ttk.Label(self, text='Location:', justify=tk.RIGHT, anchor="e")
		self.location_label.grid(sticky=tk.E, row=3, column=0)
		
		self.location_menu = tk.StringVar()
		self.location_menu.set("Pick a location:")
		
		self.location_dropdown = tk.OptionMenu(self, self.location_menu, *self.state_dict.keys(), command=self.location_menu_clicked)
		self.location_dropdown.grid(row=3, column=1, sticky=tk.E)

		self.location_message = ttk.Label(self, text='', foreground='red')
		self.location_message.grid(row=3, column=3, sticky=tk.W)
		
		# output
		self.output_label = ttk.Label(self, text='Save as:', justify=tk.RIGHT, anchor="e")
		self.output_label.grid(sticky = tk.E, row=4, column=0)
		
		self.output_var = tk.StringVar()
		self.output_entry = ttk.Entry(self, textvariable=self.output_var, width=30)
		self.output_entry.grid(row=4, column=1, sticky=tk.NSEW)
		
		self.output_button = ttk.Button(self, text='Save as', command=self.output_button_clicked)
		self.output_button.grid(row=4, column=2, padx=10)
		
		self.output_message = ttk.Label(self, text='', foreground='red')
		self.output_message.grid(row=4, column=3, sticky=tk.W)

		# submit		
		self.submit_button = ttk.Button(self, text='Submit', command=self.submit_button_clicked)
		self.submit_button.grid(row=5, column=2, pady=15, padx=10)
		
		self.submit_message = ttk.Label(self, text='', foreground='blue')
		self.submit_message.grid(row=5, column=1, sticky=tk.W)

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

	def input_button_clicked(self):
		"""
		Handle input button click event
		:return:
		"""
		filename = filedialog.askopenfilename(title='Open an input file', initialdir='~/', filetypes=[('Comma-Separated File', '*.csv')])
		if filename:
			self.input_entry['state'] = "normal"
			self.input_entry.delete(0,tk.END)
			self.input_entry.insert(0,filename)
			self.check_input()

	def show_input_error(self, message):
		"""
		Show an error message
		:param message:
		:return:
		"""
		self.update_idletasks()
		self.input_message['text'] = message
		self.input_message['foreground'] = 'red'

	
	def hide_input_error(self):
		"""
		hide an error message
		:param message:
		:return:
		"""
		self.input_message['text'] = ''
		
	def round_menu_clicked(self, event):
		# if clicked at all - user forced to pick dropdown item
		# if error message on screen, remove
		self.round_message['text'] = ''
	
	def show_round_error(self, message):
		self.update_idletasks()
		self.round_message['text'] = message
		self.round_message['foreground'] = 'red'
		
	def hide_round_error(self):
		self.round_message['text'] = ''

	def check_round(self):
		# check if round was selected
		if self.round_menu.get() == "Pick a round number:":
			self.show_round_error("Please select a number.")
		else:
			self.hide_round_error()
			return True
		
		return False

	def location_menu_clicked(self, event):
		# if clicked at all - user forced to pick dropdown item
		# if error message on screen, remove
		self.location_message['text'] = ''
	
	def show_location_error(self, message):
		self.update_idletasks()
		self.location_message['text'] = message
		self.location_message['foreground'] = 'red'
		
	def hide_location_error(self):
		self.location_message['text'] = ''

	def check_location(self):
		# check if location was selected
		if self.location_menu.get() == "Pick a location:":
			self.show_location_error("Please select a location.")
		else:
			self.hide_location_error()
			return True
		
		return False

	def output_button_clicked(self):
		"""
		Handle output button click event
		:return:
		"""
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

	def check_input(self):
		# check if input file path exists
		if not os.path.exists(self.input_entry.get()):
			self.show_input_error("This path does not exist.")
		else:
			self.hide_input_error()
			
			# check the extension
			if not fnmatch.fnmatch(os.path.join(self.input_entry.get()), "*.csv"):
				self.show_input_error("Comma-separated files (.csv) only.")
			else:
				self.hide_input_error()
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
				self.show_output_error("Microsoft Excel (.xlsx) only.")
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
		self.check_input()
		self.check_round()
		self.check_location()
		self.check_output()
		
		# ask if we're ready to be submitted
		# here we stack checks based on priority for report
		
		# for appendix, all input fields need to be filled out for report to run
		if self.check_input():
			if self.check_round():
				if self.check_location():
					if self.check_output():
				 		input_path = os.path.join(self.input_entry.get())
				 		input_round = int(self.round_menu.get())
				 		input_location = self.state_dict.get(self.location_menu.get())
				 		output_path = os.path.join(self.output_entry.get())
				 		# ask controller to run files
				 		self.controller.run_scores_topline(input_path, input_round, input_location, output_path)