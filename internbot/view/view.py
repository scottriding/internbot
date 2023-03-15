from view import topline_view
from view import appendix_view
from view import crosstabs_view

import tkinter as tk
from tkinter import ttk

class View(ttk.Frame):
	def __init__(self, parent):
		super().__init__(parent)

		self.parent = parent
        
		# set the controller
		self.controller = None

		# create widgets
		# input frequencies file section
		self.top_button = ttk.Button(self, text='Topline', command=self.run_topline)
		self.top_button.grid(row=1, column=0, padx=10)

		self.appendix_button = ttk.Button(self, text='Appendix', command=self.run_appendix)
		self.appendix_button.grid(row=1, column=1, padx=10)

		self.xtabs_button = ttk.Button(self, text='Crosstabs', command=self.run_xtabs)
		self.xtabs_button.grid(row=1, column=2, padx=10)

	def set_controller(self, controller):
		"""
		Set the controller
		:param controller:
		:return:
		"""
		self.controller = controller

	def run_topline(self):
		# create a popup-window
		topline_window = tk.Toplevel()
		
		# create topline view and place it in window
		self.topline = topline_view.ToplineView(topline_window, self.controller)
		self.topline.grid(row=0, column=0, padx=10, pady=10)

	def show_topline_matcher(self, csv_questions, qsf_questions):
		self.topline.show_matcher(csv_questions, qsf_questions)
		
	def hide_topline_matcher(self):
		self.topline.hide_matcher()

	def topline_show_success(self, message):
		self.topline.show_loading_success(message)

	def topline_show_error(self, message):
		self.topline.show_loading_error(message)

	def run_appendix(self):
		# create a popup-window
		appendix_window = tk.Toplevel()
		
		# create appendix view and place it in window
		self.appendix = appendix_view.AppendixView(appendix_window, self.controller)
		self.appendix.grid(row=0, column=0, padx=10, pady=10)

	def appendix_show_success(self, message):
		self.appendix.show_loading_success(message)

	def appendix_show_error(self, message):
		self.appendix.show_loading_error(message)

	def run_xtabs(self):
		# create a popup-window
		xtabs_window = tk.Toplevel()
		
		# create appendix view and place it in window
		self.xtabs = crosstabs_view.CrosstabsView(xtabs_window, self.controller)
		self.xtabs.grid(row=0, column=0, padx=10, pady=10)

	def crosstabs_show_success(self, message):
		self.xtabs.show_loading_success(message)

	def crosstabs_show_error(self, message):
		self.xtabs.show_loading_error(message)







