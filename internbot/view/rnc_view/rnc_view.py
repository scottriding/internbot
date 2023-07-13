from view.rnc_view import issue_trended_view
from view.rnc_view import scores_topline_view
from view.rnc_view import trended_score_view

import os
import fnmatch

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

class RNCView(ttk.Frame):
	def __init__(self, parent, controller):
		super().__init__(parent)

		self.parent = parent
		self.controller = controller
		
		# create widgets
		self.issue_trended_button = ttk.Button(self, text='Issue Trended', command=self.run_issue_trended)
		self.issue_trended_button.grid(row=1, column=1, padx=10, sticky="NSEW")

		self.scores_topline_button = ttk.Button(self, text='Scores Topline', command=self.run_scores_topline)
		self.scores_topline_button.grid(row=1, column=2, padx=10, sticky="NSEW")
		
		self.trended_score_button = ttk.Button(self, text='Trended Score', command=self.run_trended_score)
		self.trended_score_button.grid(row=1, column=3, padx=10, sticky="NSEW")

	def run_issue_trended(self):
		# figure out where main window has ended up
		root_x = self.parent.winfo_rootx()
		root_y = self.parent.winfo_rooty()
		
		# create a popup-window
		issue_trended_window = tk.Toplevel()
		
		offset_x = root_x + 10
		offset_y = root_y + 10
		
		issue_trended_window.geometry(f'+{offset_x}+{offset_y}')
		
		# create issue trended view and place it in window
		self.issue_trended = issue_trended_view.IssueTrendedView(issue_trended_window, self.controller)
		self.issue_trended.grid(row=0, column=0, padx=10, pady=10)

	def issue_trended_show_success(self, message):
		self.issue_trended.show_loading_success(message)

	def issue_trended_show_error(self, message):
		self.issue_trended.show_loading_error(message)

	def scores_topline_show_success(self, message):
		self.scores_topline.show_loading_success(message)

	def scores_topline_show_error(self, message):
		self.scores_topline.show_loading_error(message)
		
	def rnc_trended_show_success(self, message):
		self.trended_score.show_loading_success(message)

	def rnc_trended_show_error(self, message):
		self.trended_score.show_loading_error(message)

	def run_scores_topline(self):
		# figure out where main window has ended up
		root_x = self.parent.winfo_rootx()
		root_y = self.parent.winfo_rooty()
		
		# create a popup-window
		scores_topline_window = tk.Toplevel()
		
		offset_x = root_x + 10
		offset_y = root_y + 10
		
		scores_topline_window.geometry(f'+{offset_x}+{offset_y}')
		
		# create issue trended view and place it in window
		self.scores_topline = scores_topline_view.ScoresToplineView(scores_topline_window, self.controller)
		self.scores_topline.grid(row=0, column=0, padx=10, pady=10)

	def run_trended_score(self):
		# figure out where main window has ended up
		root_x = self.parent.winfo_rootx()
		root_y = self.parent.winfo_rooty()
		
		# create a popup-window
		trended_score_window = tk.Toplevel()
		
		offset_x = root_x + 10
		offset_y = root_y + 10
		
		trended_score_window.geometry(f'+{offset_x}+{offset_y}')
		
		# create issue trended view and place it in window
		self.trended_score = trended_score_view.TrendedScoreView(trended_score_window, self.controller)
		self.trended_score.grid(row=0, column=0, padx=10, pady=10)





