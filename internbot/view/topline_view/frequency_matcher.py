import os
import fnmatch

import tkinter as tk
from tkinter import ttk

class Matcher(tk.Frame):
	def __init__(self, parent, csv_questions, qsf_questions, controller):

		tk.Frame.__init__(self, parent)
		self.canvas = tk.Canvas(self, borderwidth=0)
		self.frame = tk.Frame(self.canvas)
		self.vsb = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
		self.canvas.configure(yscrollcommand=self.vsb.set)

		self.vsb.pack(side="right", fill="y")
		self.canvas.pack(side="left", fill="both", expand=True)
		self.canvas.create_window((4,4), window=self.frame, anchor="nw",
								  tags="self.frame")

		self.frame.bind("<Configure>", self.onFrameConfigure)
		
		self.qsf_questions = qsf_questions
		self.csv_questions = csv_questions

		current_row = self.populate()
		
		# submit		
		self.submit_button = ttk.Button(self.frame, text='Submit', command=self.submit_button_clicked)
		self.submit_button.grid(row=current_row, column=2, pady=15, padx=10)
		
		self.controller = controller

	def populate(self):
		list_csv = self.generate_string_list(self.csv_questions)
		list_csv.insert(0, "Pick a new question name:")
		
		self.drop_menus = {}
		
		i = 1
		for question in self.qsf_questions.questions:
			if question.parent == "CompositeQuestion":
				for subquestion in question.questions:
					# qsf
					label = tk.Label(self.frame, text="%s:" % subquestion.name, justify=tk.RIGHT, anchor="e")
					label.grid(sticky = tk.E, row=i, column=0)
					
					# matched
					match_label = ttk.Entry(self.frame, justify=tk.LEFT)
					match_label.grid(sticky = tk.E, row=i, column=1)
					match_label.insert(0, subquestion.assigned)
					
					# rematch
					menu = tk.StringVar()
					menu.set("Pick a new question name:")
					
					drop = tk.OptionMenu(self.frame, menu, *list_csv)
					drop.grid(row=i, column=2)
					
					self.drop_menus[subquestion.name] = menu
					
					i += 1
			else:
				# qsf
				label = tk.Label(self.frame, text="%s:" % question.name, justify=tk.RIGHT, anchor="e")
				label.grid(sticky = tk.E, row=i, column=0)
				
				# matched
				match_label = ttk.Entry(self.frame, justify=tk.LEFT)
				match_label.grid(sticky = tk.E, row=i, column=1)
				match_label.insert(0, question.assigned)
				
				# rematch
				menu = tk.StringVar()
				menu.set("Pick a new question name:")
				
				drop = tk.OptionMenu(self.frame, menu, *list_csv)
				drop.grid(row=i, column=2)
				
				self.drop_menus[question.name] = menu
				
				i += 1
			
		return i

	def onFrameConfigure(self, event):
		'''Reset the scroll region to encompass the inner frame'''
		self.canvas.configure(scrollregion=self.canvas.bbox("all"))

	def generate_string_list(self, questions):
		to_list = []
		for question in questions.questions:
			if question.parent == "CompositeQuestion":
				for sub_question in question.questions:
					to_list.append(sub_question.name)
			else:
				to_list.append(question.name)

		return to_list

	def submit_button_clicked(self):
		"""
		Handle submit button click event
		:return:
		"""
		for key, value in self.drop_menus.items():
			if value.get() != "Pick a new question name:":				
				question_to_change = self.qsf_questions.find_question_by_name(key)
				rematch = self.csv_questions.find_question_by_name(value.get())
				
				question_to_change.type = rematch.type
				question_to_change.responses = rematch.responses
				
		self.controller.finalize_topline(self.qsf_questions)
				