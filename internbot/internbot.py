from model import model
from view import view

## outside modules
import tkinter as tk
from tkinter import ttk
import time

import os
import fnmatch
import sys

class Controller:

	def __init__(self, model, view):
		self.model = model
		self.view = view
		
		self.__topline_templates = {}
		self.__topline_templates["Y2"] = os.path.join(template_folder, "topline_template.docx")
		self.__topline_templates["QUALTRICS"] = ""

		self.__appendix_templates = {}
		self.__appendix_templates["Y2"] = os.path.join(template_folder, "appendix_template.docx")
		self.__appendix_templates["QUALTRICS"] = ""

		self.__template_logos = {}
		self.__template_logos["Y2"] = os.path.join(image_folder, "y2_xtabs.png")
		self.__template_logos["Qualtrics"] = os.path.join(image_folder, "QLogo.png")
		self.__template_logos["Whatsapp"] = os.path.join(image_folder, "whatsapp.png")
		self.__template_logos["Facebook"] = os.path.join(image_folder, "FB.png")

	def run_appendix(self, verbatims_path, output_path):
		# for future updates - current template is set only to Y2
		template_path = self.__appendix_templates.get("Y2")

		try:
			questions = self.model.build_appendix_model(verbatims_path)
			self.model.build_appendix_report(questions, output_path, template_path)
			self.view.appendix_show_success("Done!")
		except ValueError as error:
			self.view.appendix_show_error(error)

	def run_format_crosstabs(self, cross_path, output_path, template):
		try:
			self.model.format_qresearch_report(cross_path, self.__template_logos.get(template), output_path)
			self.view.crosstabs_show_success("Done!")
		except ValueError as error:
			self.view.crosstabs_show_error(error)

	def run_topline(self, frequency_path, output_path, survey_path, groups):
		# for future updates - current template is set only to Y2
		self.template_path = self.__topline_templates.get("Y2")
		self.output_path = output_path

		try:
			survey = self.model.survey(survey_path)
			questions = self.model.build_document_model(frequency_path, groups, survey)
			if survey:
				csv_questions = self.model.build_document_model(frequency_path, groups, None)
				self.view.show_topline_matcher(csv_questions, questions, groups)
			else:
				self.model.build_document_report(questions, groups, self.template_path, self.output_path)
				self.view.topline_show_success("Done!")
		except ValueError as error:
			self.view.topline_show_error(error)

	def finalize_topline(self, qsf_questions, groups):
		try:
			self.view.hide_topline_matcher()
			self.model.build_document_report(qsf_questions, groups, self.template_path, self.output_path)
			self.view.topline_show_success("Done!")
		except ValueError as error:
			self.view.topline_show_error(error)

class Internbot(tk.Tk):

	def __init__(self):
		super().__init__()
		
		self.title('internbot 1.4.1')

		# create a model
		internbot_model = model.Model()
		
		# create a view and place it on the root window
		internbot_view = view.View(self)
		internbot_view.grid(row=0, column=0, padx=10, pady=10)
		
		#create a controller
		controller = Controller(internbot_model, internbot_view)
		
		# set the controller to view
		internbot_view.set_controller(controller)

if __name__ == '__main__':
	print("Welcome to internbot!")
	
	template_folder = 'resources/templates/'
	image_folder = 'resources/images/'

	#template_folder = os.path.join(sys._MEIPASS, 'resources/templates/')
	#image_folder = os.path.join(sys._MEIPASS, 'resources/images/')
	
	internbot = Internbot()
	internbot.mainloop()	