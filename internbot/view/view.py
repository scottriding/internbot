from view import topline_view
from view import appendix_view
from view import crosstabs_view
from view import rnc_view

import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext

class View(tk.Tk):
	def __init__(self):
		super().__init__()
		self.__controller = None
		
		# tk.Tk settings
		self.title("internbot 1.4.2")
		
		self.geometry("500x250")
		self.columnconfigure(0, weight=1)
		self.rowconfigure(0, weight=1)

		# menu (buttons above application)
		menubar = tk.Menu(self)
		
		filemenu = tk.Menu(menubar, tearoff=0)
		report_menu = tk.Menu(filemenu, tearoff=0)

		topline_menu = tk.Menu(report_menu, tearoff=0)
		topline_menu.add_command(label="Results", command=self.run_topline)
		topline_menu.add_command(label="Appendix", command=self.run_appendix)
		
		report_menu.add_cascade(label="Topline", menu=topline_menu)
		report_menu.add_command(label="Crosstabs", command=self.run_xtabs)

		filemenu.add_cascade(label="New", menu=report_menu)

		filemenu.add_separator()
		filemenu.add_command(label="Exit", command=self.quit)
		
		helpmenu = tk.Menu(menubar, tearoff=0)
		helpmenu.add_command(label="Help Index", command=self.do_nothing)
		helpmenu.add_command(label="About...", command=self.do_nothing)
		
		menubar.add_cascade(label="File", menu=filemenu)
		menubar.add_cascade(label="Help", menu=helpmenu)
		
		self.config(menu=menubar)
		
		
		# frame (body of window)
		self.__frame = tk.Frame(self, bg="white")
		self.__frame.grid(column=0, row=0, sticky="news")
		self.__frame.columnconfigure(0, weight=1)
		self.__frame.rowconfigure(0, weight=1)
		
		self.__saved_text = ["Welcome to internbot!"]
		self.__default_text = "Welcome to internbot!"
		self.__project_filename = None
		
		self.__text = scrolledtext.ScrolledText(self.__frame, bd=0, bg="white", fg="black", highlightthickness = 0, borderwidth=0, font=("Trade Gothic LT Pro", 14), undo=True, autoseparators=True, maxundo=-1, wrap=tk.WORD)
		self.__text.grid(column=0, row=0, sticky="news")
		self.__text.insert(tk.END, self.__default_text)

	@property
	def controller(self):
		return self.__controller
	
	@controller.setter
	def controller(self, controller):
		self.__controller = controller

	def do_nothing(self):
		filewin = tk.Toplevel(self)
		button = tk.Button(filewin, text="Button")
		button.pack()

	def run_topline(self):
		# figure out where main window has ended up
		root_x = self.winfo_rootx()
		root_y = self.winfo_rooty()
		
		# create a popup-window
		topline_window = tk.Toplevel()
		
		offset_x = root_x + 10
		offset_y = root_y + 10
		
		topline_window.geometry(f'+{offset_x}+{offset_y}')
		
		# create topline view and place it in window
		self.topline = topline_view.ToplineView(topline_window, self.controller)
		self.topline.grid(row=0, column=0, padx=10, pady=10)

	def show_topline_matcher(self, csv_questions, qsf_questions, groups):
		self.topline.show_matcher(csv_questions, qsf_questions, groups)
		
	def hide_topline_matcher(self):
		self.topline.hide_matcher()

	def topline_show_success(self, message):
		self.topline.show_loading_success(message)

	def topline_show_error(self, message):
		self.topline.show_loading_error(message)

	def run_appendix(self):
		# figure out where main window has ended up
		root_x = self.winfo_rootx()
		root_y = self.winfo_rooty()
		
		# create a popup-window
		appendix_window = tk.Toplevel()
		
		offset_x = root_x + 10
		offset_y = root_y + 10
		
		appendix_window.geometry(f'+{offset_x}+{offset_y}')
		
		# create appendix view and place it in window
		self.appendix = appendix_view.AppendixView(appendix_window, self.controller)
		self.appendix.grid(row=0, column=0, padx=10, pady=10)

	def appendix_show_success(self, message):
		self.appendix.show_loading_success(message)

	def appendix_show_error(self, message):
		self.appendix.show_loading_error(message)

	def run_xtabs(self):
		# figure out where main window has ended up
		root_x = self.winfo_rootx()
		root_y = self.winfo_rooty()
		
		# create a popup-window
		xtabs_window = tk.Toplevel()
		
		offset_x = root_x + 10
		offset_y = root_y + 10
		
		xtabs_window.geometry(f'+{offset_x}+{offset_y}')
		
		# create appendix view and place it in window
		self.xtabs = crosstabs_view.CrosstabsView(xtabs_window, self.controller)
		self.xtabs.grid(row=0, column=0, padx=10, pady=10)

	def crosstabs_show_success(self, message):
		self.xtabs.show_loading_success(message)

	def crosstabs_show_error(self, message):
		self.xtabs.show_loading_error(message)








