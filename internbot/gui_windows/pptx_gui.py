import base
import crosstabs
import topline
import rnc_automation
import tkinter
from tkinter import messagebox
from tkinter import filedialog
import os, subprocess, platform
import csv
from collections import OrderedDict
from gui_windows.years_window import YearsWindow
import threading



class PowerPointView(object):

    def __init__(self, main_window, mov_x, mov_y, window_width, window_height, header_font, header_color, bot_render, resources_filepath):
        self.__window = main_window
        self.mov_x = mov_x
        self.mov_y = mov_y
        self.window_width = window_width
        self.window_height = window_height
        self.header_font = header_font
        self.header_color = header_color
        self.fpath = os.path.join(os.path.expanduser("~"), "Desktop")
        self.__embedded_fields = []
        self.bot_render = bot_render
        self.resources_filepath = resources_filepath

    def pptx_menu(self):
        """
        Function sets up menu for entry of round number and open of topline file.
        :return: None
        """
        self.redirect_window = tkinter.Toplevel(self.__window)
        self.redirect_window.withdraw()
        width = 200
        height = 350
        self.redirect_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width/2 - width/2, self.mov_y + self.window_height / 2 - height / 2))
        self.redirect_window.title("Topline Menu")
        message = "Please open a survey file."
        tkinter.Label(self.redirect_window, text=message, font=self.header_font, fg=self.header_color).pack(side=tkinter.TOP,
                                                                                                  pady=10)

        # round details
        round_frame = tkinter.Frame(self.redirect_window, width=20)
        round_frame.pack(side=tkinter.TOP, padx=20)

        round_label = tkinter.Label(round_frame, text="Round number:", width=15)
        round_label.pack(side=tkinter.LEFT)
        self.round_entry = tkinter.Entry(round_frame)
        self.round_entry.pack(side=tkinter.RIGHT, expand=True)

        btn_qsf = tkinter.Button(self.redirect_window, text="Run Charts", command=self.read_qsf_topline, height=3, width=20)
        btn_cancel = tkinter.Button(self.redirect_window, text="Cancel", command=self.redirect_window.destroy, height=3,
                                    width=20)
        btn_bot = tkinter.Button(self.redirect_window, image=self.bot_render, borderwidth=0, highlightthickness=0,
                                relief=tkinter.FLAT, bg="white", height=65, width=158, command=self.pptx_help_window)

        btn_bot.pack(side=tkinter.TOP, padx=10)
        btn_qsf.pack(side=tkinter.TOP, padx=10)
        btn_cancel.pack(side=tkinter.TOP, padx=10)

        self.round_entry.focus_set()
        self.redirect_window.deiconify()

    def pptx_help_window(self):
        """
        Funtion sets up help window to give the user info about round numbers.
        :return: None
        """
        help_window = tkinter.Toplevel(self.__window)
        help_window.withdraw()
        width = 250
        height = 300
        help_window.geometry("%dx%d+%d+%d" % (
            width, height, self.mov_x + self.window_width / 2 - width / 2,
            self.mov_y + self.window_height / 2 - height / 2))
        help_window.title("Power Point Help")
        message = "\nPower Point Chart Help"
        tkinter.Label(help_window, text=message, font=self.header_font, fg=self.header_color).pack(side=tkinter.TOP,
                                                                                                   padx=10)
        info_message = "Help message here"

        tkinter.Label(help_window, text=info_message, justify=tkinter.LEFT).pack(side=tkinter.TOP)
        btn_ok = tkinter.Button(help_window, text="Ok", command=help_window.destroy, height=3, width=20)
        btn_ok.pack(padx=5, pady=10, side=tkinter.TOP, expand=False)
        help_window.deiconify()

        def enter_pressed(event):
            help_window.destroy()

        help_window.bind("<Return>", enter_pressed)
        help_window.bind("<KP_Enter>", enter_pressed)

    def read_qsf_topline(self):
        """
        Funtion reads in a topline file
        :return: None
        """
        print("Reading in QSF Topline")
        self.filename = filedialog.askopenfilename(initialdir=self.fpath, title="Select survey file", filetypes=(("Qualtrics files", "*.qsf"), ("comma seperated files", "*.csv"), ("all files", "*.*")))
        if self.filename is not "":
            round_int = self.round_entry.get()
            if round_int != "" and round_int != 1:
                thread = threading.Thread(target=self.year_window_setup(int(round_int)))
                thread.start()
            else:
                self.build_topline_report()

    def year_window_setup(self, rounds):
        """
        Funtion sets up the window for entry of trend labels for trended Toplines
        :return: None
        """

        self.year_window = tkinter.Toplevel(self.__window)
        self.year_window.withdraw()
        width = 300
        height = 200 + rounds * 25
        self.year_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        self.year_window.title("Trended Years Entry")
        message = "Please input the applicable years \nfor the trended topline report."
        tkinter.Label(self.year_window, text=message, font=self.header_font, fg=self.header_color).pack(expand=True)

        year_frame = tkinter.Frame(self.year_window)
        year_frame.pack(side=tkinter.TOP, expand=True)
        self.year_window_obj = YearsWindow(self.__window, self.year_window, rounds)
        self.year_window_obj.packing_years(year_frame)
        btn_finish = tkinter.Button(self.year_window, text="Done", command=self.build_topline_leadup, height=3, width=17)
        btn_cancel = tkinter.Button(self.year_window, text="Cancel", command=self.year_window.destroy, height=3, width=17)
        btn_finish.pack(ipadx=5, side=tkinter.LEFT, expand=False)
        btn_cancel.pack(ipadx=5, side=tkinter.RIGHT, expand=False)
        self.year_window.deiconify()

        def enter_pressed(event):
            self.build_topline_leadup()

        self.year_window.bind("<Return>", enter_pressed)
        self.year_window.bind("<KP_Enter>", enter_pressed)

    def build_topline_leadup(self):
        """
        Function serves as a segue between the year_window_setup and the YearsWindow object
        for trended topline reports.
        :return: None
        """
        years = self.year_window_obj.read_years()
        self.build_topline_report(years)

    def build_topline_report(self, years=[]):
        """
        Function builds topline report and requests if the user would like to have the files opened.
        :param years: array of trend titles
        :return: None
        """

        template_file = open(os.path.join(self.resources_filepath, "topline_template.docx"), "r")
        slides_generator = None
        if ".qsf" in self.filename:
            survey = base.QSFSurveyCompiler().compile(self.filename)
            ask_freqs = messagebox.askokcancel("Frequency file", "Please select .csv file with survey result frequencies.")
            if ask_freqs is True:
                frequency_file = filedialog.askopenfilename(initialdir=self.fpath, title="Select frequency file", filetypes=(("Comma separated files", "*.csv"), ("all files", "*.*")))
                if frequency_file is not "":
                    slides_generator = topline.PowerPoint.SlidesGenerator(frequency_file, years, survey)
        elif ".csv" in self.filename:
            pass

        ask_template = messagebox.askokcancel("Template", "Please select your template file (.pptx).")
        if ask_template is True:
            template_file = filedialog.askopenfilename(initialdir=self.fpath, title="Select frequency file", filetypes=(("power point files", "*.pptx"), ("all files", "*.*")))
            if template_file is not "":
                ask_output = messagebox.askokcancel("Output directory", "Please select the directory for finished report.")
                if ask_output is True:
                    savedirectory = filedialog.asksaveasfilename(defaultextension='.pptx', filetypes=[('power point files', '.pptx')])
                    if savedirectory is not "":
                        thread = threading.Thread(target=self.build_topline_worker, args=(slides_generator, template_file, savedirectory, years))
                        thread.start()

    def build_topline_worker(self, slides_generator, template_file, savedirectory, years):
        print("Running Topline Report...")
        slides_generator.generate_topline(template_file, savedirectory, years)
        print("Done!")
        self.open_file_for_user(savedirectory)
        self.redirect_window.destroy()

    def open_sound(self):

        def play_sound():
            audio_file = os.path.join(self.resources_filepath, "open.mp3")
            return_code = subprocess.call(["afplay", audio_file])

        thread_worker = threading.Thread(target=play_sound)
        thread_worker.start()

    def open_file_for_user(self, file_path):
        try:
            if os.path.exists(file_path):
                if platform.system() == 'Darwin':  # macOS
                    subprocess.call(('open', file_path))
                    self.open_sound()
                elif platform.system() == 'Windows':  # Windows
                    os.startfile(file_path)
            else:
                messagebox.showerror("Error", "Error: Could not open file for you \n" + file_path)
        except IOError:
            messagebox.showerror("Error", "Error: Could not open file for you \n" + file_path)


