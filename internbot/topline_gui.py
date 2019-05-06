import base
import crosstabs
import topline
import rnc_automation
import Tkinter
import tkMessageBox
import tkFileDialog
import os, subprocess, platform
import csv
from collections import OrderedDict
from years_window import YearsWindow

class ToplineView(object):

    def __init__(self, main_window, mov_x, mov_y, window_width, window_height, header_font, header_color, bot_render):
        self.__window = main_window
        self.mov_x = mov_x
        self.mov_y = mov_y
        self.window_width = window_width
        self.window_height = window_height
        self.header_font = header_font
        self.header_color = header_color
        self.bot_render=bot_render
        self.fpath = os.path.join(os.path.expanduser("~"), "Desktop")
        self.__embedded_fields = []

    def topline_menu(self):
        """
        Function sets up menu for entry of round number and open of topline file.
        :return: None
        """
        print "File Topline"
        self.redirect_window = Tkinter.Toplevel(self.__window)
        self.redirect_window.withdraw()
        width = 200
        height = 300
        self.redirect_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        self.redirect_window.title("Y2 Topline Report Automation")
        message = "Please open a survey file."
        Tkinter.Label(self.redirect_window, text=message, font=self.header_font, fg=self.header_color).pack(side=Tkinter.TOP,
                                                                                                  pady=10)

        # round details
        round_frame = Tkinter.Frame(self.redirect_window, width=20)
        round_frame.pack(side=Tkinter.TOP, padx=20)

        round_label = Tkinter.Label(round_frame, text="Round number:", width=15)
        round_label.pack(side=Tkinter.LEFT)
        self.round_entry = Tkinter.Entry(round_frame)
        self.round_entry.pack(side=Tkinter.RIGHT, expand=True)

        btn_open = Tkinter.Button(self.redirect_window, text="Open", command=self.read_topline, height=3, width=20)
        btn_cancel = Tkinter.Button(self.redirect_window, text="Cancel", command=self.redirect_window.destroy, height=3,
                                    width=20)
        btn_bot = Tkinter.Button(self.redirect_window, image=self.bot_render, borderwidth=0,
                                 highlightthickness=0, relief=Tkinter.FLAT, bg="white", height=65, width=158,
                                 command=self.topline_help_window)

        btn_bot.pack(side=Tkinter.TOP, padx=10, pady=5)
        btn_open.pack(side=Tkinter.TOP, padx=10, pady=5)
        btn_cancel.pack(side=Tkinter.TOP, padx=10, pady=5)
        self.redirect_window.deiconify()

    def topline_help_window(self):
        """
        Funtion sets up help window to give the user info about round numbers.
        :return: None
        """
        help_window = Tkinter.Toplevel(self.__window)
        help_window.withdraw()
        width = 250
        height = 400
        help_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        message = "\nTopline Help"
        Tkinter.Label(help_window, text=message, font=self.header_font, fg=self.header_color).pack(side=Tkinter.TOP, padx=10)
        info_message = "\nRound Number:\n\n" \
                       "Leave it blank  or enter 1 for \n" \
                       "non-trended reports.\n" \
                       "Enter  a number (2-10) for the \n" \
                       "corresponding number of years/\n" \
                       "quarters for trended reports.\n\n" \
                       "When done, select open and you\n" \
                       "will be prompted to open the \n" \
                       "Topline file and select a save\n" \
                       "directory."

        Tkinter.Label(help_window, text=info_message, justify=Tkinter.LEFT).pack(side=Tkinter.TOP)
        btn_ok = Tkinter.Button(help_window, text="Ok", command=help_window.destroy, height=3, width=20)
        btn_ok.pack(padx=5, pady=10, side=Tkinter.TOP, expand=False)
        help_window.deiconify()

    def read_topline(self):
        """
        Funtion reads in a topline file
        :return: None
        """
        try:
            filename = tkFileDialog.askopenfilename(initialdir=self.fpath, title="Select survey file", filetypes=(
            ("Qualtrics files", "*.qsf"), ("comma seperated files", "*.csv"), ("all files", "*.*")))
            if filename is not "":
                round_int = self.round_entry.get()
                is_trended = False
                if round_int == "":
                    round_int = 1
                else:
                    round_int = int(round_int)
                    if round_int != 1:
                        is_trended = True
                if ".qsf" in filename:
                    survey = base.QSFSurveyCompiler().compile(filename)
                    ask_freqs = tkMessageBox.askokcancel("Frequency file",
                                                         "Please select .csv file with survey result frequencies.")
                    if ask_freqs is True:
                        frequency_file = tkFileDialog.askopenfilename(initialdir=self.fpath,
                                                                      title="Select frequency file", filetypes=(
                            ("Command separated files", "*.csv"), ("all files", "*.*")))
                        if frequency_file is not "":
                            self.report_generator = topline.BasicReport.ReportGenerator(frequency_file, round_int,
                                                                                        survey)
                elif ".csv" in filename:
                    self.report_generator = topline.BasicReport.ReportGenerator(filename, round_int)
                    self.redirect_window.destroy()
                self.round = round_int
                if is_trended is True:
                    self.year_window_setup()
                else:
                    self.build_topline_report()
        except Exception as e:
            tkMessageBox.showerror("Error", "An error occurred\n" + str(e))

    def year_window_setup(self):
        """
        Funtion sets up the window for entry of trend labels for trended Toplines
        :return: None
        """
        try:
            self.year_window = Tkinter.Toplevel(self.__window)
            self.year_window.withdraw()
            width = 300
            height = 200 + self.round * 25
            self.year_window.geometry("%dx%d+%d+%d" % (
            width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
            self.year_window.title("Trended report years")
            message = "Please input the applicable years \nfor the trended topline report."
            Tkinter.Label(self.year_window, text=message, font=self.header_font, fg=self.header_color).pack(expand=True)

            year_frame = Tkinter.Frame(self.year_window)
            year_frame.pack(side=Tkinter.TOP, expand=True)
            self.year_window_obj = YearsWindow(self.__window, self.year_window, self.round)
            self.year_window_obj.packing_years(year_frame)
            btn_finish = Tkinter.Button(self.year_window, text="Done", command=self.build_topline_leadup, height=3,
                                        width=17)
            btn_cancel = Tkinter.Button(self.year_window, text="Cancel", command=self.year_window.destroy, height=3,
                                        width=17)
            btn_finish.pack(ipadx=5, side=Tkinter.LEFT, expand=False)
            btn_cancel.pack(ipadx=5, side=Tkinter.RIGHT, expand=False)
            self.year_window.deiconify()
        except Exception as e:
            tkMessageBox.showerror("Error", "An error occurred\n" + str(e))

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
        :param isQSF: Boolean value
        :param report: topline report object
        :param years: array of trend titles
        :return: None
        """
        try:
            template_file = open("templates_images/topline_template.docx", "r")
            ask_output = tkMessageBox.askokcancel("Output directory",
                                                  "Please select the directory for finished report.")
            if ask_output is True:
                savedirectory = tkFileDialog.askdirectory()
                if savedirectory is not "":
                    self.report_generator.generate_topline(template_file, savedirectory, years)
                    open_files = tkMessageBox.askyesno("Info", "Done!\nWould you like to open your finished files?")
                    if open_files is True:
                        self.open_file_for_user(savedirectory + "/topline_report.docx")
        except Exception as e:
            tkMessageBox.showerror("Error", "An error occurred\n" + str(e))

    def open_file_for_user(self, file_path):
        try:
            if os.path.exists(file_path):
                if platform.system() == 'Darwin':  # macOS
                    subprocess.call(('open', file_path))
                elif platform.system() == 'Windows':  # Windows
                    os.startfile(file_path)
            else:
                tkMessageBox.showerror("Error", "Error: Could not open file for you \n"+file_path)
        except IOError:
            tkMessageBox.showerror("Error", "Error: Could not open file for you \n" + file_path)