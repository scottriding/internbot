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

import sys

class RNCView(object):

    def __init__(self, main_window, mov_x, mov_y, window_width, window_height, header_font, header_color, bot_render):
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

    def rnc_menu(self):
        print "File RNC"
        self.redirect_window = Tkinter.Toplevel(self.__window)
        self.redirect_window.withdraw()
        width = 250
        height = 350
        self.redirect_window.geometry("%dx%d+%d+%d" % (
            width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))

        self.redirect_window.title("RNC Scores Topline Automation")
        message = "Please open a model scores file."
        Tkinter.Label(self.redirect_window, text=message, font=self.header_font, fg=self.header_color).pack(expand=True)
        btn_topline = Tkinter.Button(self.redirect_window, text="Scores Topline Report", command=self.scores_window,
                                     height=3, width=20)
        btn_trended = Tkinter.Button(self.redirect_window, text="Issue Trended Report",
                                     command=self.issue_trended_window, height=3, width=20)
        btn_ind_trended = Tkinter.Button(self.redirect_window, text="Trended Score Reports",
                                         command=self.trended_scores_window, height=3, width=20)
        btn_cancel = Tkinter.Button(self.redirect_window, text="Cancel", command=self.redirect_window.destroy, height=3,
                                    width=20)
        btn_bot = Tkinter.Button(self.redirect_window, image=self.bot_render, borderwidth=0,
                                 highlightthickness=0, relief=Tkinter.FLAT, bg="white", height=65, width=158,
                                 command=self.rnc_help_window)
        btn_cancel.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        btn_topline.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        btn_trended.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        btn_ind_trended.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        btn_bot.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        self.redirect_window.deiconify()

    def rnc_help_window(self):
        """
        Function serves as an intro to internbot. Explains the help bot to the user.
        :return: None
        """
        help_window = Tkinter.Toplevel(self.__window)
        help_window.withdraw()

        width = 250
        height = 250
        help_window.geometry("%dx%d+%d+%d" % (
            width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))

        message = "\nRNC Report Help"
        Tkinter.Label(help_window, text=message, font=self.header_font, fg=self.header_color).pack()
        info_message = ""
        Tkinter.Label(help_window, text=info_message, font=('Trade Gothic LT Pro', 14,)).pack()
        btn_ok = Tkinter.Button(help_window, text="Ok", command=help_window.destroy, height=3, width=20,
                                highlightthickness=0)
        btn_ok.pack(pady=5, side=Tkinter.BOTTOM, expand=False)
        help_window.deiconify()

    def scores_window(self):
        try:
            okay = tkMessageBox.askokcancel("Select", "Select a model file")
            if okay is True:
                self.filename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select model file", filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
                if self.filename is not "":
                    self.create_window = Tkinter.Toplevel(self.redirect_window)
                    self.create_window.withdraw()
                    width = 250
                    height = 100
                    self.create_window.geometry("%dx%d+%d+%d" % (width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
                    self.create_window.title("Scores Topline Report Details")
                    # location details
                    location_frame = Tkinter.Frame(self.create_window)
                    location_frame.pack(side = Tkinter.TOP, expand=True)

                    location_label = Tkinter.Label(location_frame, text="Reporting region:")
                    location_label.pack (side = Tkinter.LEFT, expand=True)
                    self.location_entry = Tkinter.Entry(location_frame)
                    self.location_entry.pack(side=Tkinter.RIGHT, expand=True)

                    # round details
                    round_frame = Tkinter.Frame(self.create_window)
                    round_frame.pack(side=Tkinter.TOP, expand=True)

                    round_label = Tkinter.Label(round_frame, text="Round number:")
                    round_label.pack(padx=5, side=Tkinter.LEFT, expand=True)
                    self.round_entry = Tkinter.Entry(round_frame)
                    self.round_entry.pack(padx=7, side=Tkinter.RIGHT, expand=True)

                    # done and cancel buttons
                    button_frame = Tkinter.Frame(self.create_window)
                    button_frame.pack(side=Tkinter.TOP, expand=True)

                    btn_cancel = Tkinter.Button(self.create_window, text="Cancel", command=self.create_window.destroy)
                    btn_cancel.pack(side=Tkinter.RIGHT, expand=True)
                    btn_done = Tkinter.Button(self.create_window, text="Done", command=self.scores_topline)
                    btn_done.pack(side=Tkinter.RIGHT, expand=True)
                    self.create_window.deiconify()
        except Exception as e:
                tkMessageBox.showerror("Error", "An error occurred\n" + str(e))

    def scores_topline(self):
        try:
            filename = self.filename
            #fields entered by user
            report_location = self.location_entry.get()
            round = self.round_entry.get()
            if round is not "" and report_location is not "":
                self.create_window.destroy()
                report = rnc_automation.ScoresToplineReportGenerator(filename, int(round))
                ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report.")
                if ask_output is True:
                    savedirectory = tkFileDialog.askdirectory()
                    if savedirectory is not "":
                        report.generate_scores_topline(savedirectory, report_location, round)
                        open_files = tkMessageBox.askyesno("Info", "Done!\nWould you like to open your finished files?")
                        if open_files is True:
                            self.open_file_for_user(savedirectory + "/scores_topline.xlsx")
        except Exception as e:
            tkMessageBox.showerror("Error", "An error occurred\n" + str(e))

    def issue_trended_window(self):
        try:
            okay = tkMessageBox.askokcancel("Select", "Select a model file")
            if okay is True:
                self.filename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select model file", filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
                if self.filename is not "":
                    self.create_window = Tkinter.Toplevel(self.redirect_window)
                    self.create_window.withdraw()
                    width=250
                    height=100
                    self.create_window.geometry("%dx%d+%d+%d" % (width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
                    self.create_window.title("Issue Trended Report Details")

                    # round details
                    round_frame = Tkinter.Frame(self.create_window)
                    round_frame.pack(side = Tkinter.TOP, expand=True)

                    round_label = Tkinter.Label(round_frame, text="Round number:")
                    round_label.pack(padx = 5, side = Tkinter.LEFT, expand=True)
                    self.round_entry = Tkinter.Entry(round_frame)
                    self.round_entry.pack(padx = 7, side=Tkinter.RIGHT, expand=True)

                    # done and cancel buttons
                    button_frame = Tkinter.Frame(self.create_window)
                    button_frame.pack(side = Tkinter.TOP, expand=True)

                    btn_cancel = Tkinter.Button(self.create_window, text = "Cancel", command = self.create_window.destroy)
                    btn_cancel.pack(side = Tkinter.RIGHT, expand=True)
                    btn_done = Tkinter.Button(self.create_window, text = "Done", command = self.issue_trended)
                    btn_done.pack(side = Tkinter.RIGHT, expand=True)
                    self.create_window.deiconify()
        except Exception as e:
            tkMessageBox.showerror("Error", "An error occurred\n" + str(e))

    def issue_trended(self):
        try:
            filename = self.filename
            #fields entered by the user
            round = self.round_entry.get()
            self.create_window.destroy()
            report = rnc_automation.IssueTrendedReportGenerator(filename, int(round))
            ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report.")
            if ask_output is True:
                savedirectory = tkFileDialog.askdirectory()
                if savedirectory is not "":
                    report.generate_issue_trended(savedirectory, round)
                    self.create_window.destroy
                    open_files = tkMessageBox.askyesno("Info", "Done!\nWould you like to open your finished files?")
                    if open_files is True:
                        self.open_file_for_user(savedirectory + "/trended.xlsx")
        except Exception as e:
            tkMessageBox.showerror("Error", "An error occurred\n" + str(e))

    def trended_scores_window(self):
        try:
            okay = tkMessageBox.askokcancel("Select", "Select a model file")
            if okay is True:
                self.filename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select model file", filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
                if self.filename is not "":
                    self.create_window = Tkinter.Toplevel(self.redirect_window)
                    self.create_window.withdraw()
                    width=250
                    height=100
                    self.create_window.geometry("%dx%d+%d+%d" % (width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
                    self.create_window.title("Trended Issue Reports Details")

                    # round details
                    round_frame = Tkinter.Frame(self.create_window)
                    round_frame.pack(side = Tkinter.TOP, expand=True)

                    round_label = Tkinter.Label(round_frame, text="Round number:")
                    round_label.pack(padx = 5, side = Tkinter.TOP, expand=True)
                    self.round_entry = Tkinter.Entry(round_frame)
                    self.round_entry.pack(padx = 7, side=Tkinter.BOTTOM, expand=True)

                    # done and cancel buttons
                    button_frame = Tkinter.Frame(self.create_window)
                    button_frame.pack(side = Tkinter.TOP, expand=True)

                    btn_cancel = Tkinter.Button(self.create_window, text = "Cancel", command = self.create_window.destroy)
                    btn_cancel.pack(side = Tkinter.RIGHT, expand=True)
                    btn_done = Tkinter.Button(self.create_window, text = "Done", command = self.trended_scores)
                    btn_done.pack(side = Tkinter.RIGHT, expand=True)

                    self.create_window.deiconify()
        except Exception as e:
            tkMessageBox.showerror("Error", "An error occurred\n" + str(e))

    def trended_scores(self):
        try:
            filename = self.filename
            #Field entered by user
            round = self.round_entry.get()
            if round is not "":
                self.create_window.destroy()
                report = rnc_automation.TrendedScoresReportGenerator(filename, int(round))
                ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report\n This report will create a lot of file and may take several minutes. The program will appear to be unresponsive.")

                if ask_output is True:
                    savedirectory = tkFileDialog.askdirectory()
                    self.create_window.update()
                    if savedirectory is not "":
                        report.generate_trended_scores(savedirectory, round)
                        self.create_window.destroy()
                        tkMessageBox.showinfo("Info", "Done!")
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