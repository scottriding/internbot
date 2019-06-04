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
import threading
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
        self.redirect_window = tkinter.Toplevel(self.__window)
        self.redirect_window.withdraw()
        width = 250
        height = 350
        self.redirect_window.geometry("%dx%d+%d+%d" % (
            width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))

        self.redirect_window.title("RNC Scores Topline Automation")
        message = "Please open a model scores file."
        tkinter.Label(self.redirect_window, text=message, font=self.header_font, fg=self.header_color).pack(expand=True)
        btn_topline = tkinter.Button(self.redirect_window, text="Scores Topline Report", command=self.scores_window,
                                     height=3, width=20)
        btn_trended = tkinter.Button(self.redirect_window, text="Issue Trended Report",
                                     command=self.issue_trended_window, height=3, width=20)
        btn_ind_trended = tkinter.Button(self.redirect_window, text="Trended Score Reports",
                                         command=self.trended_scores_window, height=3, width=20)
        btn_cancel = tkinter.Button(self.redirect_window, text="Cancel", command=self.redirect_window.destroy, height=3,
                                    width=20)
        btn_bot = tkinter.Button(self.redirect_window, image=self.bot_render, borderwidth=0,
                                 highlightthickness=0, relief=tkinter.FLAT, bg="white", height=65, width=158,
                                 command=self.rnc_help_window)
        btn_cancel.pack(ipadx=5, side=tkinter.BOTTOM, expand=False)
        btn_topline.pack(ipadx=5, side=tkinter.BOTTOM, expand=False)
        btn_trended.pack(ipadx=5, side=tkinter.BOTTOM, expand=False)
        btn_ind_trended.pack(ipadx=5, side=tkinter.BOTTOM, expand=False)
        btn_bot.pack(ipadx=5, side=tkinter.BOTTOM, expand=False)
        self.redirect_window.deiconify()

    def rnc_help_window(self):
        """
        Function serves as an intro to internbot. Explains the help bot to the user.
        :return: None
        """
        help_window = tkinter.Toplevel(self.__window)
        help_window.withdraw()

        width = 250
        height = 300
        help_window.geometry("%dx%d+%d+%d" % (width, height, self.mov_x + self.window_width/2 - width/2, self.mov_y + self.window_height/2 - height/2))

        message = "\nRNC Report Help"
        tkinter.Label(help_window, text=message, font=self.header_font, fg=self.header_color).pack()
        info_message = "Each button will prompt you to open\n" \
                       "a RNC model file.\n" \
                       "Trended Score Reports creates several reports \n" \
                       "based on model file that will be placed" \
                       "\nin a specified directory folder.\n" \
                       "Issue Trended Report creates a single,\n" \
                       "large excel file by tabbed by model.\n" \
                       "Scores Topline Report creates a general\n" \
                       "breakdown of models in a single excel file tab."

        tkinter.Label(help_window, text=info_message, font=('Trade Gothic LT Pro', 14,)).pack()
        btn_ok = tkinter.Button(help_window, text="Ok", command=help_window.destroy, height=3, width=20,
                                highlightthickness=0)
        btn_ok.pack(pady=5, side=tkinter.BOTTOM, expand=False)
        help_window.deiconify()

    def scores_window(self):
        self.filename = filedialog.askopenfilename(initialdir = self.fpath, title = "Select model file for scores report", filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
        if self.filename is not "":
            self.create_window = tkinter.Toplevel(self.redirect_window)
            self.create_window.withdraw()
            width = 250
            height = 250
            self.create_window.geometry("%dx%d+%d+%d" % (width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
            self.create_window.title("Scores Topline Report Details")
            # location details
            location_frame = tkinter.Frame(self.create_window)
            location_frame.pack(side = tkinter.TOP, expand=True)

            location_label = tkinter.Label(location_frame, text="Reporting region:")
            location_label.pack (side = tkinter.LEFT, expand=True)
            self.location_entry = tkinter.Entry(location_frame)
            self.location_entry.pack(side=tkinter.RIGHT, expand=True)
            self.location_entry.focus_set()

            # round details
            round_frame = tkinter.Frame(self.create_window)
            round_frame.pack(side=tkinter.TOP, expand=True)

            round_label = tkinter.Label(round_frame, text="Round number:")
            round_label.pack(padx=5, side=tkinter.LEFT, expand=True)
            self.round_entry = tkinter.Entry(round_frame)
            self.round_entry.pack(padx=7, side=tkinter.RIGHT, expand=True)

            # done and cancel buttons
            button_frame = tkinter.Frame(self.create_window)
            button_frame.pack(side=tkinter.TOP, expand=True)

            btn_cancel = tkinter.Button(self.create_window, text="Cancel", command=self.create_window.destroy, height=3, width=20)
            btn_cancel.pack(side=tkinter.BOTTOM, expand=True)
            btn_done = tkinter.Button(self.create_window, text="Done", command=self.scores_topline, height=3, width=20)
            btn_done.pack(side=tkinter.BOTTOM, expand=True)
            self.create_window.deiconify()

            def enter_pressed(event):
                self.scores_topline()

            self.create_window.bind("<Return>", enter_pressed)
            self.create_window.bind("<KP_Enter>", enter_pressed)

            def up_pressed(event):
                self.location_entry.focus_set()

            self.create_window.bind("<Up>", up_pressed)

            def down_pressed(event):
                self.round_entry.focus_set()

            self.create_window.bind("<Down>", down_pressed)



    def scores_topline(self):
        filename = self.filename
        #fields entered by user
        report_location = self.location_entry.get()
        round = self.round_entry.get()
        if round is not "" and report_location is not "":
            self.create_window.destroy()
            report = rnc_automation.ScoresToplineReportGenerator(filename, int(round))
            savedirectory = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('excel files', '.xlsx')])
            if savedirectory is not "":
                thread=threading.Thread(target=self.scores_topline_worker, args=(report, savedirectory, report_location, round))
                thread.start()
                self.create_window.destroy

    def scores_topline_worker(self, report, savedirectory, report_location, round):
        print("Running report...")
        report.generate_scores_topline(savedirectory, report_location, round)
        self.open_file_for_user(savedirectory)
        print("Done!")


    def issue_trended_window(self):
        self.filename = filedialog.askopenfilename(initialdir = self.fpath, title = "Select model file for issuse trended report", filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
        if self.filename is not "":
            self.create_window = tkinter.Toplevel(self.redirect_window)
            self.create_window.withdraw()
            width=250
            height=200
            self.create_window.geometry("%dx%d+%d+%d" % (width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
            self.create_window.title("Issue Trended Report Details")

            # round details
            round_frame = tkinter.Frame(self.create_window)
            round_frame.pack(side = tkinter.TOP, expand=True)

            round_label = tkinter.Label(round_frame, text="Round number:")
            round_label.pack(padx = 5, side = tkinter.LEFT, expand=True)
            self.round_entry = tkinter.Entry(round_frame)
            self.round_entry.pack(padx = 7, side=tkinter.RIGHT, expand=True)
            self.round_entry.focus_set()

            # done and cancel buttons
            button_frame = tkinter.Frame(self.create_window)
            button_frame.pack(side = tkinter.TOP, expand=True)

            btn_cancel = tkinter.Button(self.create_window, text = "Cancel", command = self.create_window.destroy,height=3, width=20)
            btn_cancel.pack(side = tkinter.BOTTOM, expand=True)
            btn_done = tkinter.Button(self.create_window, text = "Done", command = self.issue_trended,height=3, width=20)
            btn_done.pack(side = tkinter.BOTTOM, expand=True)
            self.create_window.deiconify()

            def enter_pressed(event):
                self.issue_trended()

            self.create_window.bind("<Return>", enter_pressed)
            self.create_window.bind("<KP_Enter>", enter_pressed)


    def issue_trended(self):
        filename = self.filename
        #fields entered by the user
        round = self.round_entry.get()
        self.create_window.destroy()
        report = rnc_automation.IssueTrendedReportGenerator(filename, int(round))
        savedirectory = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('excel files', '.xlsx')])
        if savedirectory is not "":
            thread = threading.Thread(target=self.issue_trended_worker, args=(report, savedirectory, round))
            thread.start()
            self.create_window.destroy


    def issue_trended_worker(self, report, savedirectory, round):
        print("Running report...")
        report.generate_issue_trended(savedirectory, round)
        self.open_file_for_user(savedirectory)
        print("Done!")


    def trended_scores_window(self):
        okay = messagebox.askokcancel("Select", "Select a model file")
        if okay is True:
            self.filename = filedialog.askopenfilename(initialdir = self.fpath, title = "Select model file", filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
            if self.filename is not "":
                self.create_window = tkinter.Toplevel(self.redirect_window)
                self.create_window.withdraw()
                width=250
                height=200
                self.create_window.geometry("%dx%d+%d+%d" % (width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
                self.create_window.title("Trended Issue Reports Details")

                # round details
                round_frame = tkinter.Frame(self.create_window)
                round_frame.pack(side = tkinter.TOP, expand=True)

                round_label = tkinter.Label(round_frame, text="Round number:")
                round_label.pack(padx = 5, side = tkinter.TOP, expand=True)
                self.round_entry = tkinter.Entry(round_frame)
                self.round_entry.pack(padx = 7, side=tkinter.BOTTOM, expand=True)
                self.round_entry.focus_set()

                # done and cancel buttons
                button_frame = tkinter.Frame(self.create_window)
                button_frame.pack(side = tkinter.TOP, expand=True)

                btn_cancel = tkinter.Button(self.create_window, text = "Cancel", command = self.create_window.destroy, height=3, width=20)
                btn_cancel.pack(side = tkinter.BOTTOM, expand=True)
                btn_done = tkinter.Button(self.create_window, text = "Done", command = self.trended_scores, height=3, width=20)
                btn_done.pack(side = tkinter.BOTTOM, expand=True)

                self.create_window.deiconify()

                def enter_pressed(event):
                    self.trended_scores()

                self.create_window.bind("<Return>", enter_pressed)
                self.create_window.bind("<KP_Enter>", enter_pressed)


    def trended_scores(self):
        filename = self.filename
        print(filename)
        #Field entered by user
        round = self.round_entry.get()
        if round is not "":
            self.create_window.destroy()
            report = rnc_automation.TrendedScoresReportGenerator(filename, int(round))
            ask_output = messagebox.askokcancel("Output directory", "Please select the directory for finished report\n This report will create a lot of file and may take several minutes. The program will appear to be unresponsive.")

            if ask_output is True:
                savedirectory = filedialog.askdirectory()
                self.create_window.update()
                if savedirectory is not "":
                    thread = threading.Thread(target=self.trended_scores_worker, args=(report, savedirectory, round))
                    thread.start()
                    self.create_window.destroy()

    def trended_scores_worker(self, report, savedirectory, round):
        print("Running reports...")
        report.generate_trended_scores(savedirectory, round)
        self.open_file_for_user(savedirectory)
        print("Done Creating Reports!")


    def open_file_for_user(self, file_path):
        try:
            if os.path.exists(file_path):
                if platform.system() == 'Darwin':  # macOS
                    subprocess.call(('open', file_path))
                elif platform.system() == 'Windows':  # Windows
                    os.startfile(file_path)
            else:
                messagebox.showerror("Error", "Error: Could not open file for you \n"+file_path)
        except IOError:
            messagebox.showerror("Error", "Error: Could not open file for you \n" + file_path)