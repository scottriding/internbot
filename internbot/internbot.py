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
from SPSS_Crosstabs_GUI import SPSSCrosstabsView
from RNC_GUI import RNCView
import sys


class Internbot:

    def __init__ (self, root):
        self.__window = root
        self.main_buttons()
        self.fpath = os.path.join(os.path.expanduser("~"), "Desktop")
        self.__embedded_fields = []

    def main_buttons(self):
        """
        Function establishes all the components of the main window
        :return: None
        """

        self.rnc = RNCView(self.__window, mov_x, mov_y, window_width, window_height, header_font, header_color, bot_render)
        #Button definitions
        button_frame =Tkinter.Frame(self.__window)
        button_frame.pack(side=Tkinter.LEFT, fill=Tkinter.BOTH, )
        btn_xtabs = Tkinter.Button(button_frame, text="Run crosstabs", padx=4, width=20, height=3, command=self.software_tabs_menu, relief=Tkinter.FLAT, highlightthickness=0)
        btn_report = Tkinter.Button(button_frame, text="Run topline report", padx=4, width=20, height=3,command=self.topline_menu, relief=Tkinter.FLAT, highlightthickness=0)
        btn_appen = Tkinter.Button(button_frame, text="Run topline appendix", padx=4, width=20, height=3,command=self.append_menu, relief=Tkinter.FLAT, highlightthickness=0)
        btn_rnc = Tkinter.Button(button_frame, text="Run RNC", padx=4, width=20, height=3,command=self.rnc.rnc_menu, relief=Tkinter.FLAT, highlightthickness=0)
        btn_terminal = Tkinter.Button(button_frame, text="Terminal Window", padx=4, width=20, height=3, command=self.terminal_window, relief=Tkinter.FLAT, highlightthickness=0)
        btn_quit = Tkinter.Button(button_frame, text="Quit", padx=4, width=20, height=3,command=self.__window.destroy, relief=Tkinter.GROOVE, highlightthickness=0)
        btn_bot = Tkinter.Button(button_frame, image=bot_render, padx=4, pady=10, width=158, height=45, borderwidth=0, highlightthickness=0, relief=Tkinter.FLAT, command=self.main_help_window)
        btn_bot.pack(padx=5, pady=3, side=Tkinter.TOP)
        btn_xtabs.pack(padx=5, side=Tkinter.TOP, expand=True)
        btn_report.pack(padx=5, side=Tkinter.TOP, expand=True)
        btn_appen.pack(padx=5, side=Tkinter.TOP, expand=True)
        btn_rnc.pack(padx=5, side=Tkinter.TOP, expand=True)
        btn_terminal.pack(padx=5, side=Tkinter.TOP, expand=True)
        btn_quit.pack(padx=5, side=Tkinter.TOP, expand=True)

        #Menubar Set Up
        # self.menubar = Tkinter.Menu(self.__window)
        # menu_xtabs = Tkinter.Menu(self.menubar, tearoff = 0)
        # menu_xtabs.add_command(label="Variable Script", command=self.variable_script)
        # menu_xtabs.add_command(label="Table Script", command=self.table_script)
        # menu_xtabs.add_command(label="Build Report", command=self.build_xtabs)
        # self.menubar.add_cascade(label="Crosstabs", menu=menu_xtabs)
        # menu_report = Tkinter.Menu(self.menubar, tearoff=0)
        # menu_report.add_command(label="Run Topline", command=self.topline_menu)
        # menu_report.add_command(label="Run Appendix", command=self.append_menu)
        # self.menubar.add_cascade(label="Topline", menu=menu_report)
        # menu_rnc = Tkinter.Menu(self.menubar, tearoff=0)
        # menu_rnc.add_command(label="Scores Topline", command=self.scores_window)
        # menu_rnc.add_command(label="Issue Trended", command=self.issue_trended_window)
        # menu_rnc.add_command(label="Trended Score", command=self.trended_scores_window)
        # self.menubar.add_cascade(label="RNC", menu=menu_rnc)
        # menu_quit = Tkinter.Menu(self.menubar, tearoff=0)
        # menu_quit.add_command(label="Good Bye", command=self.__window.destroy)
        # self.menubar.add_cascade(label="Quit", menu=menu_quit)
        # self.__window.config(menu=self.menubar)

        def help_pressed(event):
            self.main_help_window()

        self.__window.bind("<F1>", help_pressed)

    def main_help_window(self):
        """
        Function serves as an intro to internbot. Explains the help bot to the user.
        :return: None
        """
        help_window = Tkinter.Toplevel(self.__window)
        help_window.withdraw()

        width = 250
        height = 250
        help_window.geometry("%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))

        message = "\nWelcome to Internbot"
        Tkinter.Label(help_window, text=message, font=header_font, fg=header_color).pack()
        info_message = "You can find help information throughout"\
                       "\nInternbot by clicking the bot icon" \
                       "\n\nShe will tell you a little bit about" \
                       "\n what you need to input for the" \
                       "\nreport you are trying to create\n"
        Tkinter.Label(help_window, text=info_message, font=('Trade Gothic LT Pro', 14, )).pack()
        btn_ok = Tkinter.Button(help_window, text="Ok", command=help_window.destroy, height=3, width=20,  highlightthickness=0)
        btn_ok.pack(pady= 5, side=Tkinter.BOTTOM, expand=False)
        help_window.deiconify()

    def terminal_window(self):
        self.term_window = Tkinter.Toplevel(self.__window)
        self.term_window.withdraw()
        width = 500
        height = 600
        self.term_window.geometry("%dx%d+%d+%d" % (
            width, height, mov_x + window_width / 2 - width , mov_y + window_height / 2 - height/2))


        t1 = Tkinter.Text(self.term_window, height=100, width=500, fg='white', background=header_color)
        t1.pack()

        class PrintToT1(object):
            def __init__(self, stream):
                self.stream = stream

            def write(self, s):
                self.stream.write(s)
                t1.insert(Tkinter.END, s)
                self.stream.flush()

        #sys.stdout = os.fdopen(sys.stdout.fileno(), 'w', 0)
        sys.stdout = PrintToT1(sys.stdout)
        sys.stderr = PrintToT1(sys.stderr)


        #self.__window.wm_attributes("-topmost", 1)
        #self.__window.focus_force()
        print "Hello World"
        self.term_window.deiconify()

    def software_tabs_menu(self):
        """
        Function sets up the Software Type selection for crosstabs
        :return:
        """
        sft_window = Tkinter.Toplevel(self.__window)
        sft_window.withdraw()
        self.spss = SPSSCrosstabsView(self.__window, mov_x, mov_y, window_width, window_height, header_font, header_color)
        width = 200
        height = 250
        sft_window.geometry(
            "%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))

        message = "Please select crosstabs\nsoftware to use"
        Tkinter.Label(sft_window, text = message, font=header_font, fg=header_color).pack()
        btn_spss = Tkinter.Button(sft_window, text="SPSS", command=self.spss.spss_crosstabs_menu, height=3, width=20)
        btn_q = Tkinter.Button(sft_window, text="Q Research", command=self.bases_window, height=3, width=20)
        btn_cancel = Tkinter.Button(sft_window, text="Cancel", command=sft_window.destroy, height=3, width=20)
        btn_cancel.pack(side=Tkinter.BOTTOM, expand=True)
        btn_q.pack(side=Tkinter.BOTTOM, expand=True)
        btn_spss.pack(side=Tkinter.BOTTOM, expand=True)
        sft_window.deiconify()



    def bases_window(self):
        """
        Function sets up window for bases entry of Q crosstab reports
        :return:
        """

        self.base_window = Tkinter.Toplevel(self.__window)
        self.base_window.withdraw()
        width = 1100
        height = 350
        self.base_window.geometry("%dx%d+%d+%d" % (width, height, mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))
        self.base_window.title("Base assignment")
        self.loaded_qfile = False
        
        self.boxes_frame = Tkinter.Frame(self.base_window)
        self.boxes_frame.pack(padx=10, pady=10, side=Tkinter.LEFT, fill=Tkinter.BOTH)
        tables_horiz_scrollbar = Tkinter.Scrollbar(self.boxes_frame)
        tables_horiz_scrollbar.config(orient= Tkinter.HORIZONTAL )
        tables_horiz_scrollbar.pack(side=Tkinter.BOTTOM, fill=Tkinter.X)
        tables_vert_scrollbar = Tkinter.Scrollbar(self.boxes_frame)
        tables_vert_scrollbar.pack(side=Tkinter.RIGHT, fill=Tkinter.Y)
        
        self.tables_box = Tkinter.Listbox(self.boxes_frame, selectmode=Tkinter.SINGLE, width=55, height=10, yscrollcommand=tables_vert_scrollbar.set, xscrollcommand = tables_horiz_scrollbar.set, exportselection=False)
        self.tables_box.pack(expand=True, side = Tkinter.LEFT, fill=Tkinter.BOTH)

        self.bases_frame = Tkinter.Frame(self.base_window)
        self.bases_frame.pack(padx=10, pady=10, side=Tkinter.RIGHT, fill=Tkinter.BOTH)
        bases_horiz_scrollbar = Tkinter.Scrollbar(self.bases_frame)
        bases_horiz_scrollbar.config(orient=Tkinter.HORIZONTAL)
        bases_horiz_scrollbar.pack(side=Tkinter.BOTTOM, fill=Tkinter.X)
        bases_vert_scrollbar = Tkinter.Scrollbar(self.bases_frame)
        bases_vert_scrollbar.pack(side=Tkinter.RIGHT, fill=Tkinter.Y)

        self.bases_box = Tkinter.Listbox(self.bases_frame, selectmode=Tkinter.SINGLE, width =54, height = 10, yscrollcommand=bases_vert_scrollbar.set, xscrollcommand = bases_horiz_scrollbar.set, exportselection=False)
        self.bases_box.pack(expand=False, side=Tkinter.RIGHT, fill=Tkinter.BOTH)

        self.focus_index = 0

        self.tables_box.select_set(self.focus_index)
        self.bases_box.select_set(self.focus_index)

        
        tables_vert_scrollbar.config(command=self.tables_box.yview)
        tables_horiz_scrollbar.config(command=self.tables_box.xview)
        bases_vert_scrollbar.config(command=self.bases_box.yview)
        bases_horiz_scrollbar.config(command=self.bases_box.xview)

        buttons_frame = Tkinter.Frame(self.base_window)
        buttons_frame.pack(pady=10, side = Tkinter.RIGHT, fill=Tkinter.BOTH)
        btn_edit = Tkinter.Button(buttons_frame, text="Edit", command=self.parse_bases, width=20, height=3)
        btn_done = Tkinter.Button(buttons_frame, text="Done", command=self.finish_bases, width=20, height =3)
        btn_open = Tkinter.Button(buttons_frame, text="Open", command=self.qtab_build, width=20, height=3)
        btn_cancel = Tkinter.Button(buttons_frame, text="Cancel", command=self.base_window.destroy, width=20, height=3)
        btn_bot = Tkinter.Button(buttons_frame, image=bot_render, borderwidth=0,
                                 highlightthickness=0, relief=Tkinter.FLAT, bg="white", height=65, width=158,
                                 command=self.bases_help_window)
        btn_cancel.pack(pady=2, side=Tkinter.BOTTOM)
        btn_done.pack(pady=2, side=Tkinter.BOTTOM)
        btn_edit.pack(pady=2, side=Tkinter.BOTTOM)
        btn_open.pack(pady=2, side=Tkinter.BOTTOM)
        btn_bot.pack(pady=2, side=Tkinter.BOTTOM)

        def enter_pressed(event):
            self.parse_bases()

        def mouse_clicked(event):
            for index in self.tables_box.curselection():
                self.tables_box.select_clear(index)
            for index in self.bases_box.curselection():
                self.focus_index = index
                self.tables_box.select_set(self.focus_index)
                self.bases_box.select_set(self.focus_index)

        def up_pressed(event):
            for index in self.tables_box.curselection():
                self.tables_box.select_clear(index)

            for index in self.bases_box.curselection():
                self.bases_box.select_clear(index)

            if self.focus_index > 0:
                self.focus_index-=1
            self.tables_box.select_set(self.focus_index)
            self.bases_box.select_set(self.focus_index)

        def down_pressed(event):
            for index in self.tables_box.curselection():
                self.tables_box.select_clear(index)

            for index in self.bases_box.curselection():
                self.bases_box.select_clear(index)

            if self.focus_index <self.tables_box.size() -1:
                self.focus_index += 1
            self.tables_box.select_set(self.focus_index)
            self.bases_box.select_set(self.focus_index)

        def help_pressed(event):
            self.bases_help_window()

        self.base_window.bind("<F1>", help_pressed)

        self.base_window.bind("<Return>", enter_pressed)
        self.base_window.bind("<KP_Enter>", enter_pressed)

        self.base_window.bind("<Up>", up_pressed)

        self.base_window.bind("<Down>", down_pressed)

        self.base_window.bind("<Button-1>", mouse_clicked)

        self.base_window.deiconify()

    def qtab_build(self):
        """
        Funtion opens a Q Research .xlsx file and loads the table information into base_window for display to user
        :return: None
        """

        if not self.loaded_qfile:
            report_file_name = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select report file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
            if report_file_name is not "":
                self.__parser = crosstabs.Format_Q_Report.QParser(report_file_name)
                self.tables = self.__parser.format_report()
                prompts = []
                for key, value in self.tables.iteritems():
                    prompts.append(value.prompt)

                for prompt in prompts:
                    self.tables_box.insert(Tkinter.END, prompt)
                base_sizes = ["[Enter base details]"] * len(prompts)
                for size in base_sizes:
                    self.bases_box.insert(Tkinter.END, size)

                self.focus_index = 0

                self.tables_box.select_set(self.focus_index)
                self.bases_box.select_set(self.focus_index)

                self.loaded_qfile = True
                self.base_window.focus_force()
        else:
            ask_lost_work = tkMessageBox.askyesno("Select Q Research report file",
                                                  "You will lose your work. \nDo you want to continue?")
            if ask_lost_work:
                self.loaded_qfile=False
                self.qtab_build()

    def bases_help_window(self):
        """
        Function displays help information about how to use bases window.
        A part of the Q Research crosstabs process.
        :return: None
        """
        help_window = Tkinter.Toplevel(self.__window)
        help_window.lift()
        help_window.withdraw()
        width=300
        height=400
        help_window.geometry("%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))
        message = "\nBase Input Help"
        Tkinter.Label(help_window, text=message, font=header_font, fg = header_color).pack()
        info_message = "Select Open to import a Q output file."\
                       "\nThis should be a .xlsx file."\
                       "\n\nYou can change the current selection with"\
                       "\nthe up/down arrow keys or your mouse."\
                       "\n\nEdit the base and n info of the current"\
                       "\nselection by selecting Edit or hitting enter."\
                       "\n\nWhen you're done, select Done and you"\
                       "\nwill be prompted to select the directory of"\
                       "\nyour finished report.\n"
        Tkinter.Label(help_window, text=info_message, justify=Tkinter.LEFT).pack()
        btn_ok = Tkinter.Button(help_window, text="Ok", command=help_window.destroy, height=3, width=20)
        btn_ok.pack(padx=5, side=Tkinter.BOTTOM, expand=False)
        help_window.deiconify()

    def parse_bases(self):
        """
        Function calls edit_base for every entry in current selection of base_window
        :return: None
        """
        if self.loaded_qfile:
            if len(self.bases_box.curselection()) is 0:
                tkMessageBox.askokcancel("Select from Right",
                                         "Please make a selection from the list on the right.")
            else:
                for index in self.tables_box.curselection():
                    self.tables_box.select_clear(index)
                for index in self.bases_box.curselection():
                    self.focus_index = index
                    self.edit_base_window(index)

        else:
            tkMessageBox.askokcancel("Select Q Research report file",
                                     "Please open the Q Research report file first.")

    def edit_base_window(self, index):
        """
        Function defines a window for the user to input Base and N for a table.
        :param base:
        :param index:
        :return:
        """
        edit_window = Tkinter.Toplevel(self.base_window)
        width=400
        height=150
        edit_window.geometry("%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))
        edit_window.title("Add base and n")
        edit_frame =Tkinter.Frame(edit_window)
        edit_frame.pack(side=Tkinter.TOP)
        base_frame = Tkinter.Frame(edit_frame)
        base_frame.pack(side=Tkinter.LEFT, padx=5, pady=5)
        lbl_base = Tkinter.Label(base_frame, text="Base description:")
        entry_des = Tkinter.Entry(base_frame, width =30)
        entry_des.insert(0, "")
        entry_des.focus_set()
        size_frame = Tkinter.Frame(edit_frame)
        size_frame.pack(side=Tkinter.RIGHT, padx=5, pady=5)
        lbl_title = Tkinter.Label(size_frame, text="Size (n):")
        entry_size = Tkinter.Entry(size_frame, width= 10)
        entry_size.insert(0, "")

        def edit():
            """
            Internal function used by edit_base to edit the text of what appears in bases_window
            :return: None
            """
            base_desc = entry_des.get()
            base_size = entry_size.get()
            self.bases_box.delete(int(index))
            self.bases_box.insert(int(index), base_desc + " - " + base_size)
            edit_window.destroy()
            self.focus_index += 1
            self.tables_box.select_set(self.focus_index)
            self.bases_box.select_set(self.focus_index)


        btn_cancel = Tkinter.Button(edit_window, text="Cancel", command=edit_window.destroy, width = 20, height =3)
        btn_edit = Tkinter.Button(edit_window, text="Edit", command=edit, width = 20, height =3)

        lbl_base.pack(side=Tkinter.TOP)
        lbl_title.pack(side=Tkinter.TOP)
        entry_des.pack(side=Tkinter.TOP)
        entry_size.pack(side=Tkinter.TOP)

        btn_edit.pack(side=Tkinter.LEFT, padx=20)
        btn_cancel.pack(side=Tkinter.RIGHT, padx=20)

        #Key Bindings
        def enter_pressed(event):
            edit()

        def right_pressed(event):
            entry_size.focus_set()

        def left_pressed(event):
            entry_des.focus_set()

        def escape_pressed(event):
            edit_window.destroy()

        edit_window.bind("<Return>", enter_pressed)
        edit_window.bind("<KP_Enter>", enter_pressed)
        edit_window.bind("<Left>", left_pressed)
        edit_window.bind("<Right>", right_pressed)
        edit_window.bind("<Escape>", escape_pressed)

    def finish_bases(self):
        """
        Function gets save directory for finished crosstab report.
        :return: None
        """
        if self.loaded_qfile:
            index = 1
            for item in list(self.bases_box.get(0, Tkinter.END)):
                #base=item.split(" - ")
                current_table = self.tables.get(index)
                current_table.description = item
                current_table.size = 0
                index += 1

            self.__parser.add_bases(self.tables)

            ask_output = tkMessageBox.askokcancel("Select Output Directory", "Please select the directory for final report.")
            if ask_output is True: # user selected ok
                savedirectory = tkFileDialog.askdirectory()
                if savedirectory is not "":
                    self.__parser.save(savedirectory)
            self.edit_window.destroy()
        else:
            tkMessageBox.askokcancel("Select Q Research report file",
                                     "Please open the Q Research report file first.")

    def topline_menu(self):
        """
        Function sets up menu for entry of round number and open of topline file.
        :return: None
        """
        print "Here"
        self.redirect_window = Tkinter.Toplevel(self.__window)
        self.redirect_window.withdraw()
        width=200
        height=300
        self.redirect_window.geometry("%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))
        self.redirect_window.title("Y2 Topline Report Automation")
        message = "Please open a survey file."
        Tkinter.Label(self.redirect_window, text=message, font=header_font, fg=header_color).pack(side=Tkinter.TOP, pady=10)

        # round details
        round_frame = Tkinter.Frame(self.redirect_window, width=20)
        round_frame.pack(side=Tkinter.TOP, padx=20)

        round_label = Tkinter.Label(round_frame, text="Round number:", width=15)
        round_label.pack(side=Tkinter.LEFT)
        self.round_entry = Tkinter.Entry(round_frame)
        self.round_entry.pack(side=Tkinter.RIGHT, expand=True)
        
        btn_open = Tkinter.Button(self.redirect_window, text="Open", command = self.read_topline, height=3, width=20)
        btn_cancel = Tkinter.Button(self.redirect_window, text="Cancel", command=self.redirect_window.destroy, height=3, width=20)
        btn_bot = Tkinter.Button(self.redirect_window, image=bot_render,  borderwidth=0,
                                 highlightthickness=0, relief=Tkinter.FLAT, bg="white", height=65, width=158, command=self.topline_help_window)
        btn_bot.pack(side=Tkinter.TOP, padx=10, pady=5)
        btn_open.pack(side=Tkinter.TOP, padx=10, pady=5)
        btn_cancel.pack(side=Tkinter.TOP, padx=10, pady=5)
        self.redirect_window.deiconify()

    def append_menu(self):
        self.redirect_window = Tkinter.Toplevel(self.__window)
        self.redirect_window.withdraw()
        x = self.__window.winfo_x()
        y = self.__window.winfo_y()
        self.redirect_window.geometry("250x100+%d+%d" % (x + 175, y + 150))
        self.redirect_window.title("Y2 Topline Appendix Report Automation")
        message = "Please open an appendix file."
        Tkinter.Label(self.redirect_window, text = message).pack(expand=True)
        btn_open = Tkinter.Button(self.redirect_window, text="Open", command=self.read_topline, height=3, width=20)
        btn_cancel = Tkinter.Button(self.redirect_window, text="Cancel", command=self.redirect_window.destroy, height=3,
                                    width=20)
        btn_bot = Tkinter.Button(self.redirect_window, image=bot_render, borderwidth=0,
                                 highlightthickness=0, relief=Tkinter.FLAT, bg="white", height=65, width=158,
                                 command=self.topline_help_window)
        btn_bot.pack(side=Tkinter.TOP, padx=10, pady=5)
        btn_open.pack(side = Tkinter.TOP, padx=10, pady=5)
        btn_cancel.pack(side = Tkinter.TOP, padx=10, pady=5)

        self.redirect_window.deiconify()

    def topline_help_window(self):
        """
        Funtion sets up help window to give the user info about round numbers.
        :return: None
        """
        help_window = Tkinter.Toplevel(self.__window)
        help_window.withdraw()
        width=250
        height=400
        help_window.geometry("%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))
        message = "\nTopline Help"
        Tkinter.Label(help_window, text=message, font=header_font, fg=header_color).pack(side=Tkinter.TOP, padx=10)
        info_message = "\nRound Number:\n\n" \
                       "Leave it blank  or enter 1 for \n" \
                       "non-trended reports.\n" \
                       "Enter  a number (2-10) for the \n" \
                       "corresponding number of years/\n" \
                       "quarters for trended reports.\n\n"\
                       "When done, select open and you\n" \
                       "will be prompted to open the \n" \
                       "Topline file and selesct a save\n" \
                       "directory."

        Tkinter.Label(help_window, text=info_message, justify=Tkinter.LEFT).pack(side=Tkinter.TOP)
        btn_ok = Tkinter.Button(help_window, text="Ok", command=help_window.destroy, height=3, width=20)
        btn_ok.pack(padx=5, pady=10, side=Tkinter.TOP, expand=False)
        help_window.deiconify()

    def append_menu(self):
        """
        Function sets up Topline Appendix menu.
        :return: None
        """
        self.redirect_window=Tkinter.Toplevel(self.__window)
        self.redirect_window.withdraw()
        width=200
        height=300

        self.redirect_window.geometry("%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))
        self.redirect_window.title("Y2 Topline Appendix Report Automation")
        message="Please open a survey \nfile and open ends file."
        Tkinter.Label(self.redirect_window, text=message, font=header_font, fg=header_color).pack(expand=True, pady=5)
        btn_open = Tkinter.Button(self.redirect_window, text="Open", command=self.read_topline, height=3, width=20)
        btn_cancel = Tkinter.Button(self.redirect_window, text="Cancel", command=self.redirect_window.destroy, height=3,
                                    width=20)
        btn_bot = Tkinter.Button(self.redirect_window, image=bot_render, borderwidth=0,
                                 highlightthickness=0, relief=Tkinter.FLAT, bg="white", height=65, width=158,
                                 command=self.topline_help_window)
        btn_bot.pack(side=Tkinter.TOP, padx=10, pady=5)
        btn_open.pack(side=Tkinter.TOP, padx=10, pady=5)
        btn_cancel.pack(side=Tkinter.TOP, padx=10, pady=5)
        self.redirect_window.deiconify()

    def read_append(self):
        """
        reads in the appendix file and creates a .docx
        :return:
        """
        generator = topline.Appendix.AppendixGenerator()
        csvfilename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select open ends file", filetypes = (("Comma separated files", "*csv"), ("all files", "*.*")))
        if csvfilename != "":
            generator.parse_file(csvfilename)
            ask_docx = tkMessageBox.askyesno("Appendix Type", "Create a docx report?")
            if ask_docx is True:
                ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report.")
                if ask_output is True:
                    savedirectory = tkFileDialog.askdirectory()
                    generator.write_appendix(savedirectory, "appendix_template.docx", False)

            elif ask_docx is False:
                ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report.")
                if ask_output is True:
                    savedirectory = tkFileDialog.askdirectory()
                    generator.write_appendix(savedirectory, '', True)

        self.redirect_window.destroy()

    def read_topline(self):
        """
        Funtion reads in a topline file
        :return: None
        """
        try:
            filename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select survey file",filetypes = (("Qualtrics files","*.qsf"),("comma seperated files","*.csv"),("all files","*.*")))
            if filename is not "":
                self.isQSF = False
                if ".qsf" in filename:
                    pass
                elif ".csv" in filename:
                    round_int = self.round_entry.get()
                    is_trended = False
                    if round_int == "":
                        round_int = 1
                    else:
                        round_int = int(round_int)
                        if round_int != 1:
                            is_trended = True
                    self.report = topline.CSV.ReportGenerator(filename, round_int)
                    self.redirect_window.destroy()
                    self.round = round_int
                    if is_trended is True:
                        self.year_window_setup()
                    else:
                        self.build_topline_report(self.isQSF, self.report)
        except Exception as e:
            tkMessageBox.showerror("Error", "An error occurred\n"+ str(e))

    def year_window_setup(self):
        """
        Funtion sets up the window for entry of trend labels for trended Toplines
        :return: None
        """
        try:
            self.year_window = Tkinter.Toplevel(self.__window)
            self.year_window.withdraw()
            width=300
            height=200+self.round*25
            self.year_window.geometry("%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))
            self.year_window.title("Trended report years")
            message = "Please input the applicable years \nfor the trended topline report."
            Tkinter.Label(self.year_window, text=message, font=header_font, fg=header_color).pack(expand=True)

            year_frame = Tkinter.Frame(self.year_window)
            year_frame.pack(side=Tkinter.TOP, expand = True)
            self.year_window_obj = YearsWindow(self.__window, self.year_window, self.round)
            self.year_window_obj.packing_years(year_frame)
            btn_finish = Tkinter.Button(self.year_window, text="Done", command=self.build_topline_leadup, height=3, width=17)
            btn_cancel = Tkinter.Button(self.year_window, text="Cancel", command=self.year_window.destroy, height=3, width=17)
            btn_finish.pack(ipadx=5, side=Tkinter.LEFT, expand=False)
            btn_cancel.pack(ipadx=5, side=Tkinter.RIGHT, expand=False)
            self.year_window.deiconify()
        except Exception as e:
            tkMessageBox.showerror("Error", "An error occurred\n"+ str(e))

    def build_topline_leadup(self):
        """
        Function serves as a segue between the year_window_setup and the YearsWindow object
        for trended topline reports.
        :return: None
        """
        years = self.year_window_obj.read_years()
        self.build_topline_report(self.isQSF, self.report, years)

    def build_topline_report(self, isQSF, report, years=[]):
        """
        Function builds topline report and requests if the user would like to have the files opened.
        :param isQSF: Boolean value
        :param report: topline report object
        :param years: array of trend titles
        :return: None
        """
        try:
            template_file = open("topline_template.docx", "r")
            ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report.")
            if ask_output is True:
                savedirectory = tkFileDialog.askdirectory()
                if savedirectory is not "":
                    if isQSF is True:
                        pass
                    else:
                        report.generate_topline(template_file, savedirectory, years)
                        open_files = tkMessageBox.askyesno("Info", "Done!\nWould you like to open your finished files?")
                        if open_files is True:
                            self.open_file_for_user(savedirectory+"/topline_report.docx")
        except Exception as e:
            tkMessageBox.showerror("Error", "An error occurred\n" + str(e))

        self.redirect_window.title("RNC Scores Topline Automation")
        message = "Please open a model scores file."
        Tkinter.Label(self.redirect_window, text=message, font=header_font, fg=header_color).pack(expand=True)
        btn_topline = Tkinter.Button(self.redirect_window, text="Scores Topline Report", command=self.scores_window,
                                     height=3, width=20)
        btn_trended = Tkinter.Button(self.redirect_window, text="Issue Trended Report",
                                     command=self.issue_trended_window, height=3, width=20)
        btn_ind_trended = Tkinter.Button(self.redirect_window, text="Trended Score Reports",
                                         command=self.trended_scores_window, height=3, width=20)
        btn_cancel = Tkinter.Button(self.redirect_window, text="Cancel", command=self.redirect_window.destroy, height=3,
                                    width=20)
        btn_bot = Tkinter.Button(self.redirect_window, image=bot_render, borderwidth=0,
                                 highlightthickness=0, relief=Tkinter.FLAT, bg="white", height=65, width=158,
                                 command=self.rnc_help_window)
        btn_cancel.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        btn_topline.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        btn_trended.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        btn_ind_trended.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        btn_bot.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        self.redirect_window.deiconify()



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

window = Tkinter.Tk()
window.withdraw()
window.title("Internbot: 01011001 00000010") # Internbot: Y2
if platform.system() == 'Windows':  # Windows
    window.iconbitmap('y2.ico')
screen_width = window.winfo_screenwidth()

screen_height = window.winfo_screenheight()
mov_x = screen_width / 2 - 300
mov_y = screen_height / 2 - 200
window_height = 450
window_width = 600
window.geometry("%dx%d+%d+%d" % (window_width, window_height, mov_x, mov_y))


window['background'] = 'white'

y2_logo = "Y2Logo.gif"
help_bot = "Internbot.gif"
bot_render = Tkinter.PhotoImage(file=help_bot)
logo_render = Tkinter.PhotoImage(file= y2_logo)
logo_label = Tkinter.Label(window, image=logo_render, borderwidth=0, highlightthickness=0, relief=Tkinter.FLAT, padx=50)
logo_label.pack(side=Tkinter.RIGHT)

window.option_add("*Font", ('Trade Gothic LT Pro', 16, ))
window.option_add("*Button.Foreground", "midnight blue")

header_font = ('Trade Gothic LT Pro', 18, 'bold')
header_color = "midnight blue"
def close(event):
    window.withdraw() # if you want to bring it back
    sys.exit() # if you want to exit the entire thing

window.bind('<Escape>', close)

window.deiconify()
Internbot(window)
window.mainloop()