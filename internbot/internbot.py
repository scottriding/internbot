import crosstabs
import gui_windows
import tkinter
from tkinter import messagebox
from tkinter import filedialog
import os, subprocess, platform
import sys
import time
from shutil import copyfile


class Internbot:

    def __init__ (self, root):
        self.__window = root
        self.main_buttons()
        self.menu_bar_setup()
        localtime = time.asctime(time.localtime(time.time()))
        self.session_time_stamp = str("Session:  " + localtime + "   Internbot Version:  " + internbot_version)
        self.terminal_window()
        self.fpath = os.path.join(os.path.expanduser("~"), "Desktop")
        self.__embedded_fields = []
        self.terminal_open = False

    def main_buttons(self):
        """
        Function establishes all the components of the main window
        :return: None
        """

        self.rnc = gui_windows.RNCView(self.__window, mov_x, mov_y, window_width, window_height, header_font, header_color, bot_render)
        self.topline = gui_windows.ToplineView(self.__window, mov_x, mov_y, window_width, window_height, header_font, header_color, bot_render)
        self.spss = gui_windows.SPSSCrosstabsView(self.__window, mov_x, mov_y, window_width, window_height, header_font, header_color)
        self.q = gui_windows.QCrosstabsView(self.__window, mov_x, mov_y, window_width, window_height, header_font, header_color, bot_render)
        self.appendix = gui_windows.AppendixView(self.__window, mov_x, mov_y, window_width, window_height, header_font, header_color, bot_render)

        #Button definitions
        button_frame =tkinter.Frame(self.__window)
        button_frame.pack(side=tkinter.LEFT, fill=tkinter.BOTH, )
        btn_xtabs = tkinter.Button(button_frame, text="Run crosstabs", padx=4, width=20, height=3, command=self.software_tabs_menu, relief=tkinter.FLAT, highlightthickness=0)
        btn_report = tkinter.Button(button_frame, text="Run topline report", padx=4, width=20, height=3,command=self.topline.topline_menu, relief=tkinter.FLAT, highlightthickness=0)
        btn_appen = tkinter.Button(button_frame, text="Run topline appendix", padx=4, width=20, height=3,command=self.appendix.append_menu, relief=tkinter.FLAT, highlightthickness=0)
        btn_rnc = tkinter.Button(button_frame, text="Run RNC", padx=4, width=20, height=3,command=self.rnc.rnc_menu, relief=tkinter.FLAT, highlightthickness=0)
        btn_terminal = tkinter.Button(button_frame, text="Terminal Window", padx=4, width=20, height=3, command=self.reopen_terminal_window, relief=tkinter.FLAT, highlightthickness=0)
        btn_quit = tkinter.Button(button_frame, text="Quit", padx=4, width=20, height=3,command=self.__window.destroy, relief=tkinter.GROOVE, highlightthickness=0)
        btn_bot = tkinter.Button(button_frame, image=bot_render, padx=4, pady=10, width=158, height=45, borderwidth=0, highlightthickness=0, relief=tkinter.FLAT, command=self.main_help_window)
        btn_bot.pack(padx=5, pady=3, side=tkinter.TOP)
        btn_xtabs.pack(padx=5, side=tkinter.TOP, expand=True)
        btn_report.pack(padx=5, side=tkinter.TOP, expand=True)
        btn_appen.pack(padx=5, side=tkinter.TOP, expand=True)
        btn_rnc.pack(padx=5, side=tkinter.TOP, expand=True)
        btn_terminal.pack(padx=5, side=tkinter.TOP, expand=True)
        btn_quit.pack(padx=5, side=tkinter.TOP, expand=True)

    def menu_bar_setup(self):
        #Menubar Set Up
        self.menubar = tkinter.Menu(self.__window)
        menu_xtabs = tkinter.Menu(self.menubar, tearoff = 0)
        menu_xtabs.add_command(label="Crosstabs Menu", command=self.software_tabs_menu)
        menu_xtabs.add_command(label="SPSS", command=self.spss.spss_crosstabs_menu)
        menu_xtabs.add_command(label="Q Research", command=self.q.bases_window)
        menu_xtabs.add_command(label="Amazon Legacy", command=self.amazon_xtabs)
        self.menubar.add_cascade(label="Crosstabs", menu=menu_xtabs)
        menu_report = tkinter.Menu(self.menubar, tearoff=0)
        menu_report.add_command(label="Topline Menu", command=self.topline.topline_menu)
        menu_report.add_command(label="QSF and CSV", command=self.topline.read_qsf_topline)
        menu_report.add_command(label="CSV Only", command=self.topline.read_csv_topline)
        self.menubar.add_cascade(label="Topline", menu=menu_report)
        menu_appendix = tkinter.Menu(self.menubar, tearoff=0)
        menu_appendix.add_command(label="Appendix Menu", command=self.appendix.append_menu)
        menu_appendix.add_command(label="Word Appendix", command=self.appendix.doc_appendix)
        menu_appendix.add_command(label="Excel Appendix", command=self.appendix.excel_appendix_type)
        self.menubar.add_cascade(label="Appendix", menu=menu_appendix)
        menu_terminal = tkinter.Menu(self.menubar, tearoff=0)
        menu_terminal.add_command(label="Open Terminal", command=self.reopen_terminal_window)
        menu_terminal.add_command(label="Export Error Log", command=self.export_error_log)
        self.menubar.add_cascade(label="Terminal", menu=menu_terminal)
        menu_rnc = tkinter.Menu(self.menubar, tearoff=0)
        menu_rnc.add_command(label="RNC Menu", command=self.rnc.rnc_menu)
        menu_rnc.add_command(label="Scores", command=self.rnc.scores_window)
        menu_rnc.add_command(label="Issue Trended", command=self.rnc.issue_trended_window)
        menu_rnc.add_command(label="Trended Scores", command=self.rnc.trended_scores_window)
        self.menubar.add_cascade(label="RNC", menu=menu_rnc)
        menu_quit = tkinter.Menu(self.menubar, tearoff=0)
        menu_quit.add_command(label="Close Internbot", command=self.__window.destroy)
        self.menubar.add_cascade(label="Quit", menu=menu_quit)
        self.__window.config(menu=self.menubar)

    def main_help_window(self):
        """
        Function serves as an intro to internbot. Explains the help bot to the user.
        :return: None
        """
        help_window = tkinter.Toplevel(self.__window)
        help_window.withdraw()

        width = 250
        height = 500
        help_window.geometry("%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))

        message = "\nWelcome to Internbot"
        tkinter.Label(help_window, text=message, font=header_font, fg=header_color).pack()
        info_message = "You can find help information throughout"\
                       "\nInternbot by clicking the bot icon" \
                       "\n\nShe will tell you a little bit about" \
                       "\n what you need to input for the" \
                       "\nreport you are trying to create\n"
        tkinter.Label(help_window, text=info_message, font=('Trade Gothic LT Pro', 14, )).pack()
        term_message = "About the Terminal Window"
        tkinter.Label(help_window, text=term_message, font=header_font, fg=header_color).pack()
        term_info_message = "The terminal window will show info about\n" \
                            "the reports as you are running them.\n" \
                            "If an error occurs that you can't identify: \n" \
                            "Select Terminal>Export Error Log \n" \
                            "from the MenuBar. The file will appear  \n" \
                            "on your desktop. Then send a slack to the\n"\
                            "R&D channel with the error log and a link\n" \
                            "to the input file(s) of the report you were\n" \
                            "running. If you ever close the terminal window,\n" \
                            "you can reopen it with the Terminal Window\n" \
                            "button in the main window or Terminal>Open\n" \
                            " Terminal in the Menubar."
        tkinter.Label(help_window, text=term_info_message, font=('Trade Gothic LT Pro', 14,)).pack()
        btn_ok = tkinter.Button(help_window, text="Ok", command=help_window.destroy, height=3, width=20,  highlightthickness=0)
        btn_ok.pack(pady= 5, side=tkinter.BOTTOM, expand=False)
        help_window.deiconify()

        def enter_pressed(event):
            help_window.destroy()

        help_window.bind("<Return>", enter_pressed)
        help_window.bind("<KP_Enter>", enter_pressed)

    def terminal_window(self):
        self.term_window = True
        self.term_window = tkinter.Toplevel(self.__window)
        self.term_window.withdraw()
        self.term_window['background'] = header_color
        width = 500
        height = 600
        self.term_window.geometry("%dx%d+%d+%d" % (
            width, height, mov_x + window_width / 2 - width , mov_y + window_height / 2 - height/2))

        term_text = tkinter.Text(self.term_window, fg='white', height= 600, width=500, background=header_color, padx=5, pady=5)
        term_text.pack()

        self.error_log_filename = "templates_images/Error_Log.txt"
        self.error_log = open(self.error_log_filename, 'w')
        self.error_log.write("Error Log: " + self.session_time_stamp + "\n")


        class PrintToT1(object):
            def __init__(self, stream, error_log):
                self.stream = stream
                self.error_log = error_log

            def write(self, s):
                self.stream.write(s)
                term_text.insert(tkinter.END, s)
                self.error_log.write(s)
                self.stream.flush()
                term_text.see(tkinter.END)

        sys.stdout = PrintToT1(sys.stdout, self.error_log)
        sys.stderr = PrintToT1(sys.stderr, self.error_log)

        def update_terminal_flag():
            self.terminal_open = False
            self.term_window.withdraw()


        print("All information about reports and errors will appear in this window.\n")

        self.term_window.protocol('WM_DELETE_WINDOW', update_terminal_flag)
        self.term_window.deiconify()

    def reopen_terminal_window(self):
        self.terminal_open = True
        self.term_window.deiconify()

    def export_error_log(self):

        ask_export = messagebox.askokcancel("Export Error Log", "Warning: Internbot will close and you will need to restart it")
        if ask_export:
            self.error_log.close()
            copyfile(os.path.join(os.path.expanduser("~/Documents/GitHub/internbot/internbot/"), self.error_log_filename),os.path.join(os.path.expanduser("~"), "Desktop/internbot_error_log.txt"))
            self.open_file_for_user(os.path.join(os.path.expanduser("~"), "Desktop/internbot_error_log.txt"))
            self.__window.destroy()


    def software_tabs_menu(self):
        """
        Function sets up the Software Type selection for crosstabs
        :return:
        """
        sft_window = tkinter.Toplevel(self.__window)
        sft_window.withdraw()

        width = 200
        height = 250
        sft_window.geometry(
            "%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))

        message = "Please select crosstabs\nsoftware to use"
        tkinter.Label(sft_window, text = message, font=header_font, fg=header_color).pack()
        btn_spss = tkinter.Button(sft_window, text="SPSS", command=self.spss.spss_crosstabs_menu, height=3, width=20)
        btn_q = tkinter.Button(sft_window, text="Q Research", command=self.q.bases_window, height=3, width=20)
        btn_cancel = tkinter.Button(sft_window, text="Cancel", command=sft_window.destroy, height=3, width=20)
        btn_cancel.pack(side=tkinter.BOTTOM, expand=True)
        btn_q.pack(side=tkinter.BOTTOM, expand=True)
        btn_spss.pack(side=tkinter.BOTTOM, expand=True)
        sft_window.deiconify()

    def amazon_xtabs(self):
        """
        Function runs legacy report for Amazon CX Wave Series.
        """
        ask_xlsx = messagebox.askokcancel("Select XLSX Report File", "Please select combined tables .xlsx file")
        if ask_xlsx is True:
            tablefile = filedialog.askopenfilename(initialdir = self.fpath, title = "Select report file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
            if tablefile is not "":
                ask_output = messagebox.askokcancel("Select output directory", "Please select folder for final report.")
                if ask_output is True:
                    savedirectory = filedialog.askdirectory()
                    renamer = crosstabs.Format_Amazon_Report.RenameTabs()
                    renamed_wb = renamer.rename(tablefile, "templates_images/Amazon TOC.csv", savedirectory)
                    highlighter = crosstabs.Format_Amazon_Report.Highlighter(renamed_wb)
                    highlighter.highlight(savedirectory)
                    messagebox.showinfo("Finished", "The highlighted report is saved in your chosen directory.")

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

window = tkinter.Tk()
window.withdraw()
window.title("Internbot: 01011001 00000010") # Internbot: Y2
if platform.system() == 'Windows':  # Windows
    window.iconbitmap('templates_images/y2.ico')
screen_width = window.winfo_screenwidth()

screen_height = window.winfo_screenheight()
mov_x = screen_width / 2 - 300
mov_y = screen_height / 2 - 200
window_height = 450
window_width = 600
window.geometry("%dx%d+%d+%d" % (window_width, window_height, mov_x, mov_y))
window['background'] = 'white'

y2_logo = "templates_images/Y2Logo.gif"
help_bot = "templates_images/Internbot.gif"
bot_render = tkinter.PhotoImage(file=help_bot)
logo_render = tkinter.PhotoImage(file= y2_logo)
logo_label = tkinter.Label(window, image=logo_render, borderwidth=0, highlightthickness=0, relief=tkinter.FLAT, padx=50)
logo_label.pack(side=tkinter.RIGHT)

window.option_add("*Font", ('Trade Gothic LT Pro', 16, ))
window.option_add("*Button.Foreground", "midnight blue")

header_font = ('Trade Gothic LT Pro', 18, 'bold')
header_color = '#112C4E'
def close(event):
    window.withdraw() # if you want to bring it back
    sys.exit() # if you want to exit the entire thing

window.bind('<Escape>', close)

internbot_version = "1.0.0"

window.deiconify()
Internbot(window)
window.mainloop()