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
from spss_xtabs_gui import SPSSCrosstabsView
from q_xtabs_gui import QCrosstabsView
from topline_gui import ToplineView
from appendix_gui import AppendixView
from rnc_gui import RNCView
import sys
import templates_images



class Internbot:

    def __init__ (self, root):
        self.__window = root
        self.main_buttons()
        self.fpath = os.path.join(os.path.expanduser("~"), "Desktop")
        self.__embedded_fields = []
        self.terminal_open = False

    def main_buttons(self):
        """
        Function establishes all the components of the main window
        :return: None
        """

        self.rnc = RNCView(self.__window, mov_x, mov_y, window_width, window_height, header_font, header_color, bot_render)
        self.topline = ToplineView(self.__window, mov_x, mov_y, window_width, window_height, header_font, header_color, bot_render)
        self.spss = SPSSCrosstabsView(self.__window, mov_x, mov_y, window_width, window_height, header_font,
                                      header_color)
        self.q = QCrosstabsView(self.__window, mov_x, mov_y, window_width, window_height, header_font,
                                header_color, bot_render)
        self.appendix = AppendixView(self.__window, mov_x, mov_y, window_width, window_height, header_font,
                                header_color, bot_render)

        #Button definitions
        button_frame =Tkinter.Frame(self.__window)
        button_frame.pack(side=Tkinter.LEFT, fill=Tkinter.BOTH, )
        btn_xtabs = Tkinter.Button(button_frame, text="Run crosstabs", padx=4, width=20, height=3, command=self.software_tabs_menu, relief=Tkinter.FLAT, highlightthickness=0)
        btn_report = Tkinter.Button(button_frame, text="Run topline report", padx=4, width=20, height=3,command=self.topline.topline_menu, relief=Tkinter.FLAT, highlightthickness=0)
        btn_appen = Tkinter.Button(button_frame, text="Run topline appendix", padx=4, width=20, height=3,command=self.appendix.append_menu, relief=Tkinter.FLAT, highlightthickness=0)
        btn_rnc = Tkinter.Button(button_frame, text="Run RNC", padx=4, width=20, height=3,command=self.rnc.rnc_menu, relief=Tkinter.FLAT, highlightthickness=0)
        btn_terminal = Tkinter.Button(button_frame, text="Terminal Window", padx=4, width=20, height=3, command=self.reopen_terminal_window, relief=Tkinter.FLAT, highlightthickness=0)
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
        self.menubar = Tkinter.Menu(self.__window)
        menu_xtabs = Tkinter.Menu(self.menubar, tearoff = 0)
        menu_xtabs.add_command(label="Choose a Software", command=self.software_tabs_menu)
        self.menubar.add_cascade(label="Crosstabs", menu=menu_xtabs)
        menu_report = Tkinter.Menu(self.menubar, tearoff=0)
        menu_report.add_command(label="Topline Menu", command=self.topline.topline_menu)
        self.menubar.add_cascade(label="Topline", menu=menu_report)
        menu_appendix = Tkinter.Menu(self.menubar, tearoff=0)
        menu_appendix.add_command(label="Appendix Menu", command=self.appendix.append_menu)
        self.menubar.add_cascade(label="Appendix", menu=menu_appendix)
        menu_rnc = Tkinter.Menu(self.menubar, tearoff=0)
        menu_rnc.add_command(label="RNC Menu", command=self.rnc.rnc_menu)
        self.menubar.add_cascade(label="RNC", menu=menu_rnc)
        menu_quit = Tkinter.Menu(self.menubar, tearoff=0)
        menu_quit.add_command(label="Good Bye", command=self.__window.destroy)
        self.menubar.add_cascade(label="Quit", menu=menu_quit)
        self.__window.config(menu=self.menubar)

        self.terminal_window()


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
        self.term_window = True
        self.term_window = Tkinter.Toplevel(self.__window)
        self.term_window.withdraw()
        self.term_window['background'] = header_color
        width = 500
        height = 600
        self.term_window.geometry("%dx%d+%d+%d" % (
            width, height, mov_x + window_width / 2 - width , mov_y + window_height / 2 - height/2))

        term_text = Tkinter.Text(self.term_window, fg='white', height= 600, width=500, background=header_color, padx=5, pady=5)
        term_text.pack()

        class PrintToT1(object):
            def __init__(self, stream):
                self.stream = stream

            def write(self, s):
                self.stream.write(s)
                term_text.insert(Tkinter.END, s)
                self.stream.flush()
                term_text.see(Tkinter.END)

        sys.stdout = PrintToT1(sys.stdout)
        sys.stderr = PrintToT1(sys.stderr)

        def update_terminal_flag():
            self.terminal_open = False
            self.term_window.withdraw()


        print "All information about reports and errors will appear in this window.\n"
        self.term_window.protocol('WM_DELETE_WINDOW', update_terminal_flag)
        self.term_window.deiconify()

    def reopen_terminal_window(self):
        self.terminal_open = True
        self.term_window.deiconify()

    def software_tabs_menu(self):
        """
        Function sets up the Software Type selection for crosstabs
        :return:
        """
        sft_window = Tkinter.Toplevel(self.__window)
        sft_window.withdraw()

        width = 200
        height = 300
        sft_window.geometry(
            "%dx%d+%d+%d" % (width,height,mov_x + window_width / 2 - width / 2, mov_y + window_height / 2 - height / 2))

        message = "Please select crosstabs\nsoftware to use"
        Tkinter.Label(sft_window, text = message, font=header_font, fg=header_color).pack()
        btn_amaz = Tkinter.Button(sft_window, text="Amazon CX", command=self.amazon_xtabs, height=3, width=20)
        btn_spss = Tkinter.Button(sft_window, text="SPSS", command=self.spss.spss_crosstabs_menu, height=3, width=20)
        btn_q = Tkinter.Button(sft_window, text="Q Research", command=self.q.bases_window, height=3, width=20)
        btn_cancel = Tkinter.Button(sft_window, text="Cancel", command=sft_window.destroy, height=3, width=20)
        btn_cancel.pack(side=Tkinter.BOTTOM, expand=True)
        btn_q.pack(side=Tkinter.BOTTOM, expand=True)
        btn_spss.pack(side=Tkinter.BOTTOM, expand=True)
        btn_amaz.pack(side=Tkinter.BOTTOM, expand=True)
        sft_window.deiconify()

    def amazon_xtabs(self):
        """
        Function runs legacy report for Amazon CX Wave Series.
        """
        ask_xlsx = tkMessageBox.askokcancel("Select XLSX Report File", "Please select combined tables .xlsx file")
        if ask_xlsx is True:
            tablefile = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select report file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
            ask_output = tkMessageBox.askokcancel("Select output directory", "Please select folder for final report.")
            if ask_output is True:
                savedirectory = tkFileDialog.askdirectory()
                renamer = crosstabs.Format_Amazon_Report.RenameTabs()
                renamed_wb = renamer.rename(tablefile, "templates_images/Amazon TOC.csv", savedirectory)
                highlighter = crosstabs.Format_Amazon_Report.Highlighter(renamed_wb)
                highlighter.highlight(savedirectory)
                tkMessageBox.showinfo("Finished", "The highlighted report is saved in your chosen directory.")

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
bot_render = Tkinter.PhotoImage(file=help_bot)
logo_render = Tkinter.PhotoImage(file= y2_logo)
logo_label = Tkinter.Label(window, image=logo_render, borderwidth=0, highlightthickness=0, relief=Tkinter.FLAT, padx=50)
logo_label.pack(side=Tkinter.RIGHT)

window.option_add("*Font", ('Trade Gothic LT Pro', 16, ))
window.option_add("*Button.Foreground", "midnight blue")

header_font = ('Trade Gothic LT Pro', 18, 'bold')
header_color = '#112C4E'
def close(event):
    window.withdraw() # if you want to bring it back
    sys.exit() # if you want to exit the entire thing

window.bind('<Escape>', close)

window.deiconify()
Internbot(window)
window.mainloop()