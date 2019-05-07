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
import threading
import sys

class AppendixView(object):

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

    def append_menu(self):
        """
        Function sets up Topline Appendix menu.
        :return: None
        """
        self.redirect_window=Tkinter.Toplevel(self.__window)
        self.redirect_window.withdraw()
        width=200
        height=300

        self.redirect_window.geometry("%dx%d+%d+%d" % (width,height,self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        self.redirect_window.title("Y2 Topline Appendix Report Automation")
        message = "Please open a .csv file \nwith open ended responses."
        Tkinter.Label(self.redirect_window, text=message, font=self.header_font, fg=self.header_color).pack(expand=True, pady=5)
        btn_open = Tkinter.Button(self.redirect_window, text="Open", command=self.append_type, height=3, width=20)
        btn_cancel = Tkinter.Button(self.redirect_window, text="Cancel", command=self.redirect_window.destroy, height=3,
                                    width=20)
        btn_bot = Tkinter.Button(self.redirect_window, image=self.bot_render, borderwidth=0,
                                 highlightthickness=0, relief=Tkinter.FLAT, bg="white", height=65, width=158,
                                 command=self.appendix_help_window)
        btn_bot.pack(side=Tkinter.TOP, padx=10, pady=5)
        btn_cancel.pack(side=Tkinter.BOTTOM, padx=10, pady=5)
        btn_open.pack(side=Tkinter.BOTTOM, padx=10, pady=5)

        self.redirect_window.deiconify()

    def append_type(self):
        """
        Funtion reads in the appendix file and creates a .docx
        :return: None
        """
        self.choose_window = Tkinter.Toplevel(self.__window)
        self.choose_window.withdraw()
        width = 250
        height = 300
        self.choose_window.geometry("%dx%d+%d+%d" % (
            width, height, self.mov_x + self.window_width / 2 - width / 2,
            self.mov_y + self.window_height / 2 - height / 2))
        message = "Choose the format of\nyour appendix report"
        Tkinter.Label(self.choose_window, text=message, font=self.header_font, fg=self.header_color).pack(expand=True)
        btn_doc = Tkinter.Button(self.choose_window, text="Word Appendix", command=self.doc_appendix,
                                     height=3, width=20)

        btn_excel = Tkinter.Button(self.choose_window, text="Excel Appendix",
                                   command=self.excel_appendix, height=3, width=20)
        btn_cancel = Tkinter.Button(self.choose_window, text="Cancel",
                                   command=self.choose_window.destroy, height=3, width=20)

        btn_cancel.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        btn_excel.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        btn_doc.pack(ipadx=5, side=Tkinter.BOTTOM, expand=False)
        self.choose_window.deiconify()





    def doc_appendix(self):
        generator = topline.Appendix.AppendixGenerator()
        csvfilename = tkFileDialog.askopenfilename(initialdir=self.fpath, title="Select open ends file",
                                                   filetypes=(("Comma separated files", "*csv"), ("all files", "*.*")))
        generator.parse_file(csvfilename)
        savedirectory = tkFileDialog.asksaveasfilename(defaultextension='.docx', filetypes=[('word files', '.docx')])
        if savedirectory is not "":
            generator.write_appendix(savedirectory, "templates_images/appendix_template.docx", False)
        self.redirect_window.destroy()

    def excel_appendix(self):
        generator = topline.Appendix.AppendixGenerator()
        csvfilename = tkFileDialog.askopenfilename(initialdir=self.fpath, title="Select open ends file",
                                                   filetypes=(("Comma separated files", "*csv"), ("all files", "*.*")))
        generator.parse_file(csvfilename)
        savedirectory = tkFileDialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('excel files', '.xlsx')])
        if savedirectory is not "":
            generator.write_appendix(savedirectory, '', True)
        self.redirect_window.destroy()

    def appendix_help_window(self):
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
        info_message = "\nThings"

        Tkinter.Label(help_window, text=info_message, justify=Tkinter.LEFT).pack(side=Tkinter.TOP)
        btn_ok = Tkinter.Button(help_window, text="Ok", command=help_window.destroy, height=3, width=20)
        btn_ok.pack(padx=5, pady=10, side=Tkinter.TOP, expand=False)
        help_window.deiconify()


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
