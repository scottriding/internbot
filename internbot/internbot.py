import base
import crosstabs
import topline
import Tkinter
import tkMessageBox
import tkFileDialog
import os
from PIL import Image, ImageTk
from collections import OrderedDict

class Internbot:

    def __init__(self, root):
        self.__window = root
        self.main_buttons()
        self.fpath = os.path.join(os.path.expanduser("~"), "Desktop")
        self.__tables = OrderedDict()

    def main_buttons(self):
        btn_xtabs = Tkinter.Button(self.__window, text = "Run crosstabs", command = self.tabs_menu)
        btn_report = Tkinter.Button(self.__window, text = "Run topline report", command = self.topline_menu)
        btn_quit = Tkinter.Button(self.__window, text = "Quit", command = self.__window.destroy)
        btn_xtabs.pack(padx = 10, side = Tkinter.LEFT)
        btn_report.pack(padx = 10, side = Tkinter.LEFT)
        btn_quit.pack(padx = 10, side = Tkinter.LEFT)

    def tabs_menu(self):
        redirect_window = Tkinter.Tk()
        redirect_window.title("Y2 Crosstab Automation")
        message = "Please select the files to produce."
        Tkinter.Label(redirect_window, text = message).pack()
        btn_var = Tkinter.Button(redirect_window, text = "Variable script", command = self.variable_script)
        btn_tab = Tkinter.Button(redirect_window, text = "Table script", command = self.table_script)
        btn_high = Tkinter.Button(redirect_window, text = "Highlight", command = self.highlight)
        btn_cancel = Tkinter.Button(redirect_window, text = "Cancel", command = redirect_window.destroy)
        btn_var.pack(padx = 5, side = Tkinter.LEFT)
        btn_tab.pack(padx = 5, side = Tkinter.LEFT)
        btn_high.pack(padx = 5, side = Tkinter.LEFT)
        btn_cancel.pack(padx = 5, side = Tkinter.LEFT)

    def topline_menu(self):
        redirect_window = Tkinter.Toplevel()
        redirect_window.title("Y2 Topline Report Automation")
        message = "Please open a survey file."
        Tkinter.Label(redirect_window, text = message).pack()
        btn_open = Tkinter.Button(redirect_window, text = "Open", command = self.open_topline)
        btn_cancel = Tkinter.Button(redirect_window, text = "Cancel", command = redirect_window.destroy)
        btn_open.pack(padx = 15, side = Tkinter.LEFT)
        btn_cancel.pack(padx = 15, side = Tkinter.LEFT)

    def variable_script(self):
        ask_qsf = tkMessageBox.askokcancel("Select Qualtrics File", "Please select the Qualtrics survey .qsf file.")
        if ask_qsf is True: # user selected ok
            qsffilename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select Qualtrics survey file",filetypes = (("Qualtrics file","*.qsf"),("all files","*.*")))
            compiler = base.QSFSurveyCompiler()
            survey = compiler.compile(qsffilename)
            ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished variable script.")
            if ask_output is True: # user selected ok
                savedirectory = tkFileDialog.askdirectory()
                variables = crosstabs.Generate_Prelim_SPSS_Script.SPSSTranslator()
                tables = crosstabs.Generate_Prelim_SPSS_Script.TableDefiner()    
                variables.define_variables(survey, savedirectory)
                tables.define_tables(survey, savedirectory)

    def table_script(self):
        script = crosstabs.Generate_Table_Script.TableScript()
        ask_tables = tkMessageBox.askokcancel("Select Tables to Run.csv File", "Please select the tables to run .csv file.")
        if ask_tables is True:
            self.tablesfilename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select tables file",filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
            ask_banners = tkMessageBox.askokcancel("Banner selection", "Please insert/select the banners for this report.")
            if ask_banners is True:
                names = crosstabs.Generate_Table_Script.TablesParser().pull_table_names(self.tablesfilename)
                titles = crosstabs.Generate_Table_Script.TablesParser().pull_table_titles(self.tablesfilename)
                index = 0
                while index < len(names):
                    self.__tables[names[index]] = titles[index]
                    index += 1
                self.banner_window(names, titles)

    def banner_window(self, names, titles):
        self.edit_window = Tkinter.Toplevel()
        table_frame = Tkinter.Frame(self.edit_window)
        banner_frame = Tkinter.Frame(self.edit_window)

        table_frame.pack(fill=Tkinter.BOTH, expand=Tkinter.YES)
        banner_frame.pack(fill=Tkinter.BOTH, expand=Tkinter.YES)

        Tkinter.Label(table_frame, text = "Table name").pack()
        Tkinter.Label(banner_frame, text = "Banners").pack()

        table_frame.grid(row = 0, column = 0, rowspan = 15, sticky="nsew")
        banner_frame.grid(row = 0, column = 2, rowspan = 15, sticky="nsew")

        scrollbar_tables_vert = Tkinter.Scrollbar(table_frame, orient="vertical")
        scrollbar_tables_vert.pack(side=Tkinter.RIGHT, fill=Tkinter.Y)

        scrollbar_tables_horiz = Tkinter.Scrollbar(table_frame, orient="horizontal")
        scrollbar_tables_horiz.pack(side=Tkinter.BOTTOM, fill=Tkinter.X)

        self.tables_box = Tkinter.Listbox(table_frame, selectmode='multiple', width=80, height=15, \
                                         yscrollcommand=scrollbar_tables_vert.set, \
                                         xscrollcommand=scrollbar_tables_horiz.set)
        
        self.tables_box.pack(expand=True, fill=Tkinter.Y)

        scrollbar_tables_vert.config(command=self.tables_box.yview)
        scrollbar_tables_horiz.config(command=self.tables_box.xview)

        index = 0
        while index < len(names):
            self.tables_box.insert(Tkinter.END, names[index] + ": " + titles[index])
            index += 1

        scrollbar_banner_vert = Tkinter.Scrollbar(banner_frame, orient="vertical")
        scrollbar_banner_vert.pack(side=Tkinter.RIGHT, fill=Tkinter.Y)

        scrollbar_banner_horiz = Tkinter.Scrollbar(banner_frame, orient="horizontal")
        scrollbar_banner_horiz.pack(side=Tkinter.BOTTOM, fill=Tkinter.X)

        self.banners_box = Tkinter.Listbox(banner_frame, width=20, height=15, \
                                           yscrollcommand=scrollbar_banner_vert.set, \
                                           xscrollcommand=scrollbar_banner_horiz.set)

        self.banners_box.pack(expand=True, fill=Tkinter.Y)

        scrollbar_banner_vert.config(command=self.banners_box.yview)
        scrollbar_banner_horiz.config(command=self.banners_box.xview)

        btn_up = Tkinter.Button(self.edit_window, text = "Up", command = self.shift_banner_up).grid(row = 3, column = 1)
        btn_down = Tkinter.Button(self.edit_window, text = "Down", command = self.shift_banner_down).grid(row = 4, column = 1)
        btn_insert = Tkinter.Button(self.edit_window, text = "Insert", command = self.insert_banner).grid(row = 9, column = 1)
        btn_edit = Tkinter.Button(self.edit_window, text =   "Edit", command = self.parse_selection).grid(row = 10, column = 1)
        btn_create = Tkinter.Button(self.edit_window, text = "Create", command = self.create_banner).grid(row = 11, column = 1)
        btn_remove = Tkinter.Button(self.edit_window, text = "Remove", command = self.remove_banner).grid(row = 12, column = 1)

        btn_done = Tkinter.Button(self.edit_window, text = "Done", command = self.finish_banner).grid(row = 15, column = 3)

    def shift_banner_up(self):
        old_index = -1
        new_index = -1
        for index in self.tables_box.curselection():
            table = self.tables_box.get(int(index))
            old_index = index
            new_index = old_index - 1
            if old_index > -1 and new_index > -1:
                self.tables_box.delete(old_index)
                self.tables_box.insert(new_index, table)
                self.tables_box.selection_clear(old_index, old_index)
                self.tables_box.selection_set(new_index)

        old_index = -1
        new_index = -1
        for index in self.banners_box.curselection():
            table = self.banners_box.get(int(index))
            old_index = index
            new_index = old_index - 1
            if old_index > -1 and new_index > -1:
                self.banners_box.delete(old_index)
                self.banners_box.insert(new_index, table)
                self.banners_box.selection_clear(old_index, old_index)
                self.banners_box.selection_set(new_index)

    def shift_banner_down(self):
        old_index = -1
        new_index = -1
        for index in self.tables_box.curselection():
            table = self.tables_box.get(int(index))
            old_index = index
            new_index = old_index + 1
            if new_index < len(self.__tables):
                self.tables_box.delete(old_index)
                self.tables_box.insert(new_index, table)
                self.tables_box.selection_clear(old_index, old_index)
                self.tables_box.selection_set(new_index)

        old_index = -1
        new_index = -1
        for index in self.banners_box.curselection():
            table = self.banners_box.get(int(index))
            old_index = index
            new_index = old_index + 1
            if new_index < len(self.__tables):
                self.banners_box.delete(old_index)
                self.banners_box.insert(new_index, table)
                self.banners_box.selection_clear(old_index, old_index)
                self.banners_box.selection_set(new_index)

    def insert_banner(self):
        indexes_to_clear = []
        for index in self.tables_box.curselection():
            table = self.tables_box.get(int(index))
            question = table.split(": ")
            name = question[0]
            title = question[1]
            self.banners_box.insert(Tkinter.END, name + ": " + title)
            indexes_to_clear.append(index)
        indexes_to_clear.sort(reverse=True)
        for index in indexes_to_clear:
            self.tables_box.delete(index)

    def create_banner(self, initial_name='', intial_title=''):
        create_window = Tkinter.Toplevel()
        create_window.title("Create banner")
        lbl_name = Tkinter.Label(create_window, text="Banner name:")
        lbl_title = Tkinter.Label(create_window, text="Banner title:")
        entry_name = Tkinter.Entry(create_window)
        entry_name.insert(0, initial_name)
        entry_title = Tkinter.Entry(create_window)
        entry_title.insert(0, intial_title)
        def create():
            name = entry_name.get()
            title = entry_title.get()
            self.banners_box.insert(Tkinter.END, name + ": " + title)
            self.__tables[name] = title
            create_window.destroy()
        
        createButton = Tkinter.Button(create_window, text="Create", command=create)
        cancelButton = Tkinter.Button(create_window, text="Cancel", command=create_window.destroy)

        lbl_name.grid(row=0, column=0)
        lbl_title.grid(row=0, column=1)
        entry_name.grid(row=1, column=0)
        entry_title.grid(row=1, column=1)
        createButton.grid(row=2, column=0)
        cancelButton.grid(row=2, column=1)

    def parse_selection(self):
        name = ''
        title = ''
        for index in self.tables_box.curselection():
            table = self.tables_box.get(int(index))
            question = table.split(": ")
            name = question[0]
            title = question[1]
        if name is not '':
            self.edit_table(name, title, index)
        else:
            for index in self.banners_box.curselection():
                banner = self.banners_box.get(int(index))
                question = banner.split(": ")
                name = question[0]
                title = question[1]
                self.edit_banner(name, title, index)

    def edit_table(self, initial_names='', initial_title='', index=-1):
        edit_window = Tkinter.Toplevel()
        edit_window.title("Edit banner")
        lbl_banner = Tkinter.Label(edit_window, text="Banner name:")
        entry_banner = Tkinter.Entry(edit_window)
        entry_banner.insert(0, initial_names)
        lbl_title = Tkinter.Label(edit_window, text="Title:")
        entry_title = Tkinter.Entry(edit_window)
        entry_title.insert(0, initial_title)

        def edit():
            banner_name = entry_banner.get()
            title_name = entry_title.get()
            self.tables_box.delete(int(index))
            self.tables_box.insert(int(index), banner_name + ": " + title_name)
            edit_window.destroy()

        btn_cancel = Tkinter.Button(edit_window, text="Cancel", command=edit_window.destroy)
        btn_edit = Tkinter.Button(edit_window, text="Edit", command=edit)
        lbl_banner.grid(row = 0, column = 0)
        lbl_title.grid(row = 0, column = 1)
        entry_banner.grid(row = 1, column = 0)
        entry_title.grid(row = 1, column = 1)

        btn_edit.grid(row = 2, column = 0)
        btn_cancel.grid(row = 2, column = 1)

    def edit_banner(self, initial_names='', initial_title='', index=-1):
        edit_window = Tkinter.Toplevel()
        edit_window.title("Edit banner")
        lbl_banner = Tkinter.Label(edit_window, text="Banner name:")
        entry_banner = Tkinter.Entry(edit_window)
        entry_banner.insert(0, initial_names)
        lbl_title = Tkinter.Label(edit_window, text="Title:")
        entry_title = Tkinter.Entry(edit_window)
        entry_title.insert(0, initial_title)

        def edit():
            banner_name = entry_banner.get()
            title_name = entry_title.get()
            self.banners_box.delete(int(index))
            self.banners_box.insert(int(index), banner_name + ": " + title_name)
            edit_window.destroy()

        btn_cancel = Tkinter.Button(edit_window, text="Cancel", command=edit_window.destroy)
        btn_edit = Tkinter.Button(edit_window, text="Edit", command=edit)
        lbl_banner.grid(row = 0, column = 0)
        lbl_title.grid(row = 0, column = 1)
        entry_banner.grid(row = 1, column = 0)
        entry_title.grid(row = 1, column = 1)

        btn_edit.grid(row = 2, column = 0)
        btn_cancel.grid(row = 2, column = 1)

    def remove_banner(self):
        for index in self.tables_box.curselection():
            self.tables_box.delete(int(index))
        for index in self.banners_box.curselection():
            item = self.banners_box.get(int(index))
            self.banners_box.delete(int(index))
            question = item.split(": ")
            count = 0
            for key in self.__tables:
                if key in question[0]:
                    index = count
                else:
                    count += 1
            self.tables_box.insert(index, question[0] + ": " + question[1])

    def finish_banner(self):
        to_return = OrderedDict()
        for item in list(self.banners_box.get(0, Tkinter.END)):
            question = item.split(": ")
            to_return[question[0]] = question[1]
        self.edit_window.destroy()
        banners = []
        banners = to_return.keys()
        ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished table script.")
        if ask_output is True:
            savedirectory = tkFileDialog.askdirectory()
            crosstabs.Generate_Table_Script.TableScript().compile_scripts(self.tablesfilename, savedirectory, banners)

    def highlight(self):
        ask_xlsx = tkMessageBox.askokcancel("Select Tables Microsoft Excel File", "Please select the combined table .xlsx file.")
        if ask_xlsx is True:
            tabsfilename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select combined Crosstabs report.",filetypes = (("Microsoft Excel","*.xlsx"),("all files","*.*")))
            renamer = crosstabs.Polish_Final_Report.RenameTabs()
            ask_tables = tkMessageBox.askokcancel("Select Tables .csv File", "Please select the Tables to run.csv file.")
            if ask_tables is True:
                tablesfilename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select Comma Seperated file",filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
                ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished table script.")
                if ask_output is True:
                    savedirectory = tkFileDialog.askdirectory()
                    renamer.rename(tabsfilename, tablesfilename, savedirectory)
                    highlighter = crosstabs.Polish_Final_Report.Highlighter(savedirectory)
                    highlighter.highlight(savedirectory)
                    tkMessageBox.showinfo("Finished", "The highlighted report is saved as \"Highlighted.xlsx\" in your chosen directory.")

    def open_topline(self):
        filename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select survey file",filetypes = (("Qualtrics files","*.qsf"),("comma seperated files","*.csv"),("all files","*.*")))
        isQSF = False
        if ".qsf" in filename:
            compiler = base.QSFSurveyCompiler()
            survey = compiler.compile(filename)
            report = topline.QSF.ReportGenerator(survey)
            isQSF = True
            build_report(isQSF, report)
        elif ".csv" in filename:
            report = topline.CSV.ReportGenerator(filename)
            build_report(isQSF, report)

    def build_report(isQSF, report):
        template_file = open("topline_template.docx", "r")
        ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report.")
        if ask_output is True:
            savedirectory = tkFileDialog.askdirectory()
            if isQSF is True:
                ask_freq = tkMessageBox.askokcancel("Frequency file", "Please select the topline .csv frequency file.")
                if ask_freq is True:
                    freqfilename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select frequency file",filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
                    report.generate_basic_topline(freqfilename, template_file, savedirectory)
            else:
                report.generate_basic_topline(template_file, savedirectory)

window = Tkinter.Tk()
window.title("Internbot: 01011001 00000010") # Internbot: Y2
y2_logo = Image.open("y2Logo.jpg")
render = ImageTk.PhotoImage(y2_logo)
Tkinter.Label(window, image=render).pack()
Internbot(window)
window.mainloop()