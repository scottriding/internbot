import base
import crosstabs
import topline
import rnc_automation
import Tkinter
import tkMessageBox
import tkFileDialog
import os
import csv
from PIL import Image, ImageTk
from collections import OrderedDict

class Internbot:

    def __init__(self, root):
        self.__window = root
        self.main_buttons()
        self.fpath = os.path.join(os.path.expanduser("~"), "Desktop")
        self.__embedded_fields = []

    def main_buttons(self):
        btn_xtabs = Tkinter.Button(self.__window, text = "Run crosstabs", command = self.tabs_menu)
        btn_report = Tkinter.Button(self.__window, text = "Run topline report", command = self.topline_menu)
        btn_rnc = Tkinter.Button(self.__window, text = "Run RNC", command = self.rnc_menu)
        btn_quit = Tkinter.Button(self.__window, text = "Quit", command = self.__window.destroy)
        btn_xtabs.pack(padx = 5, side = Tkinter.LEFT, expand=True)
        btn_report.pack(padx = 5, side = Tkinter.LEFT, expand=True)
        btn_rnc.pack(padx = 5, side = Tkinter.LEFT, expand=True)
        btn_quit.pack(padx = 5, side = Tkinter.LEFT, expand=True)

    def tabs_menu(self):
        redirect_window = Tkinter.Toplevel(self.__window)
        message = "Please select the files to produce."
        Tkinter.Label(redirect_window, text = message).pack()
        btn_var = Tkinter.Button(redirect_window, text = "Variable script", command = self.variable_script)
        btn_tab = Tkinter.Button(redirect_window, text = "Table script", command = self.table_script)
        btn_high = Tkinter.Button(redirect_window, text = "Highlight", command = self.highlight)
        btn_cancel = Tkinter.Button(redirect_window, text = "Cancel", command = redirect_window.destroy)
        btn_var.pack(padx = 5, side = Tkinter.LEFT, expand=True)
        btn_tab.pack(padx = 5, side = Tkinter.LEFT, expand=True)
        btn_high.pack(padx = 5, side = Tkinter.LEFT, expand=True)
        btn_cancel.pack(padx = 5, side = Tkinter.LEFT, expand=True)

    def topline_menu(self):
        redirect_window = Tkinter.Toplevel(self.__window)
        redirect_window.title("Y2 Topline Report Automation")
        message = "Please open a survey file."
        Tkinter.Label(redirect_window, text = message).pack(expand=True)
        btn_open = Tkinter.Button(redirect_window, text = "Open", command = self.open_topline)
        btn_cancel = Tkinter.Button(redirect_window, text = "Cancel", command = redirect_window.destroy)
        btn_open.pack(ipadx = 10, side = Tkinter.LEFT, expand=True)
        btn_cancel.pack(ipadx = 10, side = Tkinter.LEFT, expand=True)

    def rnc_menu(self):
        redirect_window = Tkinter.Toplevel(self.__window)
        redirect_window.title("RNC Scores Topline Automation")
        message = "Please open a model scores file."
        Tkinter.Label(redirect_window, text = message).pack(expand=True)
        btn_topline = Tkinter.Button(redirect_window, text = "Score Topline", command = self.scores_topline)
        btn_trended = Tkinter.Button(redirect_window, text = "Issue Trended", command = self.issue_trended)
        brn_ind_trended = Tkinter.Button(redirect_window, text = "Trended Score", command = self.trended_scores)
        btn_cancel = Tkinter.Button(redirect_window, text = "Cancel", command = redirect_window.destroy)
        btn_topline.pack(ipadx = 5, side = Tkinter.LEFT, expand=True)
        btn_trended.pack(ipadx = 5, side = Tkinter.LEFT, expand=True)
        brn_ind_trended.pack(ipadx = 5, side = Tkinter.LEFT, expand=True)
        btn_cancel.pack(ipadx = 5, side = Tkinter.LEFT, expand=True)

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
                bases = crosstabs.Generate_Table_Script.TablesParser().pull_table_bases(self.tablesfilename)
                self.banner_window(names, titles, bases)

    def banner_window(self, names, titles, bases):
        self.edit_window = Tkinter.Toplevel(self.__window)
        self.edit_window.title("Banner selection")

        titles_frame = Tkinter.Frame(self.edit_window)
        titles_frame.pack()

        #w = Tkinter.Label(self.edit_window, text="Red", bg="red", fg="white")
        #w.pack(side=Tkinter.LEFT)
        #w = Tkinter.Label(self.edit_window, text="Green", bg="green", fg="black")
        #w.pack(side=Tkinter.LEFT)
        #w = Tkinter.Label(self.edit_window, text="Blue", bg="blue", fg="white")
        #w.pack(side=Tkinter.LEFT)

        self.boxes_frame = Tkinter.Frame(self.edit_window)
        self.boxes_frame.pack(fill=Tkinter.BOTH)

        self.tables_box = Tkinter.Listbox(self.edit_window, selectmode="multiple", width=80, height=15)

        self.tables_box.pack(padx = 15, pady=10,expand=True, side = Tkinter.LEFT, fill=Tkinter.BOTH)

        self.banners_box = Tkinter.Listbox(self.edit_window)
        self.banners_box.pack(padx = 15, pady=10, expand=True, side=Tkinter.RIGHT, fill=Tkinter.BOTH)

        index = 0
        while index < len(names):
            self.tables_box.insert(Tkinter.END, names[index] + ": " + titles[index])
            index += 1

        btn_up = Tkinter.Button(self.edit_window, text = "Up", command = self.shift_up)
        btn_down = Tkinter.Button(self.edit_window, text = "Down", command = self.shift_down)
        btn_insert = Tkinter.Button(self.edit_window, text = "Insert", command = self.insert_banner)
        btn_edit = Tkinter.Button(self.edit_window, text =   "Edit", command = self.parse_selection)
        btn_create = Tkinter.Button(self.edit_window, text = "Create", command = self.create_banner)
        btn_remove = Tkinter.Button(self.edit_window, text = "Remove", command = self.remove_banner)

        btn_done = Tkinter.Button(self.edit_window, text = "Done", command = self.finish_banner)

        btn_done.pack(side=Tkinter.BOTTOM, pady=15)
        btn_remove.pack(side=Tkinter.BOTTOM)
        btn_create.pack(side=Tkinter.BOTTOM)
        btn_edit.pack(side=Tkinter.BOTTOM)

        btn_insert.pack(side=Tkinter.BOTTOM, pady=5)
        btn_down.pack(side=Tkinter.BOTTOM)
        btn_up.pack(side=Tkinter.BOTTOM)

    def shift_up(self):
        current_banners = []
        for banner in self.banners_box.curselection():
            current_banners.append(banner)

        if len(current_banners) > 0:
            self.shift_banners_up()

        current_tables = []
        for table in self.tables_box.curselection():
            current_tables.append(table)

        if len(current_tables) > 0:
            self.shift_tables_up(current_tables)

    def shift_banners_up(self):
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

    def shift_tables_up(self, current_tables):
        shifted_indexes = []
        index = 0
        while index < len(current_tables):
            selection_index = current_tables[index]
            new_index = selection_index - 1
            if new_index > -1:
                if new_index not in shifted_indexes:
                    shifted_indexes.append(new_index)
                else:
                    shifted_indexes.append(selection_index)
            else:
                shifted_indexes.append(selection_index)
            index += 1
        
        iteration = 0
        self.tables_box.selection_clear(0, Tkinter.END)
        for table_index in current_tables:
            is_highlighted = False
            current_table = self.tables_box.get(table_index)
            if '#C00201' in self.tables_box.itemconfig(table_index)['foreground']:
                is_highlighted = True
            self.tables_box.delete(table_index)
            self.tables_box.insert(shifted_indexes[iteration], current_table)
            if is_highlighted:
                self.tables_box.itemconfig(shifted_indexes[iteration], {'fg': '#C00201'})
            self.tables_box.selection_set(shifted_indexes[iteration])
            iteration += 1

    def shift_down(self):
        current_banners = []
        for banner in self.banners_box.curselection():
            current_banners.append(banner)

        if len(current_banners) > 0:
            self.shift_banners_down()

        current_tables = []
        for table in self.tables_box.curselection():
            current_tables.append(table)

        if len(current_tables) > 0:
            self.shift_tables_down(current_tables)
        
    def shift_banners_down(self):
        old_index = -1
        new_index = -1
        for index in self.banners_box.curselection():
            table = self.banners_box.get(int(index))
            old_index = index
            new_index = old_index + 1
            if new_index < self.banners_box.size():
                self.banners_box.delete(old_index)
                self.banners_box.insert(new_index, table)
                self.banners_box.selection_clear(old_index, old_index)
                self.banners_box.selection_set(new_index)

    def shift_tables_down(self, current_tables):
        shifted_indexes = []
        index = len(current_tables) - 1
        while index >= 0:
            selection_index = current_tables[index]
            new_index = selection_index + 1
            if new_index < self.tables_box.size():
                if new_index not in shifted_indexes:
                    shifted_indexes.append(new_index)
                else:
                    shifted_indexes.append(selection_index)
            else:
                shifted_indexes.append(selection_index)
            index -= 1

        current_tables.sort(reverse=True)
        shifted_indexes.sort(reverse=True)

        iteration = 0
        self.tables_box.selection_clear(0, Tkinter.END)
        for table_index in current_tables:
            is_highlighted = False
            current_table = self.tables_box.get(table_index)
            if '#C00201' in self.tables_box.itemconfig(table_index)['foreground']:
                is_highlighted = True
            self.tables_box.delete(table_index)
            self.tables_box.insert(shifted_indexes[iteration], current_table)
            if is_highlighted:
                self.tables_box.itemconfig(shifted_indexes[iteration], {'fg': '#C00201'})
            self.tables_box.selection_set(shifted_indexes[iteration])
            iteration += 1

    def insert_banner(self):
        indexes_to_clear = []
        for index in self.tables_box.curselection():
            table = self.tables_box.get(int(index))
            question = table.split(": ")
            name = question[0]
            title = question[1]
            self.banners_box.insert(Tkinter.END, name + ": " + title)
            self.tables_box.itemconfig(index, {'fg': '#C00201'})
            self.tables_box.selection_clear(index, index)
        
    def create_banner(self, initial_name='', intial_title=''):
        create_window = Tkinter.Toplevel(self.edit_window)
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
            self.__embedded_fields.append(name)
            self.tables_box.insert(0, name + ": " + title)
            self.tables_box.itemconfig(0, {'fg': '#C00201'})
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
        edit_window = Tkinter.Toplevel(self.edit_window)
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
        edit_window = Tkinter.Toplevel(self.edit_window)
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
        indexes_to_delete = []
        for index in self.tables_box.curselection():
            if '#C00201' not in self.tables_box.itemconfig(index)['foreground']:
                item = self.tables_box.get(index)
                indexes_to_delete.append(index)
                question = item.split(": ")
        indexes_to_delete.sort(reverse=True)
        for index in indexes_to_delete:
            self.tables_box.delete(int(index))
        for index in self.banners_box.curselection():
            item = self.banners_box.get(int(index))
            self.banners_box.delete(int(index))
            banner_to_remove = item.split(": ")
            index = 0
            while index < self.tables_box.size():
                to_compare = self.tables_box.get(index).split(": ")
                if to_compare[0] == banner_to_remove[0]:
                    self.tables_box.itemconfig(index, {'fg': '#000000'})
                index += 1

    def finish_banner(self):
        table_order = OrderedDict()
        banner_list = OrderedDict()
        for item in list(self.tables_box.get(0, Tkinter.END)):
            question = item.split(": ")
            table_order[question[0]] = question[1]
        for item in list(self.banners_box.get(0, Tkinter.END)):
            question = item.split(": ")
            banner_list[question[0]] = question[1]
        self.edit_window.destroy()
        banners = banner_list.keys()
        ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished table script.")
        if ask_output is True:
            savedirectory = tkFileDialog.askdirectory()
            crosstabs.Generate_Table_Script.TableScript().compile_scripts(self.tablesfilename, savedirectory, banners, self.__embedded_fields)
            self.reorder_tablesfile(savedirectory, table_order)

    def reorder_tablesfile(self, savedirectory, table_order):
        pass

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
            self.build_report(isQSF, report)
        elif ".csv" in filename:
            report = topline.CSV.ReportGenerator(filename)
            self.build_report(isQSF, report)

    def build_report(self, isQSF, report):
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

    def scores_topline(self):
        ask_xlsx_file = tkMessageBox.askyesno("Previous rounds", "Do you already have a topline score file and need to add a round?")
        if ask_xlsx_file is "Yes":
            pass
        else:
            filename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select model file", filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
            report = rnc_automation.ScoresToplineReportGenerator(filename)
            ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report.")
            if ask_output is True:
                savedirectory = tkFileDialog.askdirectory()
                if savedirectory is not "":
                    report.generate_scores_topline("Montana", savedirectory)

    def issue_trended(self):
        ask_xlsx_file = tkMessageBox.askyesno("Previous rounds", "Do you already have a trended file and need to add a round?")
        if ask_xlsx_file is "Yes":
            pass
        else:
            filename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select model file", filetypes = (("comma seperated files", "*.csv"), ("all files", "*.*")))
            report = rnc_automation.IssueTrendedReportGenerator(filename)
            ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report.")
            if ask_output is True:
                savedirectory = tkFileDialog.askdirectory()
                if savedirectory is not "":
                    report.generate_issue_trended(savedirectory)

    def trended_scores(self):
        filename = tkFileDialog.askopenfilename(initialdir = self.fpath, title = "Select model file", filetypes = (("comma seperated files", "*.csv"), ("all files", "*.*")))
        report = rnc_automation.TrendedScoresReportGenerator(filename)
        ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report.")
        if ask_output is True:
            savedirectory = tkFileDialog.askdirectory()
            if savedirectory is not "":
                report.generate_trended_scores(savedirectory)

window = Tkinter.Tk()
window.title("Internbot: 01011001 00000010") # Internbot: Y2
y2_logo = Image.open("y2Logo.jpg")
render = ImageTk.PhotoImage(y2_logo)
Tkinter.Label(window, image=render).pack()
Internbot(window)
window.mainloop()