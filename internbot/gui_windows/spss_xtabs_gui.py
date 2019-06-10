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



class SPSSCrosstabsView(object):

    def __init__(self, main_window, mov_x, mov_y, window_width, window_height, header_font, header_color):
        self.__window = main_window
        self.mov_x = mov_x
        self.mov_y = mov_y
        self.window_width = window_width
        self.window_height = window_height
        self.header_font = header_font
        self.header_color = header_color
        self.fpath = os.path.join(os.path.expanduser("~"), "Desktop")
        self.__embedded_fields = []

    def spss_crosstabs_menu(self):
        """
        Function sets up a menu for SPSS crosstabs.
        :return:
        """
        redirect_window = tkinter.Toplevel(self.__window)
        redirect_window.withdraw()
        width = 200
        height = 300
        redirect_window.geometry("%dx%d+%d+%d" % (width,height,self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        redirect_window.title("SPSS Menu")
        message = "Please select the\nfiles to produce."
        tkinter.Label(redirect_window, text = message, font=self.header_font, fg=self.header_color).pack()
        btn_var = tkinter.Button(redirect_window, text="Variable script", command=self.variable_script, height=3, width=20)
        btn_tab = tkinter.Button(redirect_window, text="Table script", command=self.table_script, height=3, width=20)
        btn_compile = tkinter.Button(redirect_window, text="Build report", command=self.build_xtabs, height=3, width=20)
        btn_cancel = tkinter.Button(redirect_window, text="Cancel", command=redirect_window.destroy, height=3, width=20)
        btn_cancel.pack(padx=5, side=tkinter.BOTTOM, expand=True)
        btn_compile.pack(padx=5, side=tkinter.BOTTOM, expand=True)
        btn_tab.pack(padx=5, side=tkinter.BOTTOM, expand=True)
        btn_var.pack(padx=5, side=tkinter.BOTTOM, expand=True)
        redirect_window.deiconify()

    def variable_script(self):
        """
        Set up for creation of Tables to run and rename variables for SPSS crosstabs.
        :return: None
        """
        print("Running variable script...")
        ask_qsf = messagebox.askokcancel("Select Qualtrics File", "Please select the Qualtrics survey .qsf file.")
        if ask_qsf is True: # user selected ok
            qsffilename = filedialog.askopenfilename(initialdir = self.fpath, title = "Select Qualtrics survey file",filetypes = (("Qualtrics file","*.qsf"),("all files","*.*")))
            if qsffilename is not "":
                compiler = base.QSFSurveyCompiler()
                survey = compiler.compile(qsffilename)
                ask_output = messagebox.askokcancel("Output directory", "Please select the directory for finished variable script.")
                if ask_output is True: # user selected ok
                    savedirectory = filedialog.askdirectory()
                    if savedirectory is not "":
                        variables = crosstabs.Format_SPSS_Report.Generate_Prelim_SPSS_Script.SPSSTranslator()
                        tables = crosstabs.Format_SPSS_Report.Generate_Prelim_SPSS_Script.TableDefiner()
                        variables.define_variables(survey, savedirectory)
                        tables.define_tables(survey, savedirectory)
                        open_files = messagebox.askyesno("Info", "Done!\nWould you like to open your finished files?")
                        if open_files is True:
                            self.open_file_for_user(savedirectory+"/Tables to run.csv")
                            self.open_file_for_user(savedirectory+"/rename variables.sps")


    def table_script(self):
        """
        Set up for Banner selection from a Tables to run file.
        :return:
        """
        print("Running table script...")
        ask_tables = messagebox.askokcancel("Select Tables to Run.csv File", "Please select the tables to run .csv file.")
        if ask_tables is True:
            self.tablesfilename = filedialog.askopenfilename(initialdir = self.fpath, title = "Select tables file",filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
            if self.tablesfilename is not "":
                ask_banners = messagebox.askokcancel("Banner selection", "Please insert/select the banners for this report.")
                if ask_banners is True:
                    names = crosstabs.Format_SPSS_Report.Generate_Table_Script.TablesParser().pull_table_names(self.tablesfilename)
                    titles = crosstabs.Format_SPSS_Report.Generate_Table_Script.TablesParser().pull_table_titles(self.tablesfilename)
                    bases = crosstabs.Format_SPSS_Report.Generate_Table_Script.TablesParser().pull_table_bases(self.tablesfilename)
                    self.banner_window(names, titles, bases)


    def banner_window(self, names, titles, bases):
        print("Inputting banners...")
        self.edit_window = tkinter.Toplevel(self.__window)
        self.edit_window.withdraw()
        width = 1500
        height = 500
        self.edit_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))

        self.edit_window.title("Banner selection")
        self.boxes_frame = tkinter.Frame(self.edit_window)
        self.boxes_frame.pack(side=tkinter.LEFT, fill=tkinter.BOTH)
        horiz_scrollbar = tkinter.Scrollbar(self.boxes_frame)
        horiz_scrollbar.config(orient=tkinter.HORIZONTAL)
        horiz_scrollbar.pack(side=tkinter.BOTTOM, fill=tkinter.X)
        vert_scrollbar = tkinter.Scrollbar(self.boxes_frame)
        vert_scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)

        self.tables_box = tkinter.Listbox(self.boxes_frame, selectmode="multiple", width=80, height=15,
                                          yscrollcommand=vert_scrollbar.set, xscrollcommand=horiz_scrollbar.set)
        self.tables_box.pack(padx=15, pady=10, expand=True, side=tkinter.LEFT, fill=tkinter.BOTH)
        self.banners_box = tkinter.Listbox(self.edit_window)
        self.banners_box.pack(padx=15, pady=10, expand=True, side=tkinter.RIGHT, fill=tkinter.BOTH)

        index = 0
        while index < len(names):
            self.tables_box.insert(tkinter.END, names[index] + ": " + titles[index])
            index += 1

        vert_scrollbar.config(command=self.tables_box.yview)
        horiz_scrollbar.config(command=self.tables_box.xview)

        buttons_frame = tkinter.Frame(self.edit_window)
        buttons_frame.pack(side=tkinter.RIGHT, fill=tkinter.BOTH)
        btn_up = tkinter.Button(buttons_frame, text="Up", command=self.shift_up)
        btn_down = tkinter.Button(buttons_frame, text="Down", command=self.shift_down)
        btn_insert = tkinter.Button(buttons_frame, text="Insert", command=self.insert_banner)
        btn_edit = tkinter.Button(buttons_frame, text="Edit", command=self.parse_selection)
        btn_create = tkinter.Button(buttons_frame, text="Create", command=self.create_banner)
        btn_remove = tkinter.Button(buttons_frame, text="Remove", command=self.remove_banner)
        btn_done = tkinter.Button(buttons_frame, text="Done", command=self.finish_banner)

        btn_done.pack(side=tkinter.BOTTOM, pady=15)
        btn_remove.pack(side=tkinter.BOTTOM)
        btn_create.pack(side=tkinter.BOTTOM)
        btn_edit.pack(side=tkinter.BOTTOM)
        btn_insert.pack(side=tkinter.BOTTOM, pady=5)
        btn_down.pack(side=tkinter.BOTTOM)
        btn_up.pack(side=tkinter.BOTTOM)

        self.edit_window.deiconify()



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
        self.tables_box.selection_clear(0, tkinter.END)
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
        self.tables_box.selection_clear(0, tkinter.END)
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
            self.banners_box.insert(tkinter.END, name + ": " + title)
            self.tables_box.itemconfig(index, {'fg': '#C00201'})
            self.tables_box.selection_clear(index, index)

    def create_banner(self, initial_name='', intial_title=''):
        create_window = tkinter.Toplevel(self.edit_window)
        create_window.title("Create banner")
        lbl_name = tkinter.Label(create_window, text="Banner name:")
        lbl_title = tkinter.Label(create_window, text="Banner title:")
        entry_name = tkinter.Entry(create_window)
        entry_name.insert(0, initial_name)
        entry_title = tkinter.Entry(create_window)
        entry_title.insert(0, intial_title)
        def create():
            name = entry_name.get()
            title = entry_title.get()
            self.banners_box.insert(tkinter.END, name + ": " + title)
            self.__embedded_fields.append(name)
            self.tables_box.insert(0, name + ": " + title)
            self.tables_box.itemconfig(0, {'fg': '#C00201'})
            create_window.destroy()

        createButton = tkinter.Button(create_window, text="Create", command=create)
        cancelButton = tkinter.Button(create_window, text="Cancel", command=create_window.destroy)

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
        edit_window = tkinter.Toplevel(self.edit_window)
        edit_window.title("Edit banner")
        lbl_banner = tkinter.Label(edit_window, text="Banner name:")
        entry_banner = tkinter.Entry(edit_window)
        entry_banner.insert(0, initial_names)
        lbl_title = tkinter.Label(edit_window, text="Title:")
        entry_title = tkinter.Entry(edit_window)
        entry_title.insert(0, initial_title)

        def edit():
            banner_name = entry_banner.get()
            title_name = entry_title.get()
            self.tables_box.delete(int(index))
            self.tables_box.insert(int(index), banner_name + ": " + title_name)
            edit_window.destroy()

        btn_cancel = tkinter.Button(edit_window, text="Cancel", command=edit_window.destroy)
        btn_edit = tkinter.Button(edit_window, text="Edit", command=edit)
        lbl_banner.grid(row = 0, column = 0)
        lbl_title.grid(row = 0, column = 1)
        entry_banner.grid(row = 1, column = 0)
        entry_title.grid(row = 1, column = 1)

        btn_edit.grid(row = 2, column = 0)
        btn_cancel.grid(row = 2, column = 1)

    def edit_banner(self, initial_names='', initial_title='', index=-1):
        edit_window = tkinter.Toplevel(self.edit_window)
        edit_window.title("Edit banner")
        lbl_banner = tkinter.Label(edit_window, text="Banner name:")
        entry_banner = tkinter.Entry(edit_window)
        entry_banner.insert(0, initial_names)
        lbl_title = tkinter.Label(edit_window, text="Title:")
        entry_title = tkinter.Entry(edit_window)
        entry_title.insert(0, initial_title)

        def edit():
            banner_name = entry_banner.get()
            title_name = entry_title.get()
            self.banners_box.delete(int(index))
            self.banners_box.insert(int(index), banner_name + ": " + title_name)
            edit_window.destroy()

        btn_cancel = tkinter.Button(edit_window, text="Cancel", command=edit_window.destroy)
        btn_edit = tkinter.Button(edit_window, text="Edit", command=edit)
        lbl_banner.grid(row=0, column=0)
        lbl_title.grid(row=0, column=1)
        entry_banner.grid(row=1, column=0)
        entry_title.grid(row=1, column=1)

        btn_edit.grid(row=2, column=0)
        btn_cancel.grid(row=2, column=1)

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

        self.variable_entry = None
        ask_trended = messagebox.askyesno("Trended Follow-up", "Is this a trended report?")
        if ask_trended is True:
            self.filter_variable_window()
        else:
            self.save_table_script()


    def filter_variable_window(self):
        self.create_window = tkinter.Toplevel(self.__window)
        self.create_window.withdraw()
        x = self.__window.winfo_x()
        y = self.__window.winfo_y()
        self.create_window.geometry("250x100+%d+%d" % (x + 175, y + 150))
        self.create_window.title("Trended report details")

        # filtering  variable
        variable_frame = tkinter.Frame(self.create_window)
        variable_frame.pack(side = tkinter.TOP, expand=True)

        variable_frame = tkinter.Label(variable_frame, text="Trended variable:")
        variable_frame.pack (side = tkinter.LEFT, expand=True)
        self.variable_entry = tkinter.Entry(variable_frame)
        self.variable_entry.pack(side=tkinter.RIGHT, expand=True)

        # done and cancel buttons
        button_frame = tkinter.Frame(self.create_window)
        button_frame.pack(side = tkinter.TOP, expand=True)

        btn_cancel = tkinter.Button(self.create_window, text = "Cancel", command = self.create_window.destroy)
        btn_cancel.pack(side = tkinter.RIGHT, expand=True)
        btn_done = tkinter.Button(self.create_window, text = "Done", command = self.save_table_script)
        btn_done.pack(side = tkinter.RIGHT, expand=True)
        self.create_window.deiconify()

    def save_table_script(self):

        generator = crosstabs.Format_SPSS_Report.Generate_Table_Script.TableScriptGenerator()
        table_order = OrderedDict()
        banner_list = OrderedDict()
        for item in list(self.tables_box.get(0, tkinter.END)):
            question = item.split(": ")
            table_order[question[0]] = question[1]
        for item in list(self.banners_box.get(0, tkinter.END)):
            question = item.split(": ")
            banner_list[question[0]] = question[1]
        self.edit_window.destroy()
        ask_output = messagebox.askokcancel("Output directory", "Please select the directory for finished table script.")
        filtering_variable = None
        if self.variable_entry is not None:
            variable = str(self.variable_entry.get())
            if variable != "":
                filtering_variable = variable
            self.create_window.destroy()
        if ask_output is True:
            savedirectory = filedialog.askdirectory()
            if savedirectory is not "":
                 generator.compile_scripts(self.tablesfilename, banner_list.keys(), self.__embedded_fields, filtering_variable, savedirectory)
                 open_files = messagebox.askyesno("Info","Done!\nWould you like to open your finished files?")
                 if open_files is True:
                     self.open_file_for_user(savedirectory + "/table script.sps")


    def build_xtabs(self):
        print("Building report...")
        ask_directory = messagebox.askokcancel("Select Tables Folder", "Please select the folder containing SPSS generated .xlsx table files.")
        if ask_directory is True:
            tablesdirectory = filedialog.askdirectory()
            builder = crosstabs.Format_SPSS_Report.Parse_SPSS_Tables.CrosstabGenerator(tablesdirectory)
            ask_output = messagebox.askokcancel("Output directory", "Please select the directory for finished report.")
            if ask_output is True:
                outputdirectory = filedialog.askdirectory()
                if outputdirectory is not "":
                    builder.write_report(outputdirectory)
                    open_files = messagebox.askyesno("Info", "Done!\nWould you like to open your finished files?")
                    if open_files is True:
                        self.open_file_for_user(outputdirectory + "/Crosstab Report.xlsx")

    def open_sound(self):

        def play_sound():
            audio_file = "/Library/internbot/1.0.0/templates_images/open.mp3"
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