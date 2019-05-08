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

class QCrosstabsView(object):

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
        self.bases_history = ["Select from history"]

    def bases_window(self):
        """
        Function sets up window for bases entry of Q crosstab reports
        :return:
        """
        print "File Q"
        self.base_window = Tkinter.Toplevel(self.__window)
        self.base_window.withdraw()
        width = 1100
        height = 420
        self.base_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        self.base_window.title("Base assignment")
        self.loaded_qfile = False

        self.boxes_frame = Tkinter.Frame(self.base_window)
        self.boxes_frame.pack(padx=10, pady=10, side=Tkinter.LEFT, fill=Tkinter.BOTH)
        tables_horiz_scrollbar = Tkinter.Scrollbar(self.boxes_frame)
        tables_horiz_scrollbar.config(orient=Tkinter.HORIZONTAL)
        tables_horiz_scrollbar.pack(side=Tkinter.BOTTOM, fill=Tkinter.X)
        tables_vert_scrollbar = Tkinter.Scrollbar(self.boxes_frame)
        tables_vert_scrollbar.pack(side=Tkinter.RIGHT, fill=Tkinter.Y)

        self.tables_box = Tkinter.Listbox(self.boxes_frame, selectmode=Tkinter.SINGLE, width=55, height=10,
                                          yscrollcommand=tables_vert_scrollbar.set,
                                          xscrollcommand=tables_horiz_scrollbar.set, exportselection=False)
        self.tables_box.pack(expand=True, side=Tkinter.LEFT, fill=Tkinter.BOTH)

        self.bases_frame = Tkinter.Frame(self.base_window)
        self.bases_frame.pack(padx=10, pady=10, side=Tkinter.RIGHT, fill=Tkinter.BOTH)
        bases_horiz_scrollbar = Tkinter.Scrollbar(self.bases_frame)
        bases_horiz_scrollbar.config(orient=Tkinter.HORIZONTAL)
        bases_horiz_scrollbar.pack(side=Tkinter.BOTTOM, fill=Tkinter.X)
        bases_vert_scrollbar = Tkinter.Scrollbar(self.bases_frame)
        bases_vert_scrollbar.pack(side=Tkinter.RIGHT, fill=Tkinter.Y)

        self.bases_box = Tkinter.Listbox(self.bases_frame, selectmode=Tkinter.SINGLE, width=54, height=10,
                                         yscrollcommand=bases_vert_scrollbar.set,
                                         xscrollcommand=bases_horiz_scrollbar.set, exportselection=False)
        self.bases_box.pack(expand=False, side=Tkinter.RIGHT, fill=Tkinter.BOTH)

        self.focus_index = 0

        self.tables_box.select_set(self.focus_index)
        self.bases_box.select_set(self.focus_index)

        tables_vert_scrollbar.config(command=self.tables_box.yview)
        tables_horiz_scrollbar.config(command=self.tables_box.xview)
        bases_vert_scrollbar.config(command=self.bases_box.yview)
        bases_horiz_scrollbar.config(command=self.bases_box.xview)

        buttons_frame = Tkinter.Frame(self.base_window)
        buttons_frame.pack(pady=10, side=Tkinter.RIGHT, fill=Tkinter.BOTH)
        self.btn_edit = Tkinter.Button(buttons_frame, text="Edit", command=self.parse_bases, width=20, height=3)
        self.btn_done = Tkinter.Button(buttons_frame, text="Done", command=self.finish_bases, width=20, height=3)
        self.btn_open = Tkinter.Button(buttons_frame, text="Open", command=self.qtab_build, width=20, height=3)
        self.btn_load = Tkinter.Button(buttons_frame, text="Load CSV", command=self.load_csv, width=20, height=3)
        self.btn_cancel = Tkinter.Button(buttons_frame, text="Cancel", command=self.base_window.destroy, width=20,
                                         height=3)
        self.btn_bot = Tkinter.Button(buttons_frame, image=self.bot_render, borderwidth=0,
                                      highlightthickness=0, relief=Tkinter.FLAT, bg="white", height=65, width=158,
                                      command=self.bases_help_window)
        self.btn_cancel.pack(pady=2, side=Tkinter.BOTTOM)
        self.btn_done.pack(pady=2, side=Tkinter.BOTTOM)
        self.btn_edit.pack(pady=2, side=Tkinter.BOTTOM)
        self.btn_load.pack(pady=2, side=Tkinter.BOTTOM)
        self.btn_open.pack(pady=2, side=Tkinter.BOTTOM)
        self.btn_bot.pack(pady=2, side=Tkinter.BOTTOM)

        self.btn_done.config(state=Tkinter.DISABLED)
        self.btn_load.config(state=Tkinter.DISABLED)
        self.btn_edit.config(state=Tkinter.DISABLED)

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
                self.focus_index -= 1
            self.tables_box.select_set(self.focus_index)
            self.bases_box.select_set(self.focus_index)

        def down_pressed(event):
            for index in self.tables_box.curselection():
                self.tables_box.select_clear(index)

            for index in self.bases_box.curselection():
                self.bases_box.select_clear(index)

            if self.focus_index < self.tables_box.size() - 1:
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
            report_file_name = tkFileDialog.askopenfilename(initialdir=self.fpath, title="Select report file",
                                                            filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
            if report_file_name is not "":
                self.__parser = crosstabs.Format_Q_Report.QParser(report_file_name)
                self.tables = self.__parser.format_report()
                prompts = []
                for key, value in self.tables.iteritems():
                    prompts.append(value.prompt)

                for prompt in prompts:
                    self.tables_box.insert(Tkinter.END, prompt)
                base_sizes = ["[Enter base description] ~ [Enter base size]"] * len(prompts)
                for size in base_sizes:
                    self.bases_box.insert(Tkinter.END, size)

                self.focus_index = 0

                self.tables_box.select_set(self.focus_index)
                self.bases_box.select_set(self.focus_index)

                self.loaded_qfile = True
                self.base_window.focus_force()
                self.btn_done.config(state=Tkinter.NORMAL)
                self.btn_load.config(state=Tkinter.NORMAL)
                self.btn_edit.config(state=Tkinter.NORMAL)
        else:
            ask_lost_work = tkMessageBox.askyesno("Select Q Research report file",
                                                  "You will lose your work. \nDo you want to continue?")
            if ask_lost_work:
                self.loaded_qfile = False
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
        width = 300
        height = 400
        help_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        message = "\nBase Input Help"
        Tkinter.Label(help_window, text=message, font=self.header_font, fg=self.header_color).pack()
        info_message = "Select Open to import a Q output file." \
                       "\nThis should be a .xlsx file." \
                       "\n\nYou can change the current selection with" \
                       "\nthe up/down arrow keys or your mouse." \
                       "\n\nEdit the base and n info of the current" \
                       "\nselection by selecting Edit or hitting enter." \
                       "\n\nWhen you're done, select Done and you" \
                       "\nwill be prompted to select the directory of" \
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

    def unicode_dict_reader(self, utf8_data, **kwargs):
        csv_reader = csv.DictReader(utf8_data, **kwargs)
        for row in csv_reader:
            if row['base description'] != "":
                yield {unicode(key, 'iso-8859-1'): unicode(value, 'iso-8859-1') for key, value in row.iteritems()}

    def load_csv(self):
        base_file_name = tkFileDialog.askopenfilename(initialdir=self.fpath, title="Select base file",
                                                      filetypes=(
                                                      ("comma separated files", "*.csv"), ("all files", "*.*")))
        if base_file_name is not "":
            base_details = self.unicode_dict_reader(open(base_file_name))
            index = 0
            for base_detail in base_details:
                if index < len(self.tables):
                    base = base_detail["base description"]
                    n = base_detail["base size"]
                    self.bases_box.delete(int(index))
                    self.bases_box.insert(int(index), base + " ~ " + n)
                    index += 1

    def edit_base_window(self, index):
        """
        Function defines a window for the user to input Base and N for a table.
        :param base:
        :param index:
        :return:
        """
        self.edit_window = Tkinter.Toplevel(self.base_window)
        width = 400
        height = 150
        self.edit_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        self.edit_window.title("Add base and n")
        edit_frame = Tkinter.Frame(self.edit_window)
        edit_frame.pack(side=Tkinter.TOP)
        entry_frame = Tkinter.Frame(edit_frame)
        entry_frame.pack(side=Tkinter.TOP)
        base_frame = Tkinter.Frame(entry_frame)
        base_frame.pack(side=Tkinter.LEFT, padx=5, pady=5)
        lbl_base = Tkinter.Label(base_frame, text="Base description:")
        entry_des = Tkinter.Entry(base_frame, width=30)
        spinbox_frame = Tkinter.Frame(edit_frame)
        spinbox_frame.pack(side=Tkinter.BOTTOM)
        entry_des.focus_set()
        size_frame = Tkinter.Frame(entry_frame)
        size_frame.pack(side=Tkinter.RIGHT, padx=5, pady=5)
        lbl_title = Tkinter.Label(size_frame, text="Size (n):")
        entry_size = Tkinter.Entry(size_frame, width=20)

        current_base = self.bases_box.get(int(index)).split(" ~ ")

        entry_des.insert(0, current_base[0])
        entry_size.insert(0, current_base[1])

        def edit():
            """
            Internal function used by edit_base to edit the text of what appears in bases_window
            :return: None
            """

            base_desc = str.strip(entry_des.get())
            base_size = str.strip(entry_size.get())
            self.bases_box.delete(int(index))
            self.bases_box.insert(int(index), base_desc + " ~ " + base_size)
            if not self.bases_history.__contains__(base_desc + " ~ " + base_size) and not (entry_des.get() =="[Enter base description]" or entry_size.get() == "[Enter base size]"):
                self.bases_history.append(base_desc + " ~ " + base_size)
                self.bases_history.remove("Select from history")
                self.bases_history = sorted(self.bases_history)
                self.bases_history.insert(0, "Select from history")
            self.edit_window.destroy()
            self.focus_index += 1
            self.tables_box.select_set(self.focus_index)
            self.bases_box.select_set(self.focus_index)

        btn_cancel = Tkinter.Button(self.edit_window, text="Cancel", command=self.edit_window.destroy, width=20,
                                    height=3)
        btn_edit = Tkinter.Button(self.edit_window, text="Edit", command=edit, width=20, height=3)

        string_var = Tkinter.StringVar(edit_frame)
        string_var.set(self.bases_history[0])  # default value

        optionbox_bases = Tkinter.OptionMenu(edit_frame, string_var, *self.bases_history)
        self.optionmenu_seletion_index = 0

        def update_from_option(*args):
            if not (string_var.get() == "Select from history"):
                base = string_var.get().split(" ~ ")
                entry_des.delete(0, Tkinter.END)
                entry_size.delete(0, Tkinter.END)
                entry_des.insert(0, base[0])
                entry_size.insert(0, base[1])
            else:

                entry_size.delete(0, Tkinter.END)
                entry_des.delete(0, Tkinter.END)
                entry_des.insert(0, "[Enter base description]")
                entry_size.insert(0, "[Enter base size]")

        string_var.trace('w', update_from_option)
        optionbox_bases.config(width=40)
        optionbox_bases.pack(side=Tkinter.BOTTOM)

        lbl_base.pack(side=Tkinter.TOP)
        lbl_title.pack(side=Tkinter.TOP)
        entry_des.pack(side=Tkinter.TOP)
        entry_size.pack(side=Tkinter.TOP)

        btn_edit.pack(side=Tkinter.LEFT, padx=20)
        btn_cancel.pack(side=Tkinter.RIGHT, padx=20)
        entry_des.selection_range(0, Tkinter.END)

        # Key Bindings
        def enter_pressed(event):
            edit()

        def right_pressed(event):
            entry_size.focus_set()
            entry_size.selection_range(0, Tkinter.END)

        def left_pressed(event):
            entry_des.focus_set()
            entry_des.selection_range(0, Tkinter.END)

        def down_pressed(event):
            optionbox_bases.focus_set()
            if self.optionmenu_seletion_index < len(self.bases_history)-1:
                self.optionmenu_seletion_index += 1
                string_var.set(self.bases_history[self.optionmenu_seletion_index])

        def up_pressed(event):
            if self.optionmenu_seletion_index > 0:
                optionbox_bases.focus_set()
                self.optionmenu_seletion_index -= 1
                string_var.set(self.bases_history[self.optionmenu_seletion_index])
            elif self.optionmenu_seletion_index == 0:
                entry_des.focus_set()
                entry_des.selection_range(0, Tkinter.END)
                self.optionmenu_seletion_index -= 1
                string_var.set(self.bases_history[self.optionmenu_seletion_index])
                entry_size.delete(0, Tkinter.END)
                entry_des.delete(0, Tkinter.END)
                entry_des.insert(0, "[Enter base description]")
                entry_size.insert(0, "[Enter base size]")

            else:
                entry_des.focus_set()
                entry_des.selection_range(0, Tkinter.END)

        def escape_pressed(event):
            self.edit_window.destroy()

        self.edit_window.bind("<Return>", enter_pressed)
        self.edit_window.bind("<KP_Enter>", enter_pressed)
        self.edit_window.bind("<Down>", down_pressed)
        self.edit_window.bind("<Up>", up_pressed)
        self.edit_window.bind("<Left>", left_pressed)
        self.edit_window.bind("<Right>", right_pressed)

        self.edit_window.bind("<Escape>", escape_pressed)

    def finish_bases(self):
        """
        Function gets save directory for finished crosstab report.
        :return: None
        """
        if self.loaded_qfile:
            index = 1
            for item in list(self.bases_box.get(0, Tkinter.END)):
                base = item.split(" ~ ")
                current_table = self.tables.get(index)
                current_table.description = base[0]
                current_table.size = base[1]
                index += 1

            self.__parser.add_bases(self.tables)


            savedirectory = tkFileDialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('excel files', '.xlsx')])
            print savedirectory
            if savedirectory is not "":
                self.__parser.save(savedirectory)
                self.open_file_for_user(savedirectory)
            self.base_window.destroy()


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

    def saveBox(self, title=None, fileName=None, dirName=None, fileExt=".txt", fileTypes=None, asFile=False):

        if fileTypes is None:
            fileTypes = [('all files', '.*'), ('text files', '.txt')]
        # define options for opening
        options = {}
        options['defaultextension'] = fileExt
        options['filetypes'] = fileTypes

        if asFile:
            return tkFileDialog.asksaveasfile(mode='w', **options)
        # will return "" if cancelled
        else:
            return tkFileDialog.asksaveasfilename(**options)

