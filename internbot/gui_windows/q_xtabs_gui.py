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
        self.base_window = tkinter.Toplevel(self.__window)
        self.base_window.withdraw()
        width = 1100
        height = 420
        self.base_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        self.base_window.title("Q Research Crosstabs Base Entry")
        self.loaded_qfile = False

        self.boxes_frame = tkinter.Frame(self.base_window)
        self.boxes_frame.pack(padx=10, pady=10, side=tkinter.LEFT, fill=tkinter.BOTH)
        tables_horiz_scrollbar = tkinter.Scrollbar(self.boxes_frame)
        tables_horiz_scrollbar.config(orient=tkinter.HORIZONTAL)
        tables_horiz_scrollbar.pack(side=tkinter.BOTTOM, fill=tkinter.X)
        tables_vert_scrollbar = tkinter.Scrollbar(self.boxes_frame)
        tables_vert_scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)

        self.tables_box = tkinter.Listbox(self.boxes_frame, selectmode=tkinter.SINGLE, width=55, height=10,
                                          yscrollcommand=tables_vert_scrollbar.set,
                                          xscrollcommand=tables_horiz_scrollbar.set, exportselection=False)
        self.tables_box.pack(expand=True, side=tkinter.LEFT, fill=tkinter.BOTH)

        self.bases_frame = tkinter.Frame(self.base_window)
        self.bases_frame.pack(padx=10, pady=10, side=tkinter.RIGHT, fill=tkinter.BOTH)
        bases_horiz_scrollbar = tkinter.Scrollbar(self.bases_frame)
        bases_horiz_scrollbar.config(orient=tkinter.HORIZONTAL)
        bases_horiz_scrollbar.pack(side=tkinter.BOTTOM, fill=tkinter.X)
        bases_vert_scrollbar = tkinter.Scrollbar(self.bases_frame)
        bases_vert_scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)

        self.bases_box = tkinter.Listbox(self.bases_frame, selectmode=tkinter.SINGLE, width=54, height=10,
                                         yscrollcommand=bases_vert_scrollbar.set,
                                         xscrollcommand=bases_horiz_scrollbar.set, exportselection=False)
        self.bases_box.pack(expand=False, side=tkinter.RIGHT, fill=tkinter.BOTH)

        self.focus_index = 0

        self.tables_box.select_set(self.focus_index)
        self.bases_box.select_set(self.focus_index)

        tables_vert_scrollbar.config(command=self.tables_box.yview)
        tables_horiz_scrollbar.config(command=self.tables_box.xview)
        bases_vert_scrollbar.config(command=self.bases_box.yview)
        bases_horiz_scrollbar.config(command=self.bases_box.xview)

        buttons_frame = tkinter.Frame(self.base_window)
        buttons_frame.pack(pady=10, side=tkinter.RIGHT, fill=tkinter.BOTH)
        self.btn_edit = tkinter.Button(buttons_frame, text="Edit", command=self.parse_bases, width=20, height=3)
        self.btn_done = tkinter.Button(buttons_frame, text="Done", command=self.finish_bases, width=20, height=3)
        self.btn_open = tkinter.Button(buttons_frame, text="Open", command=self.qtab_open, width=20, height=3)
        self.btn_csv = tkinter.Button(buttons_frame, text="CSV Options", command=self.csv_options_menu, width=20, height=3)
        self.btn_cancel = tkinter.Button(buttons_frame, text="Cancel", command=self.base_window.destroy, width=20,
                                         height=3)
        self.btn_bot = tkinter.Button(buttons_frame, image=self.bot_render, borderwidth=0,
                                      highlightthickness=0, relief=tkinter.FLAT, bg="white", height=65, width=158,
                                      command=self.bases_help_window)
        self.btn_cancel.pack(pady=2, side=tkinter.BOTTOM)
        self.btn_done.pack(pady=2, side=tkinter.BOTTOM)
        self.btn_edit.pack(pady=2, side=tkinter.BOTTOM)
        self.btn_csv.pack(pady=2, side=tkinter.BOTTOM)
        self.btn_open.pack(pady=2, side=tkinter.BOTTOM)
        self.btn_bot.pack(pady=2, side=tkinter.BOTTOM)

        self.btn_done.config(state=tkinter.DISABLED)
        self.btn_csv.config(state=tkinter.DISABLED)
        self.btn_edit.config(state=tkinter.DISABLED)

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

    def qtab_open(self):
        """
        Funtion opens a Q Research .xlsx file and loads the table information into base_window for display to user
        :return: None
        """
        print("Opening Q Research file...")
        if not self.loaded_qfile:
            report_file_name = filedialog.askopenfilename(initialdir=self.fpath, title="Select report file",
                                                            filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
            if report_file_name is not "":
                self.__parser = crosstabs.Format_Q_Report.QParser(report_file_name)
                self.tables = self.__parser.format_report()
                prompts = []
                for key, value in self.tables.items():
                    prompts.append(value.prompt)

                for prompt in prompts:
                    self.tables_box.insert(tkinter.END, prompt)
                base_sizes = ["[Enter base description] ~ [Enter base size]"] * len(prompts)
                for size in base_sizes:
                    self.bases_box.insert(tkinter.END, size)

                self.focus_index = 0

                self.tables_box.select_set(self.focus_index)
                self.bases_box.select_set(self.focus_index)

                self.loaded_qfile = True
                self.base_window.focus_force()
                self.btn_done.config(state=tkinter.NORMAL)
                self.btn_csv.config(state=tkinter.NORMAL)
                self.btn_edit.config(state=tkinter.NORMAL)
        else:
            ask_lost_work = messagebox.askyesno("Select Q Research report file",
                                                  "You will lose your work. \nDo you want to continue?")
            if ask_lost_work:
                self.loaded_qfile = False
                self.qtab_open()

    def bases_help_window(self):
        """
        Function displays help information about how to use bases window.
        A part of the Q Research crosstabs process.
        :return: None
        """
        help_window = tkinter.Toplevel(self.__window)
        help_window.lift()
        help_window.withdraw()
        width = 300
        height = 400
        help_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        message = "\nBases Help"
        tkinter.Label(help_window, text=message, font=self.header_font, fg=self.header_color).pack()
        info_message = "Select Open to import a Q output file." \
                       "\nThis should be a .xlsx file." \
                       "\n\nYou can change the current selection with" \
                       "\nthe up/down arrow keys or your mouse." \
                       "\n\nEdit the base and n info of the current" \
                       "\ntable by selecting Edit or hitting enter." \
                       "\n\nWhen you're done, select Done and you" \
                       "\nwill be prompted to select the directory of" \
                       "\nyour finished report.\n"
        tkinter.Label(help_window, text=info_message, justify=tkinter.LEFT).pack()
        btn_ok = tkinter.Button(help_window, text="Ok", command=help_window.destroy, height=3, width=20)
        btn_ok.pack(padx=5, side=tkinter.BOTTOM, expand=False)
        help_window.deiconify()

        def enter_pressed(event):
            help_window.destroy()

        help_window.bind("<Return>", enter_pressed)
        help_window.bind("<KP_Enter>", enter_pressed)

    def parse_bases(self):
        """
        Function calls edit_base for every entry in current selection of base_window
        :return: None
        """
        if self.loaded_qfile:
            if len(self.bases_box.curselection()) is 0:
                messagebox.askokcancel("Select from Right",
                                         "Please make a selection from the list on the right.")
            else:
                for index in self.tables_box.curselection():
                    self.tables_box.select_clear(index)
                for index in self.bases_box.curselection():
                    self.focus_index = index
                    self.edit_base_window(index)

        else:
            messagebox.askokcancel("Select Q Research report file",
                                     "Please open the Q Research report file first.")

    def unicode_dict_reader(self, utf8_data, **kwargs):
        csv_reader = csv.DictReader(utf8_data, **kwargs)
        for row in csv_reader:
            if row['Base Description'] != "":
                yield {key: value for key, value in row.items()}

    def csv_options_menu(self):
        """
        Funtion asks the user if they would like to create a .docx or a .xlsx
        :return: None
        """
        self.choose_window = tkinter.Toplevel(self.__window)
        self.choose_window.withdraw()
        width = 250
        height = 250
        self.choose_window.geometry("%dx%d+%d+%d" % (
            width, height, self.mov_x + self.window_width / 2 - width / 2,
            self.mov_y + self.window_height / 2 - height / 2))
        message = "Export/Load CSV file with\nyour bases and n's"
        tkinter.Label(self.choose_window, text=message, font=self.header_font, fg=self.header_color).pack(expand=False)
        btn_export = tkinter.Button(self.choose_window, text="Export CSV", command=self.export_csv, height=3, width=20)

        btn_load = tkinter.Button(self.choose_window, text="Load CSV", command=self.load_csv, height=3, width=20)
        btn_cancel = tkinter.Button(self.choose_window, text="Cancel", command=self.choose_window.destroy, height=3, width=20)

        btn_cancel.pack(ipadx=5, side=tkinter.BOTTOM, expand=False)
        btn_load.pack(ipadx=5, side=tkinter.BOTTOM, expand=False)
        btn_export.pack(ipadx=5, side=tkinter.BOTTOM, expand=False)
        self.choose_window.deiconify()

    def load_csv(self):
        base_file_name = filedialog.askopenfilename(initialdir=self.fpath, title="Select base file",
                                                      filetypes=(
                                                      ("comma separated files", "*.csv"), ("all files", "*.*")))
        if base_file_name is not "":
            base_details = self.unicode_dict_reader(open(base_file_name))
            index = 0
            for base_detail in base_details:
                if index < len(self.tables):
                    base = base_detail["Base Description"]
                    n = base_detail["Base Size (n)"]
                    self.bases_box.delete(int(index))
                    self.bases_box.insert(int(index), base + " ~ " + n)
                    index += 1
        self.choose_window.destroy()

    def export_csv(self):
        savedirectory = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('comma separated values', '.csv')])
        if savedirectory is not "":
            with open(savedirectory, 'w') as csvfile:
                csvwriter = csv.writer(csvfile, delimiter=',', quotechar="'")
                csvwriter.writerow(
                    ['Table Name'] +
                    ['Prompt'] +
                    ['Base Description'] +
                    ['Base Size (n)']
                )
                for key, value in self.tables.items():
                    csvwriter.writerow([
                        str("Table "+str(key)),
                        value.prompt.translate(str.maketrans(' ',' ', ',')),
                    ])
            self.open_file_for_user(savedirectory)

    def edit_base_window(self, index):
        """
        Function defines a window for the user to input Base and N for a table.
        :param base:
        :param index:
        :return:
        """
        self.edit_window = tkinter.Toplevel(self.base_window)
        width = 400
        height = 150
        self.edit_window.geometry("%dx%d+%d+%d" % (
        width, height, self.mov_x + self.window_width / 2 - width / 2, self.mov_y + self.window_height / 2 - height / 2))
        self.edit_window.title("Add base and n")
        edit_frame = tkinter.Frame(self.edit_window)
        edit_frame.pack(side=tkinter.TOP)
        entry_frame = tkinter.Frame(edit_frame)
        entry_frame.pack(side=tkinter.TOP)
        base_frame = tkinter.Frame(entry_frame)
        base_frame.pack(side=tkinter.LEFT, padx=5, pady=5)
        lbl_base = tkinter.Label(base_frame, text="Base description:")
        entry_des = tkinter.Entry(base_frame, width=30)
        spinbox_frame = tkinter.Frame(edit_frame)
        spinbox_frame.pack(side=tkinter.BOTTOM)
        entry_des.focus_set()
        size_frame = tkinter.Frame(entry_frame)
        size_frame.pack(side=tkinter.RIGHT, padx=5, pady=5)
        lbl_title = tkinter.Label(size_frame, text="Size (n):")
        entry_size = tkinter.Entry(size_frame, width=20)

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

        btn_cancel = tkinter.Button(self.edit_window, text="Cancel", command=self.edit_window.destroy, width=20,
                                    height=3)
        btn_edit = tkinter.Button(self.edit_window, text="Edit", command=edit, width=20, height=3)

        string_var = tkinter.StringVar(edit_frame)
        string_var.set(self.bases_history[0])  # default value

        optionbox_bases = tkinter.OptionMenu(edit_frame, string_var, *self.bases_history)
        self.optionmenu_seletion_index = 0

        def update_from_option(*args):
            if not (string_var.get() == "Select from history"):
                base = string_var.get().split(" ~ ")
                entry_des.delete(0, tkinter.END)
                entry_size.delete(0, tkinter.END)
                entry_des.insert(0, base[0])
                entry_size.insert(0, base[1])
            else:

                entry_size.delete(0, tkinter.END)
                entry_des.delete(0, tkinter.END)
                entry_des.insert(0, "[Enter base description]")
                entry_size.insert(0, "[Enter base size]")

        string_var.trace('w', update_from_option)
        optionbox_bases.config(width=40)
        optionbox_bases.pack(side=tkinter.BOTTOM)

        lbl_base.pack(side=tkinter.TOP)
        lbl_title.pack(side=tkinter.TOP)
        entry_des.pack(side=tkinter.TOP)
        entry_size.pack(side=tkinter.TOP)

        btn_edit.pack(side=tkinter.LEFT, padx=20)
        btn_cancel.pack(side=tkinter.RIGHT, padx=20)
        entry_des.selection_range(0, tkinter.END)

        # Key Bindings
        def enter_pressed(event):
            edit()

        def right_pressed(event):
            entry_size.focus_set()
            entry_size.selection_range(0, tkinter.END)

        def left_pressed(event):
            entry_des.focus_set()
            entry_des.selection_range(0, tkinter.END)

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
                entry_des.selection_range(0, tkinter.END)
                self.optionmenu_seletion_index -= 1
                string_var.set(self.bases_history[self.optionmenu_seletion_index])
                entry_size.delete(0, tkinter.END)
                entry_des.delete(0, tkinter.END)
                entry_des.insert(0, "[Enter base description]")
                entry_size.insert(0, "[Enter base size]")

            else:
                entry_des.focus_set()
                entry_des.selection_range(0, tkinter.END)

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
            for item in list(self.bases_box.get(0, tkinter.END)):
                base = item.split(" ~ ")
                current_table = self.tables.get(index)
                current_table.description = base[0]
                current_table.size = base[1]
                index += 1

            self.__parser.add_bases(self.tables)


            savedirectory = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('excel files', '.xlsx')])

            if savedirectory is not "":
                thread = threading.Thread(target=self.finish_bases_worker, args=(True, savedirectory))
                thread.start()
            self.base_window.destroy()

    def finish_bases_worker(self, necessary, savedirectory):
        print("Running report...")
        self.__parser.save(savedirectory)
        print("Done!")
        self.open_file_for_user(savedirectory)

    def open_sound(self):

        def play_sound():
            audio_file = os.path.expanduser("~/Documents/GitHub/internbot/internbot/templates_images/open.mp3")
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

    def saveBox(self, title=None, fileName=None, dirName=None, fileExt=".txt", fileTypes=None, asFile=False):

        if fileTypes is None:
            fileTypes = [('all files', '.*'), ('text files', '.txt')]
        # define options for opening
        options = {}
        options['defaultextension'] = fileExt
        options['filetypes'] = fileTypes

        if asFile:
            return filedialog.asksaveasfile(mode='w', **options)
        # will return "" if cancelled
        else:
            return filedialog.asksaveasfilename(**options)

