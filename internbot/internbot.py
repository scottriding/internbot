"""internbot main

internbot is responsible for automation for Y2 Analytics that can
be handled via python featuring:
- Topline Report documentation in Microsoft Word
- Cross tab variable and table scripting for SPSS and Microsoft Excel cell highlighting
- MaxDiff web service (to be implemented)

@Authors: Scotty Riding and Kathryn Riding
@Version: May 16, 2018
"""
import base
import crosstabs
import topline
import Tkinter
import tkMessageBox
import tkFileDialog
import os

"""variable_script button event

variable_script methodology is responsible for creating SPSS rename variable script for refactoring a survey data file
and a Tables to Run.csv document for a table of contents tab in the final cross tabs report.

variable_script uses both the path to a Qualtrics (.qsf) file and output directory folder inputs from a user and sends those details
to parsing and compiling code.

@throws: error message if there is trouble parsing or compiling files. - TO BE IMPLEMENTED
"""
def variable_script():
    fpath = os.path.join(os.path.expanduser("~"), "Desktop")
    ask_qsf = tkMessageBox.askokcancel("Select Qualtrics File", "Please select the Qualtrics survey .qsf file.")
    if ask_qsf is True: # user selected ok
        qsffilename = tkFileDialog.askopenfilename(initialdir = fpath,title = "Select Qualtrics survey file",filetypes = (("Qualtrics file","*.qsf"),("all files","*.*")))
        compiler = base.QSFSurveyCompiler()
        survey = compiler.compile(qsffilename)
        ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished variable script.")
        if ask_output is True: # user selected ok
            savedirectory = tkFileDialog.askdirectory()
            variables = crosstabs.Generate_Prelim_SPSS_Script.SPSSTranslator()
            tables = crosstabs.Generate_Prelim_SPSS_Script.TableDefiner()    
            variables.define_variables(survey, savedirectory)
            tables.define_tables(survey, savedirectory)

"""table_script button event

table_script methodology is responsible for creating SPSS table script file to create
cross tabs tables.

table_script uses both the path to a comma-seperated (.csv) file and output directory folder inputs from a user and sends
those details to parsing and compiling code.

@throws: error message if there is trouble parsing or compiling files. - TO BE IMPLEMENTED
"""
def table_script():
    fpath = os.path.join(os.path.expanduser("~"), "Desktop")
    script = crosstabs.Generate_Table_Script.TableScript()
    ask_tables = tkMessageBox.askokcancel("Select Tables to Run.csv File", "Please select the tables to run .csv file.")
    if ask_tables is True:
        tablesfilename = tkFileDialog.askopenfilename(initialdir = fpath,title = "Select tables file",filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
        ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished table script.")
        if ask_output is True:
            savedirectory = tkFileDialog.askdirectory()
            script.compile_scripts(tablesfilename, savedirectory)

"""highlight button event

highlight methodology is responsible for renaming tabs and highlighting significant cells in a cross tabs report Microsoft Excel file.

highlight uses both the path to a Microsoft Excel (.xlsx) file and output directory folder inputs from a user and sends
those details to parsing and highlighting code.

@throws: error message if there is trouble parsing or compiling files. - TO BE IMPLEMENTED
"""
def highlight():
    fpath = os.path.join(os.path.expanduser("~"), "Desktop")
    ask_xlsx = tkMessageBox.askokcancel("Select Tables Microsoft Excel File", "Please select the combined table .xlsx file.")
    if ask_xlsx is True: # user selected ok
        renamer = crosstabs.Polish_Final_Report.RenameTabs()
        ask_tables = tkMessageBox.askokcancel("Select Tables .csv File", "Please select the Tables to run.csv file.")
        if ask_tables is True: # user selected ok
            tablesfilename = tkFileDialog.askopenfilename(initialdir = fpath,title = "Select Comma Seperated file",filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
            ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished table script.")
            if ask_output is True: # user selected ok
                savedirectory = tkFileDialog.askdirectory()
                renamer.rename(tabsfilename, tablesfilename, savedirectory)
                highlighter = crosstabs.Polish_Final_Report.Highlighter(savedirectory)
                highlighter.highlight(savedirectory)
                tkMessageBox.showinfo("Finished", "The highlighted report is saved as \"Highlighted.xlsx\" in your chosen directory.")

"""open_topline button event

open_topline methodology is responsible for creating a basic topline report.

open_topline parses a file path to build an appropriate report object from either a Qualtrics survey fie (.qsf) or comma-seperated file (.csv).

@throws: error message if there is trouble parsing or compiling files. - TO BE IMPLEMENTED
"""
def open_topline():
    fpath = os.path.join(os.path.expanduser("~"), "Desktop")
    filename = tkFileDialog.askopenfilename(initialdir = fpath,title = "Select survey file",filetypes = (("Qualtrics files","*.qsf"),("comma seperated files","*.csv"),("all files","*.*")))
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

"""build_report button event

build_report methodology is responsible for creating the final topline report with an output directory folder input from a user.

@throws: error message if there is trouble parsing or compiling files. - TO BE IMPLEMENTED
"""
def build_report(isQSF, report):
    fpath = os.path.join(os.path.expanduser("~"), "Desktop")
    template_file = open("topline_template.docx", "r")
    ask_output = tkMessageBox.askokcancel("Output directory", "Please select the directory for finished report.")
    if ask_output is True:
        savedirectory = tkFileDialog.askdirectory()
        if isQSF is True:
            ask_freq = tkMessageBox.askokcancel("Frequency file", "Please select the topline .csv frequency file.")
            if ask_freq is True:
                freqfilename = tkFileDialog.askopenfilename(initialdir = fpath,title = "Select frequency file",filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
                report.generate_basic_topline(freqfilename, template_file, savedirectory)
        else:
            report.generate_basic_topline(template_file, savedirectory)

"""redirect_xtabs button event

Creates a new window with all possible paths for creating a cross tabs report.
"""
def redirect_xtabs():
    redirect_window = Tkinter.Toplevel()
    redirect_window.title("Y2 Crosstab Automation")
    message = "Please select the files to produce."
    Tkinter.Label(redirect_window, text = message).pack()
    btn_var = Tkinter.Button(redirect_window, text = "Variable script", command = variable_script)
    btn_tab = Tkinter.Button(redirect_window, text = "Table script", command = table_script)
    btn_high = Tkinter.Button(redirect_window, text = "Highlight", command = highlight)
    btn_cancel = Tkinter.Button(redirect_window, text = "Cancel", command = redirect_window.destroy)
    btn_var.pack(padx = 5, side = Tkinter.LEFT)
    btn_tab.pack(padx = 5, side = Tkinter.LEFT)
    btn_high.pack(padx = 5, side = Tkinter.LEFT)
    btn_cancel.pack(padx = 5, side = Tkinter.LEFT)

"""redirect_topline button event

Creates a new window with buttons to redirect user to topline-centric automation.
"""
def redirect_topline():
    redirect_window = Tkinter.Toplevel()
    redirect_window.title("Y2 Topline Report Automation")
    message = "Please open survey file to use."
    Tkinter.Label(redirect_window, text = message).pack()
    btn_open = Tkinter.Button(redirect_window, text = "Open", command = open_topline)
    btn_cancel = Tkinter.Button(redirect_window, text = "Cancel", command = redirect_window.destroy)
    btn_open.pack(padx = 10, side = Tkinter.LEFT)
    btn_cancel.pack(padx = 10, side = Tkinter.LEFT)

### main window ###
window = Tkinter.Tk()
window.title("Internbot: 01011001 00000010") # Internbot: Y2

### main window buttons ###
btn_xtabs = Tkinter.Button(window, text="Run crosstabs", command=redirect_xtabs)
btn_report = Tkinter.Button(window, text="Run topline report", command=redirect_topline)
btn_quit = Tkinter.Button(window, text="Quit", command=window.destroy)
btn_xtabs.pack(padx=10,side=Tkinter.LEFT)
btn_report.pack(padx=10,side=Tkinter.LEFT)
btn_quit.pack(padx=10,side=Tkinter.LEFT)

### end - loop it all ###
window.mainloop()


