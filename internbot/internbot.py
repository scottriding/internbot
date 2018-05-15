import argparse
import base
import sys
import crosstabs
import topline
import Tkinter
import tkMessageBox
import tkFileDialog

### crosstabs menu button events
def variable_script():
    ask_qsf = tkMessageBox.askquestion("Select QSF File", "Please select the survey .qsf file.")
    if ask_qsf == 'yes':
        qsffilename = tkFileDialog.askopenfilename(initialdir = "/Users/",title = "Select Qualtrics survey file",filetypes = (("Qualtrics file","*.qsf"),("all files","*.*")))
        compiler = base.QSFSurveyCompiler()
        survey = compiler.compile(qsffilename)
        ask_output = tkMessageBox.askquestion("Output directory", "Please select the directory for finished variable script.")
        if ask_output == 'yes':
            savedirectory = tkFileDialog.askdirectory()
            variables = crosstabs.Generate_Prelim_SPSS_Script.SPSSTranslator()
            tables = crosstabs.Generate_Prelim_SPSS_Script.TableDefiner()    
            variables.define_variables(survey, savedirectory)
            tables.define_tables(survey, savedirectory)

def table_script():
    script = crosstabs.Generate_Table_Script.TableScript()
    ask_tables = tkMessageBox.askquestion("Select Tables to Run.csv File", "Please select the tables to run .csv file.")
    if ask_tables == 'yes':
        tablesfilename = tkFileDialog.askopenfilename(initialdir = "/Users/",title = "Select tables file",filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
        ask_output = tkMessageBox.askquestion("Output directory", "Please select the directory for finished table script.")
        if ask_output == 'yes':
            savedirectory = tkFileDialog.askdirectory()
            script.compile_scripts(tablesfilename, savedirectory)

def highlight():
    ask_xlsx = tkMessageBox.askquestion("Select Tables Microsoft Excel File", "Please select the combined table .xlsx file.")
    if ask_xlsx == 'yes':
        renamer = crosstabs.Polish_Final_Report.RenameTabs()
        ask_tables = tkMessageBox.askquestion("Select Tables .csv File", "Please select the Tables to run.csv file.")
        if ask_tables == 'yes':
            tablesfilename = tkFileDialog.askopenfilename(initialdir = "/Users/",title = "Select Comma Seperated file",filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
            ask_output = tkMessageBox.askquestion("Output directory", "Please select the directory for finished table script.")
            if ask_output == 'yes':
                savedirectory = tkFileDialog.askdirectory()
                renamer.rename(tabsfilename, tablesfilename, savedirectory)
                highlighter = crosstabs.Polish_Final_Report.Highlighter(savedirectory)
                highlighter.highlight(savedirectory)
                tkMessageBox.showinfo("Finished", "The highlighted report is saved as \"Highlighted.xlsx\" in your chosen directory.")

### topline menu button events
def open_topline():
    filename = tkFileDialog.askopenfilename(initialdir = "/Users/",title = "Select survey file",filetypes = (("Qualtrics files","*.qsf"),("comma seperated files","*.csv"),("all files","*.*")))
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
    else:
        pass

def build_report(isQSF, report):
    ask_template = tkMessageBox.askquestion("Template document", "Please select the topline .docx template file.")
    if ask_template == 'yes':
        tempfilename = tkFileDialog.askopenfilename(initialdir = "/Users/",title = "Select frequency file",filetypes = (("Microsoft docx files","*.docx"),("all files","*.*")))
        ask_output = tkMessageBox.askquestion("Output directory", "Please select the directory for finished report.")
        if ask_output == 'yes':
            savedirectory = tkFileDialog.askdirectory()
            if isQSF is True:
                ask_freq = tkMessageBox.askquestion("Frequency file", "Please select the topline .csv frequency file.")
                if ask_freq == 'yes':
                    freqfilename = tkFileDialog.askopenfilename(initialdir = "/Users/",title = "Select frequency file",filetypes = (("comma seperated files","*.csv"),("all files","*.*")))
                    report.generate_basic_topline(freqfilename, tempfilename, savedirectory)
            else:
                report.generate_basic_topline(tempfilename, savedirectory)

### main button events ###
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

def redirect_topline():
    redirect_window = Tkinter.Toplevel()
    redirect_window.title("Y2 Topline Report Automation")
    message = "Please open survey file to use."
    Tkinter.Label(redirect_window, text = message).pack()
    btn_open = Tkinter.Button(redirect_window, text = "Open", command = open_topline)
    btn_cancel = Tkinter.Button(redirect_window, text = "Cancel", command = redirect_window.destroy)
    btn_open.pack(padx = 10, side = Tkinter.LEFT)
    btn_cancel.pack(padx = 10, side = Tkinter.LEFT)

def clicked_xtabs():
    window.filename = tkFileDialog.askopenfilename(initialdir = "/Users/",title = "Select file",filetypes = (("qualtrics files","*.qsf"),("all files","*.*")))

### main window ###
window = Tkinter.Tk()
window.title("Internbot: 01011001 00000010")
window.geometry("380x100")

### main window buttons ###
btn_xtabs = Tkinter.Button(window, text="Run crosstabs", command=redirect_xtabs)
btn_report = Tkinter.Button(window, text="Run topline report", command=redirect_topline)
btn_quit = Tkinter.Button(window, text="Quit", command=window.destroy)
btn_xtabs.pack(padx=10,side=Tkinter.LEFT)
btn_report.pack(padx=10,side=Tkinter.LEFT)
btn_quit.pack(padx=10,side=Tkinter.LEFT)

### end - loop it all ###
window.mainloop()
