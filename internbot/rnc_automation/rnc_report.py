from openpyxl import load_workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill, Font
   
class RNCReport(object):

    def __init__ (self, models, path_to_topline=None):
        if path_to_topline != None:
            self.workbook = load_workbook(path_to_topline)
            first_round = False
        else:
            first_round = True

        self.general_style = Font( name='Calibri (Body)', 
                                size=11, 
                                bold=False, 
                                italic=False, 
                                vertAlign=None, 
                                underline='none', 
                                strike=False, 
                                color='ffffff')
        
        self.net_style = Font( name='Calibri (Body)', 
                                size=11, 
                                bold=False, 
                                italic=True, 
                                vertAlign=None, 
                                underline='none', 
                                strike=False, 
                                color='ffffff')

        self.border = Border(left=Side(style='thick'), 
                     right=Side(style='thick'), 
                     top=Side(style='thick'), 
                     bottom=Side(style='thick'))

        self.grey = PatternFill("solid", fgColor="B7B7B7")

        self.darkest_negative = PatternFill("solid", fgColor="B80001")
        self.medium_dark_negative = PatternFill("solid", fgColor="CD4748")
        self.medium_negative = PatternFill("solid", fgColor="DF8A8C")
        self.medium_light_negative = PatternFill("solid", fgColor="EAB9BB")
        self.lightest_negative = PatternFill("solid", fgColor="F1D2D5")

        self.darkest_positive = PatternFill("solid", fgColor="02A747")
        self.medium_dark_positive = PatternFill("solid", fgColor="2AB666")
        self.medium_positive = PatternFill("solid", fgColor="5AC88B")
        self.medium_light_positive = PatternFill("solid", fgColor="B6E7CE")
        self.lightest_positive = PatternFill("solid", fgColor="E6F6F0")

        self.models = models
        if first_round == True:
            self.write_model()
            self.write_diff_previous()
            self.write_diff_first()
            self.write_current_scores()
        else:
            self.insert_cols()
            self.write_current_scores()
            self.update_diff_previous()
            self.update_diff_first()
        
    def write_model(self):
        pass

    def write_diff_previous(self):
        pass

    def write_diff_first(self):
        pass

    def write_current_scores(self):
        pass

    def insert_cols(self):
        pass

    def update_diff_previous(self):
        pass

    def update_diff_first(self):
        pass
        
    def save(self, path_to_output):
        self.workbook.save(path_to_output)
        