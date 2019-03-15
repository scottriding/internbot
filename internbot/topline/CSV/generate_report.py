from question import Question, Questions
from topline_report import ToplineReport
import csv

class ReportGenerator(object):
    def __init__(self, path_to_csv, round_no = 1):
        self.__questions = Questions()
        question_data = self.unicode_dict_reader(open(path_to_csv))
        self.create_questions(question_data, round_no)

    def unicode_dict_reader(self, utf8_data, **kwargs):
        csv_reader = csv.DictReader(utf8_data, **kwargs)
        for row in csv_reader:
        	if row["variable"] is not "":
        		yield {unicode(key, 'utf-8'):unicode(value, 'utf-8') for key, value in row.iteritems()}

    def create_questions(self, question_data, round_no):
        for question in question_data:
            self.__questions.add(question, round_no)

    def generate_topline(self, path_to_template, path_to_output, years):
        report = ToplineReport(self.__questions, path_to_template, years)
        report.save(str(path_to_output) + '/topline_report.docx')