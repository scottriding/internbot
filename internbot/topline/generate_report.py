
from question import Question, Questions
from topline_report import ToplineReport
import csv

path_to_csv = '/Users/scotty/Dropbox (Y2 Analytics)/Y2 Analytics Team Folder/Projects/Qualtrics/Comcast/Data/Nominal Frequencies.csv'
path_to_template = '/Users/scotty/Dropbox (Y2 Analytics)/Y2 Analytics Team Folder/R&D/Topline automation/Template/qualtrics_template.docx'
path_to_output = '/Users/scotty/Dropbox (Y2 Analytics)/Y2 Analytics Team Folder/Projects/Qualtrics/Comcast/Deliverables/test.docx'

questions = Questions()

with open(path_to_csv,'rb') as csvfile:
    file = csv.DictReader(csvfile, quotechar = '"') # creates a reader object for file
    for question_data in file:
        questions.add(question_data)

report = ToplineReport(questions, path_to_template)
report.save(path_to_output)
