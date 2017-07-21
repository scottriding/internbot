from compile import QSFSurveyCompiler
from topline_report import ToplineReport
import csv

path_to_qsf = 'C://Users/kathryn/Downloads/Weber_School_District_Potential_Bond_Viability_Survey.qsf'
path_to_csv = 'C://Users/kathryn/Downloads/Weber School District Weighted Freqs TMF v2.csv'
path_to_template = 'C://Users/kathryn/Dropbox (Y2 Analytics)/Y2 Analytics Team Folder/R&D/Topline automation/Template/topline_template.docx'
path_to_output = 'C://Users/kathryn/Desktop/Weber School District topline.docx'

qsf = QSFSurveyCompiler()
survey = qsf.compile(path_to_qsf)

questions = []
def find_question(question_to_find, survey):
    matching_question = survey.blocks.find_question_by_name(question_to_find)
    if matching_question not in questions:
        questions.append(matching_question)
    return matching_question
    
def find_response(response_to_find, question):
    responses = question.responses
    matching_response = next((response for response in responses if response.response == response_to_find), None)
    return matching_response
    
def add_frequency(response, frequency):
    response.frequency = frequency
    
with open(path_to_csv,'rb') as csvfile:
    file = csv.DictReader(csvfile, quotechar = '"') # creates a reader object for file
    for question_data in file:
        # find question
        matching_question = find_question(question_data['name'], survey)
        # find response
        matching_response = find_response(question_data['response'], matching_question)
        # add frequency to response
        add_frequency(matching_response, question_data['frequency'])
        
report = ToplineReport(questions, path_to_template)
report.save(path_to_output)

