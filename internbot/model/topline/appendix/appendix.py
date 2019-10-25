from model.topline.appendix import document
from model.topline.appendix import spreadsheet
from model.topline.appendix import open_end_question

import csv
from collections import OrderedDict

class Appendix(object):

    def __init__(self):
        self.__questions = OrderedDict()

    def build_appendix_model(self, path_to_appendix):
        with open(path_to_appendix, encoding='utf-8-sig') as csvfile:
            text_responses = csv.DictReader(csvfile)
            for response in text_responses:
                if self.__questions.get(response['variable']) is None:
                    new_question = open_end_question.OpenEndQuestion(response['variable'], response['prompt'])
                    new_question.add_response(response['label'])
                    self.__questions[response['variable']] = new_question
                else:
                    current_question = self.__questions.get(response['variable'])
                    current_question.add_response(response['label'])

    def build_appendix_report(self, path_to_output, is_document, image_path, template_path):
        if is_document:
            builder = document.Document(template_path)
        else:
            builder = spreadsheet.Spreadsheet(image_path)

        builder.write_appendix(self.__questions)
        builder.save(path_to_output)
            
        print("Finished!")

