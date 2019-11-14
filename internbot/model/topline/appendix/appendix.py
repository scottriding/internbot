from model.topline.appendix import document
from model.topline.appendix import spreadsheet
from model.topline.appendix import open_end_question

import csv
from collections import OrderedDict

class Appendix(object):

    def __init__(self):
        self.__questions = OrderedDict()

    def unicode_dict_reader(self, utf8_data, **kwargs):
        csv_reader = csv.DictReader(utf8_data, **kwargs)
        for row in csv_reader:
            if row['variable'] != "":
                yield {key:value for key, value in row.items()}

    def build_appendix_model(self, path_to_appendix):
        text_responses = self.unicode_dict_reader(open(path_to_appendix))
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
        self.__questions = OrderedDict() ## empty out questions
            
        print("Finished!")

