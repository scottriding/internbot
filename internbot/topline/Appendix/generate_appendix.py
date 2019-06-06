import csv
from collections import OrderedDict
from topline.Appendix.format_report import SSAppendixBuilder, DocAppendixBuilder, OpenEndQuestion

class AppendixGenerator(object):

    def __init__(self):
        self.__questions = OrderedDict()
        self.is_qualtrics = False

    def unicode_dict_reader(self, utf8_data, **kwargs):
        csv_reader = csv.DictReader(utf8_data, **kwargs)
        for row in csv_reader:
            if row['variable'] != "":
                yield {key:value for key, value in row.items()}

    def parse_file(self, path_to_appendix, is_qualtrics):
        self.is_qualtrics = is_qualtrics
        text_responses = self.unicode_dict_reader(open(path_to_appendix))
        for response in text_responses:
            if self.__questions.get(response['variable']) is None:
                new_question = OpenEndQuestion(response['variable'], response['prompt'])
                new_question.add_response(response['label'])
                self.__questions[response['variable']] = new_question
            else:
                current_question = self.__questions.get(response['variable'])
                current_question.add_response(response['label'])

    def write_appendix(self, path_to_output, path_to_template = '', is_spreadsheet=False):
        if is_spreadsheet is True:
            builder = SSAppendixBuilder(self.is_qualtrics)
            builder.write_appendix(self.__questions)
            builder.save(path_to_output)
        else:
            builder = DocAppendixBuilder(path_to_template)
            builder.write_appendix(self.__questions)
            builder.save(path_to_output)
        print("Finished!")

