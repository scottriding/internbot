import csv

class TableDefiner(object):
    
    def define_tables(self, survey, path_to_output):
        questions = survey.get_questions()
        with open(path_to_output, 'wb') as csvfile:
            csvwriter = csv.writer(csvfile, delimiter=',',
                                   quotechar="'")
            self.write_header(csvwriter)
            self.write_tables(csvwriter, questions)
    
    def write_header(self, csvwriter):
        csvwriter.writerow(
            ['Empty'] + 
            ['VariableName'] + 
            ['Title'] + 
            ['Base'] + 
            ['TableIndex']
        )   
    
    def write_tables(self, csvwriter, questions):
        count = 1
        for question in questions:
            if question.type == 'TE' or question.type == 'DB' or question.type == 'Timing':
                pass
            else:
                variable_name = self.define_name(question)
                title = self.clean_prompt(question.prompt)
                base = ''
                table_index = count
                csvwriter.writerow([
                    '',
                    variable_name,
                    title,
                    base,
                    table_index
                ])
                count += 1

    def define_name(self, question):
        if question.parent == 'CompositeQuestion':
            return '$%s' % question.name
        else:
            return question.name

    def clean_prompt(self, prompt):
        result = prompt.translate(None,",'\n\t\r")
        return result