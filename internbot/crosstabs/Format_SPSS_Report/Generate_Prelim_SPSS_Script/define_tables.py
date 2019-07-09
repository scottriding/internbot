import csv

class TableDefiner(object):
    
    def define_tables(self, survey, path_to_output):
        self.count = 1
        output = str(path_to_output) + '/Tables to run.csv'
        questions = survey.get_questions()
        with open(output, 'w') as csvfile:
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
            ['TableIndex'] +
            ['Column specs']
        )   
    
    def write_tables(self, csvwriter, questions):
        for question in questions:
            if question.type == 'DB' or question.type == 'Timing':
                pass
            else:        
                if question.parent == 'CompositeQuestion' and question.type != "CompositeMultipleSelect":
                    self.likert(csvwriter, self.count, question)
                elif question.type == 'TE':
                    if question.subtype == 'ValidNumber':
                        self.text(csvwriter, self.count, question)
                else:
                    variable_name = self.define_name(question)
                    title = self.clean_prompt(question.prompt)
                    base = ''
                    table_index = self.count
                    csvwriter.writerow([
                        '',
                        variable_name,
                        title,
                        base,
                        table_index,
                        ''
                    ])
                    self.count += 1

    def define_name(self, question):
        if question.parent == 'CompositeQuestion':
            return '$%s' % question.name
        else:
            return question.name

    def group(self, question):
        for sub_question in question.questions:
            for response in sub_question.responses:
                if "agree" in response.response or "Agree" in response.response:
                    return False
        return True

    def likert(self, csvwriter, count, question):
        for sub_question in question:
            variable_name = sub_question.name
            title = '%s - %s' % (self.clean_prompt(question.prompt), \
                               self.clean_prompt(sub_question.prompt))
            base = ''
            table_index = self.count
            csvwriter.writerow([
                '',
                variable_name,
                title,
                base,
                table_index,
                ''
            ])
            self.count += 1

    def text(self, csvwriter, count, question):
        variable_name = self.define_name(question)
        title = self.clean_prompt(question.prompt)
        base = ''
        table_index = self.count
        csvwriter.writerow([
            '',
            variable_name,
            title,
            base,
            table_index,
            ''
            ])
        self.count += 1

    def clean_prompt(self, prompt):
        #Necessary python3 weirdness. str.maketrans makes a translation table for translate
        #The first two arguments: the econd replaces the first. Just used spaces, because they have to have the
        #same length and can't be None
        #The third argument is always replaced by None. So here we are removing whitespace and commas from prompt
        result = prompt.translate(str.maketrans(' ',' ','\n'))
        result = result.translate(str.maketrans(' ',' ', ','))
        result = result.translate(str.maketrans(' ',' ', '\t'))
        result = result.translate(str.maketrans(' ',' ', '\r'))
        return result