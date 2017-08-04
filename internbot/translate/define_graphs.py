class GraphDefiner(object):

    def define_graphs(self, survey, path_to_output):
        questions = survey.get_questions()
        file = open(path_to_output, "w+")
        file.write(self.define_questions(questions))
        
    def define_questions(self, questions):
        result = ''
        grouped_questions = []
        group_names = []
        for question in questions:
            if question.type == 'MC' and question.subtype in ['SAVR','SAHR','DL']:
                result += self.multiple_choice(question)
            elif question.type == 'Slider':
                result += self.numeric_scale(question)
            elif question.type == 'Composite' and question.subtype == 'SingleAnswer':
                result += self.composite_single(question)
        return result  

    def multiple_choice(self, question):
        mc_graph = ''
        if len(question.response_order) == 2:
            mc_graph += 'Pie chart \n'
        return mc_graph

    def numeric_scale(self, question):
        return ''

    def composite_single(self, question):
        return ''