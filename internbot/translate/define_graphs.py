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
            if question.type == 'MC':
                result += self.multiple_choice(question)
            elif question.type == 'Slider':
                result += self.numeric_scale(question)
            elif question.type == 'Composite':
                result += self.composite_single(question)
            elif question.type == 'TE':
                result += self.open_ended(question)
        return result  

    def multiple_choice(self, question):
        mc_graph = ''
        if len(question.response_order) == 2:
            mc_graph += '%s:\t Pie chart \n' % question.name
        elif len(question.response_order) == 5:
            mc_graph += self.define_bar_chart(question)
        return mc_graph

    def define_bar_chart(self, question):
        return ''

    def numeric_scale(self, question):
        scale_graph = '%s:\t Histogram \n' % question.name
        return scale_graph

    def composite_single(self, question):
        return ''

    def open_ended(self, question):
        return ''