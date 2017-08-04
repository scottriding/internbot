class GraphDefiner(object):

    def define_graphs(self, survey, path_to_output):
        questions = survey.get_questions()
        file = open(path_to_output, "w+")
        file.write(self.define_questions(questions))
        
    def define_questions(self, questions):
        result = ''
        cannot_determine = []
        for question in questions:
            if question.type == 'MC':
                result += self.multiple_choice(question)
            elif question.type == 'Slider':
                result += self.numeric_scale(question)
            elif question.type == 'Composite':
                result += self.composite_single(question)
            elif question.type == 'TE':
                result += self.define_bar_chart(question)
            elif question.type == 'CS':
                result += self.define_bar_chart(question)
            else:
                cannot_determine.append('%s: \t Cannot define - %s' % (question.name, question.type))
        for reason in cannot_determine:
            result += reason + '\n'
        return result  

    def multiple_choice(self, question):
        mc_graph = ''
        if len(question.response_order) == 2:
            mc_graph += '%s: \t Pie chart \n' % question.name
        elif len(question.response_order) == 5:
            mc_graph += self.define_bar_chart(question)
        return mc_graph

    def define_bar_chart(self, question):
        return '%s: \t bar chart \n' % question.name

    def numeric_scale(self, question):
        scale_graph = '%s:\t Histogram \n' % question.name
        return scale_graph

    def composite_single(self, composite_question):
        composite_graph = ''
        for sub_question in composite_question:
            composite_graph += self.define_bar_chart(sub_question)
        return composite_graph
