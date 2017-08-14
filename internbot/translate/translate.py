
class SPSSTranslator(object):
    
    def define_variables(self, survey, path_to_output):
        questions = survey.get_questions()
        file = open(path_to_output, "w+")
        file.write(self.translate_questions(questions))
        
    def translate_questions(self, questions):
        result = ''
        grouped_questions = []
        group_names = []
        for question in questions:
            if question.type == 'MC':
                result += 'VARIABLE LEVEL  %s(ORDINAL).\n' % question.name
            elif question.type == 'Slider':
                if len(question.responses) > 1:
                    grouped_questions.append(self.translate_slider_group(question, group_names))
                else:
                    for response in question.responses:
                        result += 'VARIABLE LEVEL  %s_%s(SCALE).\n' % (question.name, response.code)
            elif question.type == 'CompositeMatrix':
                grouped_questions.append(self.translate_matrix(question, group_names))
            elif question.type == 'CompositeMultipleSelect':
                sub_questions = question.questions
                try:
                    if sub_questions[0].has_carry_forward_responses is False:
                        grouped_questions.append(self.translate_multiselect(question, group_names))
                    elif sub_questions[0].has_carry_forward_responses is True:
                        grouped_questions.append(self.translate_multiselect_cf(question, group_names))
                except:
                    print question.name
            else:
                question.type
        result += self.add_groups(grouped_questions, group_names)
        return result

    def translate_slider_group(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='%s'" % (question.name, question.prompt)
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for response in question.responses:
            result += "%s_%s " % (question.name, response.code)
        result += "VALUE=1\n"
        return result 

    def translate_matrix(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='%s'" % (question.name, question.prompt)
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for sub_question in question.questions:
            result += "%s " % sub_question.name
        result += "VALUE=1\n"
        return result

    def translate_multiselect(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='%s'" % (question.name, question.prompt)
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for sub_question in question.questions:
            result += sub_question.name + ' '
        result += "VALUE=1\n"
        return result

    def translate_multiselect_cf(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='%s'" % (question.name, question.prompt)
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for sub_question in question.questions:
            result += sub_question.name
        result += "VALUE=1\n"
        return result

    def add_groups(self, grouped_questions, grouped_name):
        result = 'EXECUTE.\n\nMRSETS\n'
        for item in grouped_questions:
            result += item
        result += '  /DISPLAY NAME=['
        iteration = 1
        for name in grouped_name:
            result += name
            if iteration < len(grouped_name):
                result += " "
            iteration += 1
        result += '].'
        return result

