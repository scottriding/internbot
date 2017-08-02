
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
            if question.type == 'MC' and question.subtype in ['SAVR','SAHR','DL']:
                result += 'VARIABLE LEVEL  %s(ORDINAL).\n' % question.name
            elif question.type == 'Slider':
                result += 'VARIABLE LEVEL  %s_%sSCALE).\n' % (question.name, question.code)
            elif question.type == 'Composite' and question.subtype == 'SingleAnswer':
                grouped_questions.append(self.translate_composite(question, group_names))
            elif question.type == 'MC' and question.subtype == 'MAVR' and question.has_carry_forward_responses is False:
                grouped_questions.append(self.translate_mc_multiple(question, group_names))
            elif question.type == 'MC' and question.subtype == 'MAVR' and question.has_carry_forward_responses is True:
                print question.name
                grouped_questions.append(self.translate_mc_multiple_cf(question, group_names))
        result += self.add_groups(grouped_questions, group_names)
        return result

    def translate_composite(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='Select all that apply.'" % question.name
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for sub_question in question.questions:
            result += "%s " % sub_question.name
        result += "VALUE=1\n"
        return result

    def translate_mc_multiple(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='Select all that apply.'" % question.name
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for response in question.responses:
            result += "%s_%s " % (question.name, response.code)
        result += "VALUE=1\n"
        return result

    def translate_mc_multiple_cf(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='Select all that apply.'" % question.name
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for response in question.responses:
            result += "%s_x%s " % (question.name, response.code)
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
      
class RDSTranslator(object):

    def define_variables(self, survey, path_to_output):
        questions = survey.get_questions()
        file = open(path_to_output, "w+")
        file.write(self.translate_questions(questions))
        
    def translate_questions(self, questions):
        result = ''
        grouped_questions = []
        group_names = []
        for question in questions:
            if question.type == 'MC' and question.subtype in ['SAVR','SAHR','DL']:
                result += 'VARIABLE LEVEL  %s(ORDINAL).\n' % question.name
            elif question.type == 'Slider':
                result += 'VARIABLE LEVEL  %s_%sSCALE).\n' % (question.name, question.code)
            elif question.type == 'Composite' and question.subtype == 'SingleAnswer':
                grouped_questions.append(self.translate_composite(question, group_names))
            elif question.type == 'MC' and question.subtype == 'MAVR':
                grouped_questions.append(self.translate_mc_multiple(question, group_names))
        result += self.add_groups(grouped_questions, group_names)
        return result

    def translate_composite(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='Select all that apply.'" % question.name
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for sub_question in question.questions:
            result += "%s " % sub_question.name
        result += "VALUE=1\n"
        return result

    def translate_mc_multiple(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='Select all that apply.'" % question.name
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for response in question.responses:
            result += "%s_x%s " % (question.name, response.code)
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