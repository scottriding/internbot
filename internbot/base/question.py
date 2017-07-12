from response import Responses
from sorter import QuestionSorter, CompositeQuestionSorter


class Questions(object):

    def __init__(self):
        self.__questions = []

    def add(self, question):
        self.__questions.append(question)

    def sort(self, question_id_order):
        sorter = QuestionSorter(question_id_order)
        self.__questions = sorter.sort(self.__questions)

    def __len__(self):
        return len(self.__questions)

    def __iter__(self):
        return iter(self.__questions)

    def __repr__(self):
        result = ''
        for question in self.__questions:
            result += "\t\t%s\n" % (str(question))
        return result
        
class CompositeQuestion(object):

    def __init__(self):
        self.__questions = []
        self.__question_order = []
        self.__temp_responses = Responses()
        self.has_carry_forward_prompts = False
        self.has_carry_forward_responses = False 
        
    @property
    def id(self):
        return self.__id
        
    @id.setter
    def id(self, id):
        self.__id = str(id)
    
    @property
    def name(self):
        return self.__name
    
    @name.setter
    def name(self, name):
        self.__name = str(name)
    
    @property
    def prompt(self):
        return self.__prompt
    
    @prompt.setter
    def prompt(self, prompt):
        self.__prompt = str(prompt)
    
    @property
    def type(self):
        return self.__type

    @type.setter
    def type(self):
        self.__type = 'Composite' 
        
    @property
    def subtype(self):
        return self.__subtype

    @subtype.setter
    def subtype(self, subtype):
        self.__subtype = str(subtype)    

    @property
    def question_order(self):
        return self.__question_order

    @question_order.setter    
    def question_order(self, order):
        self.__question_order = order

    @property
    def has_carry_forward_prompts(self):
        return self.__has_carry_forward_prompts

    @has_carry_forward_prompts.setter
    def has_carry_forward_prompts(self, has_carry_forward_prompts):
        self.__has_carry_forward_prompts = bool(has_carry_forward_prompts)
        
    @property
    def carry_forward_question_id(self):
        return self.__carry_forward_question_id

    @carry_forward_question_id.setter
    def carry_forward_question_id(self, carry_forward_question_id):
        self.__carry_forward_question_id = str(carry_forward_question_id)

    @property
    def temp_responses(self):
        return self.__temp_responses
        
    def add_response(self, response, code):
        self.__temp_responses.add(response, code)
        self.sort()    
        
    def add_question(self, question):
        self.__questions.append(question)
        self.sort()
        
    def sort(self):
        sorter = CompositeQuestionSorter(self.__question_order)
        if len(self.__questions) > 0:
            self.__questions = sorter.sort(self.__questions)    

    def __len__(self):
        return len(self.__questions)

    def __iter__(self):
        return iter(self.__questions)    

    def __repr__(self):
        bool = True
        result = ''
        for question in self.__questions:
            if bool == True:
                result += '%s' % (str(question))
                bool = False
            else:    
                result += "\t\t%s\n" % (str(question))
        return result        
         

class Question(object):

    def __init__(self):
        self.__responses = Responses()
        self.__response_order = []
        self.has_carry_forward_responses = False
        self.has_carry_forward_prompts = False

    @property
    def id(self):
        return self.__id

    @id.setter
    def id(self, id):
        self.__id = str(id)

    @property
    def name(self):
        return self.__name

    @name.setter
    def name(self, name):
        self.__name = str(name)

    @property
    def type(self):
        return self.__type

    @type.setter
    def type(self, type):
        self.__type = str(type)

    @property
    def subtype(self):
        return self.__subtype

    @subtype.setter
    def subtype(self, subtype):
        self.__subtype = str(subtype)
        
    @property
    def code(self):
        return self.__code
    
    @code.setter
    def code(self, code):
        self.__code = str(code)
        
    @property
    def has_carry_forward_responses(self):
        return self.__has_carry_forward_responses

    @has_carry_forward_responses.setter
    def has_carry_forward_responses(self, has_carry_forward_responses):
        self.__has_carry_forward_responses = bool(has_carry_forward_responses)

    @property
    def carry_forward_question_id(self):
        return self.__carry_forward_question_id

    @carry_forward_question_id.setter
    def carry_forward_question_id(self, carry_forward_question_id):
        self.__carry_forward_question_id = str(carry_forward_question_id)

    @property
    def prompt(self):
        return self.__prompt

    @prompt.setter
    def prompt(self, prompt):
        self.__prompt = str(prompt)

    @property
    def n(self):
        return self.__n

    @n.setter
    def n(self, n):
        self.__n = int(n)

    @property
    def responses(self):
        return self.__responses

    @property
    def response_order(self):
        return self.__response_order

    @response_order.setter
    def response_order(self, response_order):
        for response_location in response_order:
            self.__response_order.append(str(response_location))

    def spss_name(self):
        if self.type == 'MC' and self.subtype in ['SAVR','SAHR']:
            return self.name
        elif self.type == 'Matrix' and self.subtype == 'SingleAnswer':
            return self.name
        elif self.type == 'MC' and self.subtype == 'MAVR':
            names = []
            for response in self.responses:
                names.append('%s_%s' % (self.name, response.code))
            return names
        elif self.type == 'DB':
            return self.name
        else:
            return self.name

    def add_response(self, response, code=None):
        self.__responses.add(response, code)
        #self.__responses.sort(self.__response_order)

    def __repr__(self):
        result = ''
        result += "Question: %s\n" % self.name
        result += str(self.__responses)
        return result
        


