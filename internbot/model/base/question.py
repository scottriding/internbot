from model.base import response
from model.base import sorter

class Questions(object):

    def __init__(self):
        self.__questions = []

    def add(self, question):
        self.__questions.append(question)

    def sort(self, question_id_order):
        q_sorter = sorter.QuestionSorter(question_id_order)
        self.__questions = q_sorter.sort(self.__questions)

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
        self.__temp_responses = response.Responses()
        self.__has_carry_forward_responses = False
        self.__has_mixed_responses = False
        self.__has_mixed_answers = False
        self.__has_mixed_statements = False
        self.__has_carry_forward_statements = False
        self.__has_carry_forward_answers = False

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
    def id(self):
        return self.__id
        
    @id.setter
    def id(self, id):
        self.__id = str(id)

    @property
    def subtype(self):
        return self.__subtype

    @subtype.setter
    def subtype(self, subtype):
        self.__subtype = str(subtype)

    @property
    def questions(self):
        return self.__questions

    @property
    def text_entry(self):
        return False

    @property
    def parent(self):
        return 'CompositeQuestion'

    @property
    def question_order(self):
        return self.__question_order

    @question_order.setter    
    def question_order(self, order):
        self.__question_order = order
        
    @property
    def carry_forward_question_id(self):
        return self.__carry_forward_question_id

    @carry_forward_question_id.setter
    def carry_forward_question_id(self, carry_forward_question_id):
        self.__carry_forward_question_id = str(carry_forward_question_id)

    @property
    def has_carry_forward_responses(self):
        return self.__has_carry_forward_responses

    @has_carry_forward_responses.setter
    def has_carry_forward_responses(self, type):
        self.__has_carry_forward_responses = bool(type)

    @property
    def temp_responses(self):
        return self.__temp_responses

    @property
    def has_mixed_responses(self):
        return self.__has_mixed_responses

    @has_mixed_responses.setter
    def has_mixed_responses(self, type):
        self.__has_mixed_responses = bool(type)

    @property
    def carry_forward_answers_id(self):
        return self.__carry_forward_answers_id

    @carry_forward_answers_id.setter
    def carry_forward_answers_id(self, carry_forward_answers_id):
        self.__carry_forward_answers_id = str(carry_forward_answers_id)

    @property
    def carry_forward_statements_id(self):
        return self.__carry_forward_statements_id

    @carry_forward_statements_id.setter
    def carry_forward_statements_id(self, carry_forward_statements_id):
        self.__carry_forward_statements_id = str(carry_forward_statements_id)

    @property
    def has_mixed_answers(self):
        return self.__has_mixed_answers

    @has_mixed_answers.setter
    def has_mixed_answers(self, type):
        self.__has_mixed_answers = bool(type)

    @property
    def has_mixed_statements(self):
        return self.__has_mixed_statements

    @has_mixed_statements.setter
    def has_mixed_statements(self, type):
        self.__has_mixed_statements = bool(type)

    @property
    def has_carry_forward_answers(self):
        return self.__has_carry_forward_answers

    @has_carry_forward_answers.setter
    def has_carry_forward_answers(self, type):
        self.__has_carry_forward_answers = bool(type)

    @property
    def has_carry_forward_statements(self):
        return self.__has_carry_forward_statements

    @has_carry_forward_statements.setter
    def has_carry_forward_statements(self, type):
        self.__has_carry_forward_statements = bool(type)
        
    def add_response(self, response, code):
        self.__temp_responses.add(response, code)
        self.sort()    
        
    def add_question(self, question):
        self.__questions.append(question)
        self.sort()
        
    def sort(self):
        cq_sorter = sorter.CompositeQuestionSorter(self.__question_order)
        if len(self.__questions) > 0:
            self.__questions = cq_sorter.sort(self.__questions)    

    def __len__(self):
        return len(self.__questions)

    def __iter__(self):
        return iter(self.__questions)    

    def __repr__(self):
        bool = True
        result = '%s: %s\n' % (self.name, self.prompt)
        for question in self.__questions:    
            result += "\n\t\t%s" % (str(question))
        return result

class CompositeMatrix(CompositeQuestion):

    def __init__(self):
        super(CompositeMatrix, self).__init__()

    @property
    def type(self):
        return 'CompositeMatrix'

class CompositeMultipleSelect(CompositeQuestion):

    def __init__(self):
        super(CompositeMultipleSelect, self).__init__()

    @property
    def type(self):
        return 'CompositeMultipleSelect'
        
class CompositeHotSpot(CompositeQuestion):

    def __init__(self):
        super(CompositeHotSpot, self).__init__()    

    @property
    def type(self):
        return 'CompositeHotSpot'

class CompositeConstantSum(CompositeQuestion):

    def __init__(self):
        super(CompositeConstantSum, self).__init__()    

    @property
    def type(self):
        return 'CompositeConstantSum'

class Question(object):

    def __init__(self):
        self.__responses = response.Responses()
        self.__response_order = []
        self.has_carry_forward_responses = False
        self.__text_entry = False
        self.__has_mixed_responses = False
        self.__has_carry_forward_statements = False
        self.__has_carry_forward_answers = False
        self.__n = 0
        self.__display_logic = None
        self.__stat = None

    @property
    def has_carry_forward_statements(self):
        return self.__has_carry_forward_statements

    @property
    def has_carry_forward_answers(self):
        return self.__has_carry_forward_answers

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

    @property
    def text_entry(self):
        return self.__text_entry

    @text_entry.setter
    def text_entry(self, entry):
        self.__text_entry = bool(entry)
    
    @property
    def parent(self):
        return None

    @property
    def has_mixed_responses(self):
        return self.__has_mixed_responses

    @has_mixed_responses.setter
    def has_mixed_responses(self, type):
        self.__has_mixed_responses = bool(type)

    @property
    def display_logic(self):
        return self.__display_logic

    @display_logic.setter
    def display_logic(self, logic):
        self.__display_logic = str(logic)

    @property
    def stat(self):
        return self.__stat

    @stat.setter
    def stat(self, stat):
        self.__stat = str(stat)

    def add_response(self, response, code=None):
        self.__responses.add(response, code)
        self.__responses.sort(self.__response_order)

    def add_dynamic_response(self, response, code=None):
        self.__responses.add_dynamic(response, code)
        self.__responses.sort(self.__response_order)

    def add_text_response(self, response):
        self.__responses.add_text(response)


    def add_NA(self):
        self.__responses.add_NA()

    def get_NA(self):
        return self.__responses.get_NA()

    def __repr__(self):
        result = ''
        result += "%s: %s\n" % (self.name, self.prompt)
        result += str(self.__responses)
        return result
