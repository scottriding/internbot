class Questions(object):
    
    def __init__(self, questions_data = []):
        self.__questions = {}
        for question_data in questions_data:
            self.add(question_data)
        
    def add(self, question_data):
        if self.already_exists(question_data['Name']):
            question = self.get(question_data['Name'])
            question.add_response(question_data['Response'], question_data['Frequency'])
        else:
            self.add_new(question_data)
    
    def add_new(self, question_data):
        q = Question(question_data['Name'], question_data['Prompt'], question_data['n'])
        q.add_response(question_data['Response'], question_data['Frequency'])
        self.__questions[q.name] = q
    
    def get(self, question_name):
        return self.__questions.get(question_name)
    
    def already_exists(self, question_name):
        if self.get(question_name) is None:
            return False
        else:
            return True
            
    def __len__(self):
        return len(self.__questions)
            
    def __repr__(self):
        result = ''
        for name, question in self.__questions.iteritems():
            result += str(question)
            result += '\n'
        return result
    
    def list_names(self):
        result = self.__questions.keys()
        return result
    
    def prompt(self, question_name):
        result = self.__questions[question_name].prompt
        return result
     
    def n(self, question_name):
        result = self.__questions[question_name].n
        return result
        
    def return_responses(self, question_name):
        result = self.__questions[question_name].return_response
        return result
        
    
class Question(object):
    
    def __init__(self, name, prompt, n):
        self.name      = name
        self.prompt    = prompt
        self.n         = n
        self.responses = {}
        
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
    def n(self):
        return self.__n
    
    @n.setter
    def n(self, n):
        self.__n = int(n)
    
    def add_response(self, response, freq):
        self.responses[str(response)] = float(freq)
    
    @property
    def return_response(self):
        return self.responses

        
    def __repr__(self):
        result = ""
        result += "%s: %s" % (self.name, self.prompt)
        for response, freq in self.responses.iteritems():
            result += ' %s: %s' % (response, str(freq))
        return result