from question import Questions
from sorter import BlockSorter
import re

class Blocks(object):

    def __init__(self):
        self.__blocks = []

    def add(self, block):
        self.__blocks.append(block)

    def find_by_assigned_id(self, question_id):
        for block in self.__blocks:
            if block.is_id_assigned(question_id):
                return block
        return None
        
    def sort(self, block_id_order):
        sorter = BlockSorter(block_id_order)
        self.__blocks = sorter.sort(self.__blocks)    

    def __repr__(self):
        result = ''
        for block in self.__blocks:
            result += "\t%s\n" % str(block)
        return result

    def __iter__(self):
        return(iter(self.__blocks))

class Block(object):

    def __init__ (self, block_name):
        self.__assigned_ids = []
        self.__questions = Questions()
        self.name = block_name

    def assign_id(self, question_id):
        self.__assigned_ids.append(question_id)

    def is_id_assigned(self, question_id):
        for id in self.__assigned_ids:
            if re.match('(%s)(?:\_|$)(\d+)?' % id, question_id):
                return True
        return False

    def add_question(self, question):
        self.__questions.add(question)
        self.__questions.sort(self.__assigned_ids)
    
    @property
    def blockid(self):
        return self.__blockid
     
    @blockid.setter
    def blockid(self, id):
        self.__blockid = str(id)
        

    def __repr__(self):
        result = "Block: %s\n" % (self.name)
        result += str(self.__questions)
        return result
