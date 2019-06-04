import re
from functools import cmp_to_key
class Sorter(object):

    def __init__(self, sort_order):
        self.sort_order = str(sort_order)

    def sort(self, list_to_sort):
        return sorted(list_to_sort, key=cmp_to_key(self.compare))

class BlockSorter(Sorter):

    def __init__(self, sort_order):
        super(BlockSorter, self).__init__(sort_order)
    
    def compare(self, block1, block2):
        try:
            block1_location = self.sort_order.index(block1.blockid)
            block2_location = self.sort_order.index(block2.blockid)
            if block1_location > block2_location:
                return 1
            elif block1_location < block2_location:
                return -1
            else:
                return 0
        except:
            return 0     

class QuestionSorter(Sorter):
    
    def __init__(self, sort_order):
        super(QuestionSorter, self).__init__(sort_order)
    
    def compare(self, question1, question2):        
        id1_components = re.match('(QID\d+)(_\d+)?', question1.id)
        id2_components = re.match('(QID\d+)(_\d+)?', question2.id)
        id1_location = self.sort_order.index(id1_components.group(1))
        id2_location = self.sort_order.index(id2_components.group(1))
        if id1_location > id2_location:
            return 1
        elif id1_location < id2_location:
            return -1
        elif id1_components.group(2) > id2_components.group(2):
            return 1
        elif id1_components.group(2) < id2_components.group(2):
            return -1
        else:
            return 0

class CompositeQuestionSorter(Sorter):
    
    def __init__(self, sort_order):
        super(CompositeQuestionSorter, self).__init__(sort_order)
    
    def compare(self, question1, question2):
        try:
            question1_location = self.sort_order.index(question1.code)
            question2_location = self.sort_order.index(question2.code)
            if question1_location > question2_location:
                return 1
            elif question1_location < question2_location:
                return -1
            else:
                return 0    
        except ValueError:
            return 0

class ResponseSorter(Sorter):
    
    def __init__(self, sort_order):
        super(ResponseSorter, self).__init__(sort_order)
    
    def compare(self, response1, response2):
        try:
            response1_location = self.sort_order.index(response1.code)
            response2_location = self.sort_order.index(response2.code)
            if response1_location > response2_location:
                return 1
            elif response1_location < response2_location:
                return -1
            else:
                return 0
        except ValueError:
            return 0
