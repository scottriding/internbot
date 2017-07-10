import re

class Sorter(object):

    def __init__(self, sort_order, type):
        self.__sort_order = sort_order
        self.__type = type
        
    def sort(self, list_to_sort):
        if self.__type == 'questions':
            return sorted(list_to_sort, cmp=self.question_compare)
        elif self.__type == 'responses':
            return sorted(list_to_sort, cmp=self.response_compare)
        elif self.__type == 'blocks':
            return sorted(list_to_sort, cmp=self.block_compare)
        elif self.__type == 'prompts':
            pass
    
    def question_compare(self, question1, question2):
        id1_components = re.match('(QID\d+)(_\d+)?', question1.id)
        id2_components = re.match('(QID\d+)(_\d+)?', question2.id)
        id1_location = self.__sort_order.index(id1_components.group(1))
        id2_location = self.__sort_order.index(id2_components.group(1))
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
            
    def response_compare(self, response1, response2):        
        response1_location = self.__sort_order.index(response1.code)
        response2_location = self.__sort_order.index(response2.code)
        if response1_location > response2_location:
            return 1
        elif response1_location < response2_location:
            return -1
        else:
            return 0  

    def block_compare(self, block1, block2):
        block1_location = self.__sort_order.index(block1.blockid)
        block2_location = self.__sort_order.index(block2.blockid)
        if block1_location > block2_location:
            return 1
        elif block1_location < block2_location:
            return -1
        else:
            return 0 