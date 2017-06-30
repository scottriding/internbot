from question import Questions

class Blocks(object):
    pass

class Block(object):
    def __init__ (self, block_name, block_q_count, block_q_order):
        self.name = block_name
        self.q_count = block_q_count
        self.q_order = block_q_order

    def add_question (self, question):
        pass
        
    def __repr__(self):
        result = ""
        result += "%s: %s" % (self.name, str(self.q_count))
        return result