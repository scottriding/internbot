from question import Questions

class Blocks(object):

    def __init__(self):
        self.__blocks = []

    def add(self, block):
        self.__blocks.append(block)

    def __repr__(self):
        result = ''
        for block in self.__blocks:
            result += "\t%s\n" % str(block)
        return result

    def __iter__(self):
        return(iter(self.__blocks))

class Block(object):

    def __init__ (self, block_name):
        self.name = block_name

    def __repr__(self):
        return "Block: %s" % (self.name)
