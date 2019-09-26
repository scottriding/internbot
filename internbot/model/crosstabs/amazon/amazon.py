from model.crosstabs.amazon import rename_xlsx_tabs
from model.crosstabs.amazon import highlighter

class Amazon(object):

    def __init__(self):
        self.__renamer = rename_xlsx_tabs.RenameTabs()
        self.__highlighter = highlighter.Highlighter()

    def rename(self, path_to_xlsx, path_to_toc, resources_filepath):
        return self.__renamer.rename(path_to_xlsx, path_to_toc, resources_filepath)

    def highlight(self, workbook, path_to_output, is_trended_amazon):
        self.__highlighter.highlight(workbook, path_to_output, is_trended_amazon)

