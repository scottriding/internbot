import unittest
import sys
sys.path.append("..")
import base
import os
import traceback

class TestStringMethods(unittest.TestCase):

    def test_survey_compiler(self):
        self.file = ""
        try:
            for dirname, dirnames, filenames in os.walk("qsf"):
                for file in filenames:
                    print("================ RUNNING FILE "+str(file)+" ===============")
                    self.file = file
                    file_path = os.path.join(os.path.expanduser("~/Documents/GitHub/internbot/internbot/test/qsf"), file)
                    if file_path.endswith('.qsf'):
                        survey = base.QSFSurveyCompiler().compile(file_path)
        except:
            traceback.print_exc()
            self.fail("QSF FILE "+str(self.file)+"  raised an exception while compiling")



if __name__ == '__main__':
    unittest.main()