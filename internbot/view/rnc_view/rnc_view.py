from view.rnc_view.scores_topline_view import ScoresToplineView
from view.rnc_view.issue_trended_view import IssueTrendedView
from view.rnc_view.trended_scores_view import TrendedScoresView

## outside modules
import kivy
kivy.require('1.11.1')

from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.core.text import LabelBase
from kivy.uix.relativelayout import RelativeLayout
import webbrowser


class RNCView(BoxLayout):

    def __init__(self, **kwargs):
        super(RNCView, self).__init__(**kwargs)

        self.__controller = None

        self.rnc_screen = self.create_rnc_screen()
        self.scores_screen = ScoresToplineView()
        self.issue_screen = IssueTrendedView()
        self.trended_screen = TrendedScoresView()

        self.add_widget(self.rnc_screen)

    def create_rnc_screen(self):
        rnc_screen = BoxLayout(orientation='vertical')

        back_arrow = "<"
        self.__back_btn = Button(text=back_arrow, size_hint=(.1, .05))
        self.__back_btn.font_name = "Y2"

        rnc_screen.add_widget(self.__back_btn)

        help_btn = Button(text='Help', size_hint=(.1, .05), on_press = self.rnc_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (.9529, 0.6, .2392, 1)

        rnc_screen.add_widget(help_btn)

        button_layout = BoxLayout()

        scores_btn = Button(text='Scores Topline Report', size_hint=(.1, .1))
        scores_btn.font_name = "Y2"
        scores_btn.bind(on_press=self.build_scores)

        issue_btn = Button(text='Issue Trended Report', size_hint=(.1, .1))
        issue_btn.font_name = "Y2"
        issue_btn.bind(on_press=self.build_issue)

        trended_btn = Button(text='Trended Scores Reports', size_hint=(.1, .1))
        trended_btn.font_name = "Y2"
        trended_btn.bind(on_press=self.build_tsr)

        button_layout.add_widget(scores_btn)
        button_layout.add_widget(issue_btn)
        button_layout.add_widget(trended_btn)

        rnc_screen.add_widget(button_layout)

        return rnc_screen

    @property
    def back_button(self):
        return self.__back_btn

    @property
    def controller(self):
        return self.__controller

    @controller.setter
    def controller(self, controller):
        self.__controller = controller

    def build_scores(self, instance):
        self.scores_screen.run(self.__controller)

    def build_issue(self, instance):
        self.issue_screen.run(self.__controller)

    def build_tsr(self, instance):
        self.trended_screen.run(self.__controller)

    def rnc_help(self, instance):
        help_text = "RNC reports are deliverables on model scoring \n"
        help_text += "for a targeted region designated by the RNC.\n\n" 
        help_text += "Scores Topline Report are high-level net model results.\n\n" 
        help_text += "Issue Trended Report breaks down models by RNC flags.\n\n" 
        help_text += "Trended Scores Reports are several report models\n" 
        help_text += "by flags and by weights.\n\n"
        help_text += "[ref=click][color=F3993D]Click here for RNC report examples[/color][/ref]"
        help_text += "\n\n\n\n"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/bmr116skqrlvp1l/AAC9xm56MxnmdRq6TYr_MuPka?dl=0")
        
        help_content = RelativeLayout()
        content_cancel = Button(text='confirm', 
						pos_hint={'center_x': 0.5, 'center_y': 0.15}, 
        				size_hint=(.2, .15),
        				background_normal='',
        				background_color=(0, 0.4, 1, 1))
        help_content.add_widget(content_cancel)
        help_label = (Label(text=help_text, markup=True))
        help_label.bind(on_ref_press=examples_link)
        help_label.font_family = "Y2"
        help_content.add_widget(help_label)
		
        popup = Popup(title='RNC Help',
        content=help_content,
        auto_dismiss=False,
        size_hint=(.6, .7), pos_hint={'center_x': 0.5, 'center_y': 0.5})
        

        popup.open()
        content_cancel.bind(on_press=popup.dismiss)
