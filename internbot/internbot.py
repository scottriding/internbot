## internbot modules
import base
import crosstabs
import rnc_automation
import topline

## outside modules
import kivy
kivy.require('1.11.1')

from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.image import Image
from kivy.graphics import Color, Rectangle
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.core.text import LabelBase
import webbrowser


class MainApp(App):

    def build(self):
        self.root = self.create_screens()
        self.title = "Internbot"
        self.root.bind(size=self._update_rect, pos=self._update_rect)

        with self.root.canvas.before:
            self.rect = Rectangle(size=self.root.size, pos=self.root.pos, source='data/images/DWTWNSLC.png')
        return self.root

    def _update_rect(self, instance, value):
        self.rect.pos = instance.pos
        self.rect.size = instance.size

    def create_screens(self):
        layered_menu = BoxLayout()

        self.main_screen = self.create_main_screen()
        self.crosstabs_screen = self.create_crosstabs_screen()
        self.topline_screen = self.create_topline_screen()
        self.qresearch_screen = self.create_qresearch_screen()
        self.spss_screen = self.create_spss_screen()
        self.rnc_screen = self.create_rnc_screen()

        layered_menu.add_widget(self.main_screen)
        return layered_menu

    def create_main_screen(self):
        main_screen = BoxLayout()

        ## buttons
        button_layout = BoxLayout(orientation='vertical')

        xtabs_btn = Button(text="Crosstab Reports", size_hint=(.5,.6), on_press = self.main_to_xtabs)
        xtabs_btn.font_name = "Y2"

        top_btn = Button(text="Topline Reports", size_hint=(.5,.6), on_press = self.main_to_top)
        top_btn.font_name = "Y2"
 
        rnc_btn = Button(text="RNC Reports", size_hint=(.5,.6), on_press = self.main_to_rnc)
        rnc_btn.font_name = "Y2"
        
        help_btn = Button(text="Help", size_hint=(.5,.15), on_press = self.main_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (0.07, 0.306, 0.651, 1)

        button_layout.add_widget(xtabs_btn)
        button_layout.add_widget(top_btn)
        button_layout.add_widget(rnc_btn)
        button_layout.add_widget(help_btn)

        main_screen.add_widget(button_layout)

        # y2 logo
        screen_image = Image()
        screen_image.source = "data/images/y2_white_logo.png"
        screen_image.post_hint = {'center_x': 0, 'center_y': .5}

        main_screen.add_widget(screen_image)

        return main_screen

    def create_crosstabs_screen(self):
        crosstabs_screen = BoxLayout(orientation='vertical')

        back_arrow = "<"
        back_btn = Button(text=back_arrow, size_hint=(.1, .05))
        back_btn.bind(on_press=self.xtabs_to_main)
        back_btn.font_name = "Y2"

        crosstabs_screen.add_widget(back_btn)

        help_btn = Button(text='Help', size_hint=(.1, .05), on_press = self.xtabs_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (.9529, 0.6, .2392, 1)

        crosstabs_screen.add_widget(help_btn)

        button_layout = BoxLayout()

        qresearch_btn = Button(text='QResearch', size_hint=(.3, .2), on_press = self.xtabs_to_qresearch)
        qresearch_btn.font_name = "Y2"
        
        spss_btn = Button(text='SPSS', size_hint=(.3, .2), on_press = self.xtabs_to_spss)
        spss_btn.font_name = "Y2"
        spss_btn.disabled = True
        
        button_layout.add_widget(qresearch_btn)
        button_layout.add_widget(spss_btn)

        crosstabs_screen.add_widget(button_layout)

        return crosstabs_screen

    def create_qresearch_screen(self):
        qresearch_screen = BoxLayout(orientation='vertical')

        back_arrow = "<"
        back_btn = Button(text=back_arrow, size_hint=(.1, .05))
        back_btn.bind(on_press=self.qresearch_to_xtabs)
        back_btn.font_name = "Y2"

        qresearch_screen.add_widget(back_btn)

        double_back_arrow = "<<"
        double_back_btn = Button(text=double_back_arrow, size_hint=(.1, .05))
        double_back_btn.bind(on_press=self.qresearch_to_main)
        double_back_btn.font_name = "Y2"

        qresearch_screen.add_widget(double_back_btn)

        help_btn = Button(text='Help', size_hint=(.1, .05), on_press = self.qresearch_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (.9529, 0.6, .2392, 1)

        qresearch_screen.add_widget(help_btn)

        button_layout = BoxLayout()
        toc_btn = Button(text='Create table of contents', size_hint=(.5, .2), on_press = self.create_toc)
        toc_btn.font_name = "Y2"
        
        format_btn = Button(text='Format report', size_hint=(.5, .2), on_press = self.format_qresearch_xtabs)
        format_btn.font_name = "Y2"

        button_layout.add_widget(toc_btn)
        button_layout.add_widget(format_btn)

        qresearch_screen.add_widget(button_layout)

        return qresearch_screen

    def create_spss_screen(self):
        spss_screen = BoxLayout(orientation='vertical')

        back_arrow = "<"
        back_btn = Button(text=back_arrow, size_hint=(.1, .05))
        back_btn.bind(on_press=self.spss_to_xtabs)
        back_btn.font_name = "Y2"

        spss_screen.add_widget(back_btn)

        double_back_arrow = "<<"
        double_back_btn = Button(text=double_back_arrow, size_hint=(.1, .05))
        double_back_btn.bind(on_press=self.spss_to_main)
        double_back_btn.font_name = "Y2"

        spss_screen.add_widget(double_back_btn)

        help_btn = Button(text='Help', size_hint=(.1, .05), on_press = self.spss_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (.9529, 0.6, .2392, 1)

        spss_screen.add_widget(help_btn)

        button_layout = BoxLayout()
        variable_btn = Button(text='Create variable script', size_hint=(.5, .2), on_press = self.create_spss_variables)
        variable_btn.font_name = "Y2"
        
        table_btn = Button(text='Create table script', size_hint=(.5, .2), on_press = self.create_spss_tables)
        table_btn.font_name = "Y2"

        build_btn = Button(text='Build report', size_hint=(.5, .2), on_press = self.build_spss_report)
        build_btn.font_name = "Y2"

        button_layout.add_widget(variable_btn)
        button_layout.add_widget(table_btn)
        button_layout.add_widget(build_btn)

        spss_screen.add_widget(button_layout)

        return spss_screen

    def create_topline_screen(self):
        topline_screen = BoxLayout(orientation='vertical')

        back_arrow = "<"
        back_btn = Button(text=back_arrow, size_hint=(.1, .05))
        back_btn.bind(on_press=self.top_to_main)
        back_btn.font_name = "Y2"

        topline_screen.add_widget(back_btn)

        help_btn = Button(text='Help', size_hint=(.1, .05), on_press = self.top_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (.9529, 0.6, .2392, 1)

        topline_screen.add_widget(help_btn)

        button_layout = BoxLayout()

        docx_btn = Button(text='Appendix', size_hint=(.5, .2))
        docx_btn.font_name = "Y2"

        pptx_btn = Button(text='Document', size_hint=(.5, .2))
        pptx_btn.font_name = "Y2"

        app_btn = Button(text='Powerpoint', size_hint=(.5, .2))
        app_btn.font_name = "Y2"

        button_layout.add_widget(docx_btn)
        button_layout.add_widget(pptx_btn)
        button_layout.add_widget(app_btn)

        topline_screen.add_widget(button_layout)

        return topline_screen

    def create_rnc_screen(self):
        rnc_screen = BoxLayout(orientation='vertical')

        back_arrow = "<"
        back_btn = Button(text=back_arrow, size_hint=(.1, .05))
        back_btn.bind(on_press=self.rnc_to_main)
        back_btn.font_name = "Y2"

        rnc_screen.add_widget(back_btn)

        help_btn = Button(text='Help', size_hint=(.1, .05), on_press = self.rnc_help)
        help_btn.font_name = "Y2"
        help_btn.background_normal = ''
        help_btn.background_color = (.9529, 0.6, .2392, 1)

        rnc_screen.add_widget(help_btn)

        button_layout = BoxLayout()

        scores_btn = Button(text='Scores Topline Report', size_hint=(.5, .2))
        scores_btn.font_name = "Y2"

        issues_btn = Button(text='Issue Trended Report', size_hint=(.5, .2))
        issues_btn.font_name = "Y2"

        trended_btn = Button(text='Trended Scores Reports', size_hint=(.5, .2))
        trended_btn.font_name = "Y2"

        button_layout.add_widget(scores_btn)
        button_layout.add_widget(issues_btn)
        button_layout.add_widget(trended_btn)

        rnc_screen.add_widget(button_layout)

        return rnc_screen

    def create_toc(self, instance):
        pass

    def format_qresearch_xtabs(self, instance):
        pass

    def create_spss_variables(self, instance):
        pass

    def create_spss_tables(self, instance):
        pass

    def build_spss_report(self, instance):
        pass

    def main_to_xtabs(self, instance):
        self.root.remove_widget(self.main_screen)
        self.root.add_widget(self.crosstabs_screen)
        self.rect.source = 'data/images/DWTWNSLC1.jpg'

    def xtabs_to_qresearch(self, instance):
        self.root.remove_widget(self.crosstabs_screen)
        self.root.add_widget(self.qresearch_screen)

    def xtabs_to_spss(self, instance):
        self.root.remove_widget(self.crosstabs_screen)
        self.root.add_widget(self.spss_screen)

    def qresearch_to_xtabs(self, instance):
        self.root.remove_widget(self.qresearch_screen)
        self.root.add_widget(self.crosstabs_screen)

    def spss_to_xtabs(self, instance):
        self.root.remove_widget(self.spss_screen)
        self.root.add_widget(self.crosstabs_screen)

    def main_to_top(self, instance):
        self.root.remove_widget(self.main_screen)
        self.root.add_widget(self.topline_screen)
        self.rect.source = 'data/images/DWTWNSLC2.jpg'

    def main_to_rnc(self, instance):
        self.root.remove_widget(self.main_screen)
        self.root.add_widget(self.rnc_screen)
        self.rect.source = 'data/images/DWTWNSLC3.jpg'

    def xtabs_to_main(self, instance):
        self.root.remove_widget(self.crosstabs_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'data/images/DWTWNSLC.png'

    def qresearch_to_main(self, instance):
        self.root.remove_widget(self.qresearch_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'data/images/DWTWNSLC.png'

    def spss_to_main(self, instance):
        self.root.remove_widget(self.spss_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'data/images/DWTWNSLC.png'

    def top_to_main(self, instance):
        self.root.remove_widget(self.topline_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'data/images/DWTWNSLC.png'

    def rnc_to_main(self, instance):
        self.root.remove_widget(self.rnc_screen)
        self.root.add_widget(self.main_screen)
        self.rect.source = 'data/images/DWTWNSLC.png'

    def main_help(self, instance):
        help_text = "Internbot is Y2 Analytics' report automation tool that \n" 
        help_text += "manages \"quantity\" in a report so that the analytical \n"
        help_text += "team can focus on the \"quality\" of a report.\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of reports[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/bbcle8f9nifk4bo/AAA-JGnsx1XnhLaD_5Z9oDZna?dl=0")


        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family = "Y2"

        popup = Popup(title='Internbot Help',
        content=label,
        size_hint=(None, None), size=(400, 400))

        popup.open()

    def xtabs_help(self, instance):
        help_text = "Crosstabs are a report that \"crosses\" selected dataset \n" 
        help_text += "variables against the rest of the dataset variables \n"
        help_text += "visualized in a table form with statistically significant \n"
        help_text += "values highlighted.\n\n"
        help_text += "QResearch and SPSS are different crosstab generating \n"
        help_text += "software that require different automation from internbot.\n\n"
        help_text += "[ref=click][color=F3993D]Click here for examples of crosstab reports[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/2b67i3hbj1teg5j/AACWbEIwGtqq5hK894ItmmGpa?dl=0")


        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family = "Y2"

        popup = Popup(title='Crosstabs Help',
        content=label,
        size_hint=(None, None), size=(415, 400))

        popup.open()

    def qresearch_help(self, instance):
        help_text = "QResearch is statistical software located on the new office PC.\n\n"
        help_text += "Before building a report in QResearch, you will need the \n"
        help_text += "survey file (.qsf) of the project to produce a table of contents\n"
        help_text += "in internbot (Create table of contents) and the column variables \n"  
        help_text += "or banners for building the report in the QResearch software. \n\n"
        help_text += "Once the QResearch unformatted report is generated in QResearch \n"
        help_text += "internbot will format it to Qualtrics/Y2 standards (Format report).\n\n"
        help_text += "[ref=click][color=F3993D]Click here for QResearch files and further instructions[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/luixf0mkcfxlfjc/AAC3ncnz0dAdUdQjdJXTeCEsa?dl=0")


        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family = "Y2"

        popup = Popup(title='QResearch Help',
        content=label,
        size_hint=(None, None), size=(480, 400))

        popup.open()

    def spss_help(self, instance):
        help_text = "SPSS is statistical software located on the old office PC.\n\n"
        help_text += "Before building a report in SPSS, you will need the \n"
        help_text += "survey file (.qsf) of the project to produce a variable script\n"
        help_text += "in internbot (Create variable script). And the column variables \n"  
        help_text += "or banners for creating table script in internbot\n" 
        help_text += "(Create table script). Both files will be used in generating\n" 
        help_text += "several table files in a folder that internbot will use to create\n"
        help_text += "the final report (Build report).\n\n"
        help_text += "[ref=click][color=F3993D]Click here for SPSS files and further instructions[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("")


        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family = "Y2"

        popup = Popup(title='SPSS Help',
        content=label,
        size_hint=(None, None), size=(480, 400))

        popup.open()

    def top_help(self, instance):
        help_text = "Topline reports are deliverables that give high level results\n"
        help_text += "for a survey project.\n\n"
        help_text += "An Appendix report shows the results of text entries or\n"
        help_text += "open-ended questions.\n\n"
        help_text += "Document and Powerpoint show topline frequencies verbatim\n"
        help_text += "in tables or charts, respectively.\n\n"
        help_text += "[ref=click][color=F3993D]Click here for Topline report examples[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/c653vu8jxx0d42u/AADlcD8HDL2J4HxXqDVT3bAVa?dl=0")


        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family = "Y2"

        popup = Popup(title='Topline Help',
        content=label,
        size_hint=(None, None), size=(480, 400))

        popup.open()

    def rnc_help(self, instance):
        help_text = "RNC reports are deliverables on model scoring \n"
        help_text += "for a targeted region designated by the RNC.\n\n" 
        help_text += "Scores Topline Report are high-level net model results.\n\n" 
        help_text += "Issue Trended Report breaks down models by RNC flags.\n\n" 
        help_text += "Trended Scores Reports are several report models\n" 
        help_text += "by flags and by weights.\n\n"
        help_text += "[ref=click][color=F3993D]Click here for RNC report examples[/color][/ref]"

        def examples_link(instance, value):
            webbrowser.open("https://www.dropbox.com/sh/bmr116skqrlvp1l/AAC9xm56MxnmdRq6TYr_MuPka?dl=0")


        label = Label(text=help_text, markup=True)
        label.bind(on_ref_press=examples_link)
        label.font_family = "Y2"

        popup = Popup(title='RNC Help',
        content=label,
        size_hint=(None, None), size=(480, 400))

        popup.open()

if __name__ == '__main__':
    KIVY_FONTS = [
    {
        "name": "Y2",
        "fn_regular": "data/fonts/GothamBook.otf"
    }
    ]
    for font in KIVY_FONTS:
        LabelBase.register(**font)
    MainApp().run()