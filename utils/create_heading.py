from pptx.dml.color import RGBColor
from pptx.util import Pt
class Heading:
    def __init__(self,font_size,font_color,font,position,text,slide):
        self.font_size=font_size
        self.font_color=font_color
        self.position=position
        self.text=text
        self.font=font
        self.slide=slide
    def create_heading(self):
        heading_textBox=self.slide.shapes.add_textbox(*self.position)
        heading_paragraph=heading_textBox.text_frame.paragraphs[0]
        heading_paragraph.text=self.text
        heading_paragraph.font.color.rgb=RGBColor(*self.font_color)
        heading_paragraph.font.size=Pt(self.font_size)
        try:
            heading_paragraph.font.name=self.font  # Try except due to may be font isn't available in machine.
        except:
            heading_paragraph.font.name="Calibri"
        return