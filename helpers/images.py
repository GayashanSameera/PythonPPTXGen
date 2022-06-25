from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
import math
import re
import pydash

from helpers.utils import check_tag_exist, remove_tags, get_tag_content

def replace_images(slide, shape, replacements):
    pattern = r'\+\+\+CHART (.*?) \+\+\+'
    matches = get_tag_content(pattern, shape)

    if( not matches or len(matches) < 1):
        return

    text_frame = shape.text_frame

    for match in matches:
        object_value = pydash.get(replacements, match)

        url = pydash.get(object_value, "url")
        left = pydash.get(object_value, "size.left")
        height = pydash.get(object_value, "size.height")
        top = pydash.get(object_value, "size.top")
        width = pydash.get(object_value, "size.width")
        
        slide.shapes.add_picture(url, Inches(left), Inches(top), Inches(width) ,Inches(height) )

        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                cur_text = run.text
                new_text = cur_text.replace(str(f"+++CHART {match} +++"), "")
                run.text = new_text