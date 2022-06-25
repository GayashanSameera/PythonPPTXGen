from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
import math
import re
import pydash

from helpers.utils import check_tag_exist, remove_tags, get_tag_content

def text_replace(slide, shape,replacements):
    
    pattern = r'\+\+\+INS (.*?) \+\+\+'
    matches = get_tag_content(pattern, shape)

    if( not matches or len(matches) < 1):
        return

    for match in matches:
        object_value = pydash.get(replacements, match)
        text_frame = shape.text_frame

        if shape.has_text_frame:
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    cur_text = run.text
                    new_text = cur_text.replace(str(f"+++INS {match} +++"), str(object_value))
                    run.text = new_text

        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if match in cell.text:
                        new_text = cell.text.replace(str(f"+++INS {match} +++"), str(object_value))
                        cell.text = new_text