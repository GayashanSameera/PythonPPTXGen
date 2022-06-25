from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
import math
import re
import pydash

from helpers.utils import replace_tags, get_tag_content, get_tag_from_string

def text_replace(slide, shape,replacements):
    
    pattern = r'\+\+\+INS (.*?) \+\+\+'
    matches = get_tag_content(pattern, shape)
    if( not matches or len(matches) < 1):
        return

    for match in matches:
        object_value = pydash.get(replacements, match)
        replace_tags(str(f"+++INS {match} +++"), str(object_value), shape)

def text_tag_update(text,replacements):
    current_text = text
    pattern = r'\+\+\+INS (.*?) \+\+\+'
    matches = get_tag_from_string(pattern, text)
    if( not matches or len(matches) < 1):
        return { "text": text }

    for match in matches:
        object_value = pydash.get(replacements, match, False)
        if(object_value != False):
            current_text = current_text.replace(str(f"+++INS {match} +++"), str(object_value))

    
    return { "text": current_text }