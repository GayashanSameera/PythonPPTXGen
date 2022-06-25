from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
import math
import re
import pydash

from helpers.utils import check_tag_exist, replace_tags, get_tag_content

def text_replace(slide, shape,replacements):
    
    pattern = r'\+\+\+INS (.*?) \+\+\+'
    matches = get_tag_content(pattern, shape)

    if( not matches or len(matches) < 1):
        return

    for match in matches:
        object_value = pydash.get(replacements, match)
        replace_tags(str(f"+++INS {match} +++"), str(object_value), shape)