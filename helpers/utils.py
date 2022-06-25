from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
import math
import re
import pydash

def check_tag_exist(tag, shape):
    matches = tag in shape.text
    return matches

def remove_tags(tag, shape):
    matches = tag in shape.text
    return matches

def get_tag_content(pattern,shape):
    matches = re.findall(pattern, shape.text)
    return matches

def eval_executor():
    eval('CHART_NAME + TABLE_NAME',replacements)