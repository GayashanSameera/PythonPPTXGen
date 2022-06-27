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

def replace_tags(replaced_for,replaced_text, shape):
    if shape.has_text_frame:
        text_frame = shape.text_frame
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                cur_text = run.text
                new_text = cur_text.replace(replaced_for, replaced_text)
                run.text = new_text

    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                if match in cell.text:
                    new_text = cell.text.replace(replaced_for, replaced_text)
                    cell.text = new_text

def get_tag_content(pattern,shape):
    matches = re.findall(pattern, shape.text)
    return matches

def get_tag_from_string(pattern,string):
    matches = re.findall(pattern, string)
    return matches

def eval_executor(logic, replacements):
    return eval(logic,replacements)