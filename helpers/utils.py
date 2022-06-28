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

def delete_slides(presentation, index):
    xml_slides = presentation.slides._sldIdLst  
    slides = list(xml_slides)
    try:
        slides[index]
        xml_slides.remove(slides[index])
    except ValueError:
        print("error") 

def is_extra_slide(presentation, slide_index, remove_tag):
    extra_slide_exists = False
    extra = 'EXTRA_SLIDE'
    for new_slide_shape in presentation.slides[slide_index].shapes:
        if new_slide_shape.has_text_frame:
            matches = check_tag_exist(extra, new_slide_shape)
            if(matches):
                extra_slide_exists = True
                if(remove_tag):
                    replace_tags(str(f"+++ {extra} +++"), "", new_slide_shape)
                break
    return extra_slide_exists

def remove_extra_slides(presentation):
    extra = 'EXTRA_SLIDE'
    slides = [slide for slide in presentation.slides]
    slide_indexs_to_delete = []
    for slide in slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                matches = check_tag_exist(extra, shape)
                print(matches)
                if(matches):
                    slide_indexs_to_delete.append(slides.index(slide))

    if(len(slide_indexs_to_delete) > 0):
        array_index = 0
        for s_index in slide_indexs_to_delete:
            slide_index = slides.index(slide)
            delete_slides(presentation, slide_index - array_index)
            array_index += 1 