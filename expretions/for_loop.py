from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
import math
import re
import pydash

from helpers.utils import check_tag_exist, replace_tags, get_tag_content, get_tag_from_string
from helpers.texts import text_tag_update
from helpers.images import replace_images
from helpers.tables import replace_tables


def looper(presentation, slide, shape,slides_index, replacements):
    pattern = r'\+\+\+FOR (.*?) FOR-END\+\+\+'
    matches = get_tag_content(pattern, shape)
    
    if( not matches or len(matches) < 1):
        return

    for match in matches:
        pattern_condition = r'\(\((.*?)\)\)'
        matched_condition = get_tag_from_string(pattern_condition,match)

        pattern_content = r'\<\<(.*?)\>\>'
        matched_content = get_tag_from_string(pattern_content,match)
        
        for contidion in matched_condition:
            object_value = pydash.get(replacements, contidion)
            text_result = ""
            if(object_value):
                for data in object_value:
                    updated_data = text_tag_update(matched_content[0],data)
                    if(updated_data and updated_data["text"]):
                        text_result += updated_data["text"] + "\n"


        # this is not working if you use tabspaces, but you can use spaces
        replace_tags(str(f"+++FOR {match} FOR-END+++"), text_result, shape)

