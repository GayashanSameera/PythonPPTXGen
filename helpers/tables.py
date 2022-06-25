from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
import math
import re
import pydash

from helpers.utils import check_tag_exist, replace_tags, get_tag_content

def replace_tables(presentation, slide, shape, slide_index, replacements):
    pattern = r'\+\+\+TABLE_ADD (.*?) \+\+\+'
    matches = get_tag_content(pattern, shape)

    if( not matches or len(matches) < 1):
        return

    for match in matches:
        object_value = pydash.get(replacements, match)
        if(object_value):
            replace_tags(str(f"+++TABLE_ADD {match} +++"), "", shape)
            create_table(presentation, slide, shape, slide_index, object_value)

def create_table(presentation, slide, shape, slide_index, replacements):
    row_count = replacements["row_count"]
    cols = replacements["colum_count"]
    headers = replacements["headers"]
    row_data = replacements["rows"]
    styles = replacements["styles"]
    table_count_per_slide = replacements["table_count_per_slide"]

    total_rows = len(row_data)
    total_table_count = math.ceil(total_rows / row_count)

    slide_count = math.ceil(total_table_count / table_count_per_slide) 
    extra_slide_count = slide_count - 1

    s = 0
    end = 0
    current_slide = 0
    while s < slide_count:
        j = 0
        left = 1

        if total_table_count > table_count_per_slide :
            table_count = table_count_per_slide
        else:
            table_count = total_table_count

        slide_row_start = 0
        slide_row_end = 0
        extra_slide_exists = False

        if(s > 0 and (not presentation.slides[slide_index + s])):
            break

        if(s > 0 and current_slide != s):
            for new_slide_shape in presentation.slides[slide_index + s].shapes:
                if new_slide_shape.has_text_frame:
                    extra = 'EXTRA_SLIDE'
                    matches = check_tag_exist(extra, new_slide_shape)
                    if(matches):
                        extra_slide_exists = True
                        replace_tags(str(f"+++ {extra} +++"), "", new_slide_shape)
                        break
            current_slide = s

        if(s > 0 and (not extra_slide_exists)):
            return 
            

        while j < table_count:
            shape = presentation.slides[slide_index + s].shapes.add_table(row_count + 1, cols, Inches(left) , Inches(styles["top"]), Inches(styles["width"]), Inches(styles["row_height"]))
            table = shape.table
            tables_headers(table, headers)

            slide_row_start = (row_count * j ) + 1 + end
            slide_row_end = slide_row_start + (row_count - 1)

            tables_rows(table,row_data, slide_row_start, slide_row_end ,total_rows )
            left += 2
            j += 1

        end = slide_row_end
        total_table_count -= table_count_per_slide
        s += 1

def tables_headers(table, headers):
    i = 0
    for header in headers:
        table.cell(0, i).text = header
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(200, 253, 251)
        _set_cell_border(cell,"949595", '12000')
        para = cell.text_frame.paragraphs[0]
        para.font.bold = True
        para.font.size = Pt(5)
        para.font.name = 'Comic Sans MS'
        para.font.color.rgb = RGBColor(19, 170, 246)
        table.columns[i].width = Inches(0.6)
        i += 1

def tables_rows(table, rowData, start, end,totalRows):
    j = start
    cell_start = 1

    entCount = end
    if end > totalRows:
        entCount = totalRows

    while j < entCount + 1:
        element = rowData[j - 1]
        k = 0
        if element:
            for key, value in element.items():
                table.cell(cell_start, k).text = value
                cell = table.cell(cell_start, k)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(228, 228, 228)
                _set_cell_border(cell,"949595", '12000')
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(6)
                para.font.name = 'Comic Sans MS'
                k += 1
        j += 1
        cell_start += 1

def SubElement(parent, tagname, **kwargs):
        element = OxmlElement(tagname)
        element.attrib.update(kwargs)
        parent.append(element)
        return element

def _set_cell_border(cell, border_color="000000", border_width='12700'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for lines in ['a:lnL','a:lnR','a:lnT','a:lnB']:
        ln = SubElement(tcPr, lines, w=border_width, cap='flat', cmpd='sng', algn='ctr')
        solidFill = SubElement(ln, 'a:solidFill')
        srgbClr = SubElement(solidFill, 'a:srgbClr', val=border_color)
        prstDash = SubElement(ln, 'a:prstDash', val='solid')
        round_ = SubElement(ln, 'a:round')
        headEnd = SubElement(ln, 'a:headEnd', type='none', w='med', len='med')
        tailEnd = SubElement(ln, 'a:tailEnd', type='none', w='med', len='med')
