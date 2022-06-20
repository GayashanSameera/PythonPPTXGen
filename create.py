from pptx import Presentation

def replace_tags(replacements, shapes):
    for shape in shapes:
        for match, replacement in replacements.items():
            if shape.has_text_frame:
                if (shape.text.find(match)) != -1:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            cur_text = run.text
                            new_text = cur_text.replace(str(match), str(replacement))
                            run.text = new_text
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if match in cell.text:
                            new_text = cell.text.replace(match, replacement)
                            cell.text = new_text

if __name__ == '__main__':

    prs = Presentation('input.pptx')
    slides = [slide for slide in prs.slides]
    shapes = []
    for slide in slides:
        for shape in slide.shapes:
            shapes.append(shape)

    replaces = {
            '{{var1}}': 'text 1',
            '{{var2}}': 'text 2',
            '{{var3}}': 'text 3'
            }
    replace_tags(replaces, shapes)
    prs.save('output.pptx')