from pptx import Presentation
from pptx.util import Inches
# from pptx.util import Px
# from PIL import Image

img_path = 'sample.png'

def replace_txt(replacements, shapes):
    for shape in shapes:
        for match, replacement in replacements.items():
            #find tags inside text frame
            if shape.has_text_frame:
                if (shape.text.find(match)) != -1:
                    text_frame = shape.text_frame
                    
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            cur_text = run.text
                            new_text = cur_text.replace(str(match), str(replacement))
                            run.text = new_text
            #find tags inside table frame
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if match in cell.text:
                            new_text = cell.text.replace(match, replacement)
                            cell.text = new_text


def replace_images(replacements, slide, index):
    if str(index) in replacements:
        objectKey = str(index)
        replacedData = replacements[objectKey]
        print(replacedData)

        for data in replacedData:
            print(data)
            slide.shapes.add_picture(data["path"], data["size"]["left"], data["size"]["top"], data["size"]["width"] ,data["size"]["height"] )

if __name__ == '__main__':

    prs = Presentation('indput1.pptx')
    dataBump = {
        "image_replaces": {
                        "1": [{"path": "1.png", "size":{"left":Inches(2),"top":Inches(2), "height":Inches(3), "width":Inches(8) }}]
                    },
        "text_replaces":  {
            '{{var1}}': 'LGIM Bridge',
            '{{var2}}': 'External asset workflow',
            '{{var3}}': 'Active Client page – External assets',
            '{{var4}}': 'LGIM/Client user',
            '{{var5}}': 'Active Client page – External assets',
            '{{var6}}': 'LGIM/Client user – when the user clicks on ‘Edit detail’ for a date',
            '{{var7}}': 'Active Client page – External assets',
            '{{var8}}': 'Information button text',
            '{{var10}}': "If the scheme holds assets which are not managed by LGIM, details can be entered here. This information will be used for tracking the funding level, and also for overall portfolio risk analysis and funding level projections.Enter the value of non-LGIM assets held at different dates. We will roll the asset value forward in an approximate manner between dates provided.By default any asset values provided will be treated as cash for both funding level tracking and risk analysis purposes. However, if you click on ‘Edit details’ for a given asset value then you will be able to provide a breakdown of the value by fund and asset class. If this information is provided then we will roll forward the asset value in line with index returns for the relevant asset class, and will include the external assets within overall portfolio risk analysis. Once you have entered the breakdown of the asset value at one date, this breakdown will be carried forward to subsequent dates by default, unless you choose to edit the detail within these later dates (for example, following a significant change in the portfolio). This enables you to quickly update the value of non-LGIM assets without having to re-enter the underlying detail at every date.",
            '{{var11}}': 'PV01 and IE01 data should be available from your LDI provider. Note that positive numbers should typically be entered.',
            '{{var12}}': 'PV01 - Please enter the increase in asset value that would expected in £ terms if interest rates fell 0.01%.',
            '{{var13}}':'IE01 - Please enter the increase in asset value that would expected in £ terms if price inflation expectations increased 0.01%.',
            '{{var14}}':'If you do not know this information then you can simply leave these boxes blank, and we will treat the LDI assets as cash.'
            },
        "table_replaces": {

        } 
             
    }


    slides = [slide for slide in prs.slides]
    shapes = []
    for slide in slides:
        print(slide)
        if "image_replaces" in dataBump:
            replace_images(dataBump["image_replaces"], slide, slides.index(slide))

        if "text_replaces" in dataBump:
            for shape in slide.shapes:
                shapes.append(shape)

    if "text_replaces" in dataBump:
        replace_txt(dataBump["text_replaces"], shapes)

    prs.save('output.pptx')