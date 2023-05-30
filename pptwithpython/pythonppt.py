import collections
import collections.abc
from pptx import Presentation
from pptx.util import Inches
from pptx import Presentation
from pptx.util import Pt
def add_text_to_slide(slide, text):
    # Create a text box shape on the slide
    left = top = width = height = Inches(1) # Adjust these values as per your requirement
    text_box = slide.shapes.add_textbox(left, top, width, height)
    text_frame = text_box.text_frame

    # Add the text to the text frame
    p = text_frame.add_paragraph()
    p.text = text

# Open the presentation file
# ppt_file = 'newppt.pptx'
presentation = Presentation()


# Open the text file and read the content
txt_file = 'sample_slide1_input.txt'
txt_file2 = 'sample_slide2_input.txt'

with open(txt_file, 'r') as file:
    content = file.read()
with open(txt_file2, 'r') as file:
    content1 = file.read()


# Create a new slide
slide = presentation.slides.add_slide(presentation.slide_layouts[6])
slide1 = presentation.slides.add_slide(presentation.slide_layouts[6])
# Add the text content to the slide

# Save the modified presentation
modified_ppt_file = 'modifiednew15.pptx'
presentation.save(modified_ppt_file)

presentation = Presentation('modifiednew15.pptx')
for slide in presentation.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Garamond'  # Replace 'New Font Name' with the desired font name
                    run.font.size = Pt(16)
modified_ppt_file = 'modifiednew15.pptx'
presentation.save(modified_ppt_file)






