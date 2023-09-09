from pptx import Presentation
import sys
from googletrans import Translator

prs = Presentation(sys.argv[3])
translator = Translator(service_urls=['translate.googleapis.com'])

dest = sys.argv[1]
i = 0

mode = sys.argv[2]
if not (mode == 'merge' or mode == 'overwrite'):
    print('invalid mode please use merge or overwrite')
    quit()

for slide in prs.slides:
    i = i + 1
    shapes = []
    
    # only read in text-boxes
    for shape in slide.shapes:
        if(shape.has_text_frame):
            if(shape.text != ''):
                shapes.append(shape)
    
    # Translate notes
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame
        if mode == 'merge':
            notes.text = notes.text  + '\n' + translator.translate(notes.text, dest=dest).text
        elif mode == 'overwrite': 
            notes.text = translator.translate(notes.text, dest=dest).text
    
    # translate title
    title = shapes[0].text
    if mode == 'merge':
            shapes[0].text = title + '\n\n' + translator.translate(title, dest=dest).text
    elif mode == 'overwrite': 
            shapes[0].text = translator.translate(title, dest=dest).text


    # translate content
    for shape in shapes[1:]:
            if(shape.text != ''): 
                content = shape.text
                translation = translator.translate(content, dest=dest).text

            if mode == 'merge':
                shape.text = content + '\n' + translation
            elif mode == 'overwrite': 
                shape.text = translation


    print('a slide was translated', i, '/', len(prs.slides))

prs.save(sys.argv[4])
print('done')