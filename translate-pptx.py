from pptx import Presentation
import sys
from googletrans import Translator

prs = Presentation(sys.argv[3])
translator = Translator(service_urls=['translate.googleapis.com'])

dest = sys.argv[1]
i = 0

for slide in prs.slides:
    i = i + 1
    shapes = []
    
    # only read in text-boxes
    for shape in slide.shapes:
        if(shape.has_text_frame):
            if(shape.text != ''):
                shapes.append(shape)
    
    # translate title
    title = shapes[0].text
    if sys.argv[2] == 'merge':
            shapes[0].text = title + '\n\n' + translator.translate(title, dest=dest).text
    elif sys.argv[2] == 'overwrite': 
            shapes[0].text = translator.translate(title, dest=dest).text
    else:
            print('invalid command', sys.argv[1])

    # translate content
    for shape in shapes[1:]:
            if(shape.text != ''): 
                content = shape.text
                translation = translator.translate(content, dest=dest).text

            if sys.argv[2] == 'merge':
                shape.text = content + '\n' + translation
            elif sys.argv[2] == 'overwrite': 
                shape.text = translation
            else:
                print('invalid command', sys.argv[2])


    print('a slide was translated', i, '/', len(prs.slides))

prs.save(sys.argv[4])
print('done')