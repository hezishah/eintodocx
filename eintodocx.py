from docx import Document
from docx.shared import Inches

import glob
import sys
import os

PAR_STATE_ERROR = 0
PAR_STATE_NORMAL = 1
PAR_STATE_UNDERLINE = 2
PAR_STATE_BOLD = 3

def parseEin(file):
    try:
        with open(file,mode='rb') as f:
            first = f.read(1)
            parsedList = []
            if first == b';':
                try:
                    fileText = f.read().decode('cp856')
                    lines = fileText.splitlines()
                    header = [lines[0]]
                    lineIndex = 1
                    for line in lines[lineIndex:]:
                        lineIndex += 1
                        if line.startswith(';'):
                            header.append(line)
                        else:
                            break
                    for line in lines[lineIndex:]:
                        normalText = ""
                        for c in line:
                            if c == '╩': #Underline - Bold
                                pass
                            elif c == '╦': #Underline
                                pass
                            elif c == '╔': #Bold
                                pass
                            elif c == '█': #Line Break
                                pass
                            elif c == '▄': #Half Line break
                                pass
                            else:
                                normalText += c
                        if normalText!='':
                            parsedList.append({"normal":normalText})
                    saveDocx(parsedList, file)
                    return
                except Exception as e:
                    print(e)
                    return
    except:
        pass
    return

def saveDocx(parsedList, file):
    document = Document()

    #document.add_heading('Document Title', 0)
    p = document.add_paragraph('')
    for e in parsedList:
        for k in e:
            if k == 'normal':
                p.add_run(e[k])
    '''
    p = document.add_paragraph('A plain paragraph having some ')
    p.add_run('bold').bold = True
    p.add_run(' and some ')
    p.add_run('italic.').italic = True

    document.add_heading('Heading, level 1', level=1)
    document.add_paragraph('Intense quote', style='Intense Quote')

    document.add_paragraph(
        'first item in unordered list', style='List Bullet'
    )
    document.add_paragraph(
        'first item in ordered list', style='List Number'
    )

    #document.add_picture('monty-truth.png', width=Inches(1.25))

    records = (
        (3, '101', 'Spam'),
        (7, '422', 'Eggs'),
        (4, '631', 'Spam, spam, eggs, and spam')
    )

    table = document.add_table(rows=1, cols=3)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Qty'
    hdr_cells[1].text = 'Id'
    hdr_cells[2].text = 'Desc'
    for qty, id, desc in records:
        row_cells = table.add_row().cells
        row_cells[0].text = str(qty)
        row_cells[1].text = id
        row_cells[2].text = desc

    document.add_page_break()

    '''
    document.save('out/demo-'+ os.path.split(file)[-1] +'.docx')

if __name__ == "__main__":
    if len(sys.argv) < 2:
        exit(0)
    for arg in sys.argv[1:]:
        for file in glob.glob(arg):
            parseEin(file)
    exit(0)