from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn
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
                        parsedLine = []
                        normalText = ""
                        dotCommandActive = line.startswith('.')
                        if dotCommandActive:
                            skipCount = 2
                            dotCommand = line[1]
                            if dotCommand.lower() == 'p': #pagebreak
                                pass
                            elif dotCommand.lower() == 'a': #add lines
                                skipCount+=1
                                pass
                            elif dotCommand.lower() == 'h': #header
                                pass
                            elif dotCommand.lower() == 'f': #footer
                                pass
                            else:
                                return #Unknown
                            line = line[skipCount:]
                        lastRun = {'text':'\u202B', 'bold':False,'underline':False, 'centered':False, 'left':False}
                        sequenceLen = 0
                        for c in line:
                            if c == '╩': #Underline
                                lastRun['underline'] = True
                            elif c == '╦': #Underline Bold
                                lastRun['underline'] = True
                                lastRun['bold'] = True
                            elif c == '╔': #Bold
                                lastRun['bold'] = True
                            elif c == '█': #Printer Command
                                pass
                            elif c == '▄': #Printer Command 2
                                pass
                            elif c == '┘': #Center Align
                                lastRun['centered'] = True
                            elif c == '┌': #Left Aling
                                lastRun['left'] = True
                            elif c == '\x18': #UpperMark
                                pass
                            elif c == '\x19': #LowerMark
                                pass
                            else:
                                if False: #c=='.' or c==',': #TODO: seperate dots and commas
                                    if lastRun['text']!='':
                                        parsedLine.append(lastRun)
                                    lastRun = {'text':'\u202B'+c, 'bold':False,'underline':False, 'centered':False, 'left':False}
                                    parsedLine.append(lastRun)
                                    lastRun = {'text':'\u202B', 'bold':False,'underline':False, 'centered':False, 'left':False}
                                    sequenceLen = 0
                                elif False: #c==' ':
                                    if sequenceLen:
                                        parsedLine.append(lastRun)
                                        lastRun = {'text':'\u202B', 'bold':False,'underline':False, 'centered':False, 'left':False}
                                        sequenceLen = 0
                                    lastRun['text'] += ' '
                                else:
                                    sequenceLen += 1
                                    lastRun['text'] += c
                            
                        if lastRun['text']!='':
                            parsedLine.append(lastRun)    
                        parsedList.append(parsedLine)
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
    normalStype = document.styles['Normal']
    normalStype.font.name = 'Courier New'
    normalStype.font.size = Pt(10)
    #document.add_heading('Document Title', 0)
    for e in parsedList:
        p = document.add_paragraph('')
        p.paragraph_format.space_before = 0
        p.paragraph_format.space_after = 0
        p.paragraph_format.line_spacing = 1
        ppr = p._element.get_or_add_pPr()

        w_nsmap = '{'+ppr.nsmap['w']+'}'
        bidi = None
        jc = None
        cs = None
        rtl = None
        for element in ppr:
            if element.tag == w_nsmap + 'bidi':
                bidi = element
            if element.tag == w_nsmap + 'jc':
                jc = element
            if element.tag == w_nsmap + 'cs':
                cs = element
            if element.tag == w_nsmap + 'rtl':
                rtl = element
        if bidi is None:
            bidi = OxmlElement('w:bidi')
        if jc is None:
            jc = OxmlElement('w:jc')
        if cs is None:
            cs = OxmlElement('w:cs')
        if rtl is None:
            rtl = OxmlElement('w:rtl')
        bidi.set(qn('w:val'),'1')
        jc.set(qn('w:val'),'both')
        cs.set(qn('w:cs'),'1')
        rtl.set(qn('w:rtl'),'1')
        ppr.append(bidi)
        ppr.append(jc)
        ppr.append(rtl)
        #p.style.font.rtl = True
        for run in e:
            r = p.add_run(run['text'])
            r.bold = run['bold']
            r.cs_bold = run['bold']
            r.underline = run['underline']
            if run['centered']:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if run['left']:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r.font.name = 'Courier New'
            r.font.size = Pt(10)
            #r.font.complex_script = True

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