import docx
from docx import Document
from docx.shared import RGBColor
from tkinter import *
from tkinter import filedialog


def readtxt(filename, color: tuple[int, int, int]):
    doc = docx.Document(filename)

    fullText = []
    for para in doc.paragraphs:

        # Getting the colored words from the doc
        if (getcoloredTxt(para.runs, color)):

            # Concatenating list of runs between the colored text to single a string
            sentence = "".join(r.text for r in para.runs)
            fullText.append(sentence)

    return fullText

def getcoloredTxt(runs, color):

    coloredWords, word = [], ""
    for run in runs:
        if run.font.color.rgb == RGBColor(*color):
            word += str(run.text)

        elif word != "":
            coloredWords.append(word)
            word = ""
    if word != "":
        coloredWords.append(word + "\n")

    return coloredWords

def openFile():
    global filepath
    global filepath2
    filepath = filedialog.askopenfilename(initialdir="/",
                                          title="",
                                          filetypes= (("word documents","*.docx"),
                                                      ("all files","*.*")))
    file = open(filepath,'r')
    #print(filepath)
    file.close()
    filepath2 = str(filepath)
    #filepath2 = '"' + filepath + '"'
    print(filepath2)

    return filepath2

def generateReport():
    fullText = readtxt(filename=filepath2,
                       color=(255, 0, 0))
    s = ''.join(fullText)
    w = (s.replace (']', ']\n\n'))
    w = (w.replace ('\n[', '['))
    print('\n' + w)
    paragraph = report.add_paragraph()
    runner = paragraph.add_run("\n" + filepath2)
    runner.bold = True #makes the header bold
    paragraph = report.add_paragraph(w)
    report.save('report1.docx')



if __name__ == '__main__':
    report = Document()
    window = Tk(className='TARGEST')
# set window size
    window.geometry("150x100")
    button = Button(text="Choose Document",command=openFile)
    button.pack()
    #Button(window, text="Generate Report ", command=window.destroy).pack()
    #window.mainloop()

    Button(window, text="Generate Report ", command=generateReport).pack()

    button = Button(text="End Program",command=window.destroy)
    button.pack()

    window.mainloop()

    #filepath2 = '"' + filepath + '"'
    #print(filepath2)
    #fullText = readtxt("testred.docx")
    #print(fullText)
    #filepath3 = filepath2
    #print(filepath3)

    #print(fullText)
    #lister2 = [fullText]
    #d = ''.join(fullText)
    #print(fullText)
    #print(d)
    #words = d.split('')
    #words2 = d.split('')

    #print(words)




