import csv
import os
import tkinter.filedialog as fd
import tkinter as prog
from docx import Document

root= prog.Tk()
root.title("TemplateFiller")
canvas1 = prog.Canvas(root, width = 300, height = 300)
canvas1.pack()
filename = ""
csvPath = ""

def BrowseFile (): 
    global filename 
    currdir = os.getcwd()
    filetypes = (("Word files", "*.docx"), ("All files", "*.*"))
    tempdir = fd.askopenfilename(parent=root, initialdir=currdir, title='Please select a Word Template File', filetypes=filetypes)
    
    if len(tempdir) > 0:
        print("You chose %s" % tempdir)
        filename = tempdir
        print ("\template file name = ", filename)

def BrowseCSV (): 
    global csvPath 
    currdir = os.getcwd()
    filetypes = (("CSV files", "*.csv"), ("All files", "*.*"))
    tempdir = fd.askopenfilename(parent=root, initialdir=currdir, title='Please select a CSV file', filetypes=filetypes)
    
    if len(tempdir) > 0:
        print("You chose %s" % tempdir)
        csvPath = tempdir
        print ("\template file name = ", csvPath)

def BrowseDestinationPath (): 
    global csvPath 
    currdir = os.getcwd()
    tempdir = fd.askopenfilename(parent=root, initialdir=currdir, title='Please select a directory')
    
    if len(tempdir) > 0:
        print("You chose %s" % tempdir)
        csvPath = tempdir
        print ("\template file name = ", csvPath)

def CreateTemplate (): 
    global csvPath 

    with open(csvPath, newline='') as csvfile:
        csv_reader = csv.DictReader(csvfile)
        headers = csv_reader.fieldnames
        for row in csv_reader:
            doc = Document(filename)
            print('Parsing: ' + row['Ime'])

            for columnName in headers:
                for paragraph in doc.paragraphs:
                    paragraph.text = paragraph.text.replace('<' + columnName + '>', row[columnName])
            #pcheck if a directory exists
            path =os.path.abspath("Documents/TemplateFillerStart/result")
            # Check whether the specified path exists or not
            isExist = os.path.exists(path)
            if not isExist:
                os.makedirs(path)
                print("The new directory is created!")
            doc.save('C:/Users/FHG02/Documents/TemplateFillerStart/result/' + row[headers[0]] + '.docx')        

        label1 = prog.Label(root, text= 'Finished!', fg='green', font=('helvetica', 12, 'bold'))
        canvas1.create_window(150, 250, window=label1)
    
button = prog.Button(text='OpenTemplate', command=BrowseFile, bg='grey',fg='white')
canvas1.create_window(150, 100, window=button)
button1 = prog.Button(text='OpenCSV', command=BrowseCSV, bg='grey',fg='white')
canvas1.create_window(150, 150, window=button1)
button2 = prog.Button(text='Create', command=CreateTemplate, bg='grey',fg='white')
canvas1.create_window(150, 200, window=button2)

root.mainloop()



