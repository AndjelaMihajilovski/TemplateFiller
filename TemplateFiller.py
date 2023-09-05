import csv
import os
import tkinter.filedialog as fd
import tkinter as prog
from docx import Document

root= prog.Tk()
root.title("TemplateFiller")
canvas1 = prog.Canvas(root, width = 500, height = 500)
canvas1.pack()
filename = ""
csvPath = ""
path = ""
filenameNotification = prog.Label(root, text= '', fg='red', font=('helvetica', 12, 'bold'))
csvNotification = prog.Label(root, text= '', fg='red', font=('helvetica', 12, 'bold'))
pathNotification = prog.Label(root, text= '', fg='red', font=('helvetica', 12, 'bold'))


def BrowseFile (): 
    global filename
    global filenameNotification
    filenameNotification.config(text='') 
    currdir = os.getcwd()
    filetypes = (("Word files", "*.docx"), ("All files", "*.*"))
    tempdir = fd.askopenfilename(parent=root, initialdir=currdir, title='Please select a Word Template File', filetypes=filetypes)
    
    if len(tempdir) > 0:
        print("You chose %s" % tempdir)
        filename = tempdir
        print ("\template file name = ", filename)
        filenameLabel = prog.Label(root, text= filename, fg='green', font=('helvetica', 10, 'bold'))
        canvas1.create_window(250, 120, window=filenameLabel)

def BrowseCSV (): 
    global csvPath 
    currdir = os.getcwd()
    global csvNotification
    csvNotification.config(text='')
    filetypes = (("CSV files", "*.csv"), ("All files", "*.*"))
    tempdir = fd.askopenfilename(parent=root, initialdir=currdir, title='Please select a CSV file', filetypes=filetypes)
    
    if len(tempdir) > 0:
        print("You chose %s" % tempdir)
        csvPath = tempdir
        print ("\template file name = ", csvPath)
        csvPathLabel = prog.Label(root, text= csvPath, fg='green', font=('helvetica', 10, 'bold'))
        canvas1.create_window(250, 170, window=csvPathLabel)

def BrowseDestinationPath (): 
    global path 
    global pathNotification
    pathNotification.config(text='')
    currdir = os.getcwd()
    tempdir = fd.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')
    
    if len(tempdir) > 0:
        print("You chose %s" % tempdir)
        path = tempdir
        print ("\template file name = ", path)
        pathLabel = prog.Label(root, text= path, fg='green', font=('helvetica', 10, 'bold'))
        canvas1.create_window(250, 220, window=pathLabel)

def CreateTemplate ():
    global filename
    global csvPath 
    global path 
    global filenameNotification
    global csvNotification
    global pathNotification

    if  filename == "":
        filenameNotification = prog.Label(root, text= 'You didn`t choose Word Template file!', fg='red', font=('helvetica', 12, 'bold'))
        canvas1.create_window(250, 270, window=filenameNotification)

    if  path == "":
        pathNotification = prog.Label(root, text= 'You didn`t choose save directory!', fg='red', font=('helvetica', 12, 'bold'))
        canvas1.create_window(250, 310, window=pathNotification)  

    if csvPath != "": 
        with open(csvPath, newline='') as csvfile:
            csv_reader = csv.DictReader(csvfile)
            headers = csv_reader.fieldnames
            result_dir = os.path.join(path, "result")
            if not os.path.exists(result_dir):
                os.makedirs(result_dir)
                print("The 'result' directory is created!")

            for row in csv_reader:
                doc = Document(filename)
                print('Parsing: ' + row['Name'])

                for columnName in headers:
                    for paragraph in doc.paragraphs:
                        paragraph.text = paragraph.text.replace('<' + columnName + '>', row[columnName])

                doc.save(os.path.join(result_dir, row[headers[0]] + '.docx'))     

                label1 = prog.Label(root, text= 'Finished!', fg='green', font=('helvetica', 12, 'bold'))
                canvas1.create_window(250, 270, window=label1)
    else:
        csvNotification = prog.Label(root, text= 'You didn`t choose CSV file!', fg='red', font=('helvetica', 12, 'bold'))
        canvas1.create_window(250, 290, window=csvNotification)  
    
button = prog.Button(text='OpenTemplate', command=BrowseFile, bg='grey',fg='white')
canvas1.create_window(250, 100, window=button)
button1 = prog.Button(text='OpenCSV', command=BrowseCSV, bg='grey',fg='white')
canvas1.create_window(250, 150, window=button1)
button2 = prog.Button(text='Destination', command=BrowseDestinationPath, bg='grey',fg='white')
canvas1.create_window(250, 200, window=button2)
button3 = prog.Button(text='Create', command=CreateTemplate, bg='grey',fg='white')
canvas1.create_window(250, 250, window=button3)

root.mainloop()



