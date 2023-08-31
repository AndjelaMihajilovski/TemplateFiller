import csv
import os
import tkinter as prog
from docx import Document

root= prog.Tk()

canvas1 = prog.Canvas(root, width = 300, height = 300)
canvas1.pack()

def CreateTemplate (): 

    filename = "CocktailInfoTemplate.docx"
    with open('Cocktails.csv', newline='') as csvfile:
        csv_reader = csv.DictReader(csvfile)
        headers = csv_reader.fieldnames
        for row in csv_reader:
            doc = Document(filename)
            print('Parsing: ' + row['Name'])

            for columnName in headers:
                for paragraph in doc.paragraphs:
                    paragraph.text = paragraph.text.replace('<' + columnName + '>', row[columnName])
            #pcheck if a directory exists
            path = "result"
            # Check whether the specified path exists or not
            isExist = os.path.exists(path)
            if not isExist:
                os.makedirs(path)
                print("The new directory is created!")
            doc.save('result/' + row[headers[0]] + '.docx')        

        label1 = prog.Label(root, text= 'Finished!', fg='blue', font=('helvetica', 12, 'bold'))
        canvas1.create_window(150, 200, window=label1)
    
button1 = prog.Button(text='Start', command=CreateTemplate, bg='brown',fg='white')
canvas1.create_window(150, 150, window=button1)

root.mainloop()



