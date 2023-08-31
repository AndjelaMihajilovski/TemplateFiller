import csv
import os
from docx import Document


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


