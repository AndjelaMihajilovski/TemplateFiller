import csv
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
        doc.save('result/' + row[headers[0]] + '.docx')        


