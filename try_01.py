import csv
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties
from odf.text import H, P, Span


with open ('C:\\Users\\Czerwiec\\Desktop\\VS Code work\\csv\\tomasz.czerwinski.csv', mode = 'r', encoding='utf-8') as file:
    csv_file = csv.reader(file)
    print()
    list_00 = []
    for line in csv_file:
        # if lines[3] == "dokumentacja":
        #     print(lines[0][3:] + "\t" + lines[1])

        if line[3] == "wizualizacja":
            # print(line[0][3:] + "\t" + line[1])
            list_00.append(line[0][3:] + " " + line[1])
            # print(list_00)
            doc = OpenDocumentText()
            f = H(outlinelevel=1, text = "Wizualizacja")
            doc.text.addElement(f)
            # y = line[0][3:] + " " + line[1]
            # print(lines[0][3:] + lines[1])
            # print(y)
    for x in list_00:
        print(x)
        i = P(text = x)
        
        
        doc.text.addElement(i)
        # doc.text.addElement(o)
        
      
        doc.save("this is number 2.odt")

            # x = P(text = y)
            
            # # list_01.append(y)

            # # print(list_01)

            
            # # y = P(text = list_01[1])
            # # # document = Document()
            # # # document.add_heading('Konwerter')
            # # # document.add_paragraph(x)
            # # # document.save("doc_01.docx")

            # doc.text.addElement(x)
            # doc.save("this is a try.odt")
   

            
            
    




