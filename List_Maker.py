import os
from os import walk
from win32api import *
from datetime import datetime
import time
import csv
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties, ParagraphProperties
from odf.text import H, P, Span


all_programs = ['CAD', 'CAD Licencja', 'Dystrybutor baz', 'Instalator', 'Konwerter mdb', 'Marketing', 'Panel aktualizacji', 'Podesty', 'Tłumaczenia', 'aktualizator internetowy', 'analytics', 'asystent pobierania', 'bazy danych', 'deinstalator', 'dokumentacja', 'dot4cad', 'drzwi i okna', 'edytor bazy płytek', 'edytor szafek bazy kuchennej', 'edytor szafek użytkownika', 'edytor ścian', 'elementy dowolne', 'export3D', 'instalator baz danych', 'instalator programu', 'konwerter', 'kreator ścian', 'launcher', 'listwy', 'manager', 'obrót 3d / przesuń', 'obserVeR', 'przeglądarka PDF', 'render', 'szafy wnękowe', 'ukrywacz', 'wersja Leroy Merlin', 'wersja Obi', 'wizualizacja', 'wstawianie elementów wnętrzarskich', 'wstawianie elementów wnętrzarskich w wizualizacji', 'zestawienie płytek, farb, fug, klejów']

kitchen_pro = ['blaty','wstawianie elementów agd', 'wstawianie szafek kuchennych', 'wycena']


# laptop:
source_path = "C:\\Users\\Czerwiec\\Desktop\\setup ini test\\12.06.2024 hotfix"
csv_path = "C:\\Users\\Czerwiec\\Desktop\\VS code workplaces\\csv\\csv\\tomasz.czerwinski.csv"

# source_path = "C:\\Users\\Czerwiec\\Desktop\\setup ini test\\12.06.2024 hotfix"
# csv_path = "C:\\Users\\Czerwiec\\Desktop\\VS code workplaces\\csv\\csv\\tomasz.czerwinski.csv"

bug_dict = {}

def get_version_number(file_path):
    try:
        File_information = GetFileVersionInfo(file_path, "\\")
        ms_file_version = File_information["FileVersionMS"]
        ls_file_version = File_information["FileVersionLS"]

        list_of_versions = [
            str(HIWORD(ms_file_version)),
            str(LOWORD(ms_file_version)),
            str(HIWORD(ls_file_version)),
            str(LOWORD(ls_file_version)),
        ]
        return ".".join(list_of_versions)
    except:
        return "-"


def get_creation_date(path, long=True):
    time_created = time.ctime(os.path.getmtime(path))
    t_obj = time.strptime(time_created)
    if long == True:
        return time.strftime("%Y-%m-%d %H:%M:%S", t_obj)
    elif long == False:
        return time.strftime("%d.%m.%Y", t_obj)


def make_path_list(folder_path):
    pathName_dictionary = {}
    for dirpath, dirnames, filenames in walk(folder_path):
        for file_name in filenames:
            path = os.path.abspath(os.path.join(dirpath, file_name))
            pathName_dictionary[path] = file_name
    return pathName_dictionary


def sort_files_del_from_dict(filesList, dict, indexes):

    list_of_creation_time = []
    for x in filesList:
        list_of_creation_time.append(get_creation_date(x))
    
    sortedList = list_of_creation_time.index(max(list_of_creation_time))
    list_of_indexes_to_delete = indexes[sortedList + 1 :]

    for index in sorted(list_of_indexes_to_delete, reverse=True):
        for i, name in enumerate(list(dict.keys())):
            if index == i:
                del dict[name]
            
    for i in list(dict.keys()):
        if i.endswith("txt") == True:
            del dict[i]
    return dict


def make_list_to_cut(dict):
    indexesToCut = []
    for i, v in enumerate(dict.values()):
        if list(dict.values()).count(v) > 1 and i not in indexesToCut:
            indexesToCut.append(i)
    return indexesToCut


def list_paths(i_list, paths):
    cuted_pathList = []
    for index in i_list:
        pathList = list(paths.keys())
        cuted_pathList += [pathList[index]]
    return cuted_pathList


def add_lines_to_lists(data_file, all_list, rest_list):
    cat_list_with_duplicates = []
    cat_list_with_duplicates02 = []
    cat_list_LM_with_duplicates = []
    cat_list_OBI_with_duplicates = []
    cat_list_OBI = []
    cat_list_LM = []
    cat_list02 = []
    cat_list = []

    for line in data_file:    
        
        if line[3].lower() == "wersja obi":
            cat_list_OBI_with_duplicates.append(line[3])
            cat_list_OBI = []
            [cat_list_OBI.append(x) for x in cat_list_OBI_with_duplicates if x not in cat_list_OBI]
            
        elif line[3].lower() == "wersja leroy merlin":
            cat_list_LM_with_duplicates.append(line[3])
            cat_list_LM = []
            [cat_list_LM.append(x) for x in cat_list_LM_with_duplicates if x not in cat_list_LM]

        elif line[3].lower() != "projekt" and line[3] in all_list:
            cat_list_with_duplicates.append(line[3])
            cat_list = []
            [cat_list.append(x) for x in cat_list_with_duplicates if x not in cat_list]    

        elif line[3] in rest_list:
            cat_list_with_duplicates02.append(line[3])
            cat_list02 = []
            [cat_list02.append(x) for x in cat_list_with_duplicates02 if x not in cat_list02]

    return cat_list, cat_list02, cat_list_LM, cat_list_OBI
    

def make_lines(m_name, data_dict, document):
    for key, value in data_dict:
        if key == m_name:
            headline = H(outlinelevel=1, stylename=heading03_style, text=key.upper())
            make_add_paragraph("", document)
            document.text.addElement(headline)
            
            for line in value:
                line_ = P(stylename=paragraph_style00, text="")
                boldpart = Span(stylename=boldstyle, text=line[0])
                line_.addElement(boldpart)       
                line_.addText(line[1] + "  -  " + line[2])
                document.text.addElement(line_)


def make_bug_dict(file, m_name): 
    list_00 = []
    for line in file:
        if line[3] == m_name:
            var_01 = (line[0][3:], " ", line[1])      
            list_00.append(var_01)
    bug_dict.update({m_name : list_00}) 
    return bug_dict


def save_with_current_day(document):
    doc_name = datetime.today().strftime('%d.%m.%Y')
    document.save(f"{doc_name}.odt")

def write_changed_files(dict_of_files):
    list_of_lines_to_write = []
    for key, value in dict_of_files.items():
        file_creation_date = get_creation_date(key, long=False)
        file_version = get_version_number(key)
        file_name = value
        list_of_lines_to_write.append(f"{file_name} {file_version} {file_creation_date}")
    return list_of_lines_to_write

def make_add_paragraph(text, document):
    var_01 = P(stylename=paragraph_style00, text = text)
    document.text.addElement(var_01)

def make_add_heading(text, document):
    var_01 = H(outlinelevel=1, stylename=heading01_style, text = text)
    document.text.addElement(var_01)



doc = OpenDocumentText()
doc_s = doc.styles

heading01_style = Style(name="Heading 1", family="paragraph")
heading01_style.addElement(TextProperties(attributes={"fontsize":"16pt","fontweight":"bold", 'fontfamily':"Calibri"}))
heading01_style.addElement(ParagraphProperties(attributes={"textalign": "center"}))
doc_s.addElement(heading01_style)

heading02_style = Style(name="Heading 2", family="paragraph")
heading02_style.addElement(TextProperties(attributes={"fontsize":"14pt","fontweight":"bold", 'fontfamily':"Calibri"}))
heading02_style.addElement(ParagraphProperties(attributes={"textalign": "center"}))
doc_s.addElement(heading02_style)

heading03_style = Style(name="Heading 3", family="paragraph")
heading03_style.addElement(TextProperties(attributes={'fontsize':"11pt",'fontweight':"bold", 'fontfamily':"Calibri"}))
heading03_style.addElement(ParagraphProperties(lineheight="145%"))
doc_s.addElement(heading03_style)


underline_style = Style(name="Underline", family="text")
u_prop = TextProperties(attributes={
    "textunderlinestyle":"solid",
    "textunderlinewidth":"auto",
    "textunderlinecolor":"font-color"
    })

underline_style.addElement(u_prop)
doc_s.addElement(underline_style)

boldstyle = Style(name="Bold", family="text")
boldstyle.addElement(TextProperties(attributes={"fontweight": "bold"}))
doc_s.addElement(boldstyle)

paragraph_style00 = Style(name="paragraph", family="paragraph",)
paragraph_style00.addElement(TextProperties(attributes={"fontsize": "11pt", 'fontfamily':"Calibri"}))
paragraph_style00.addElement(ParagraphProperties(lineheight="135%"))
doc_s.addElement(paragraph_style00)


# App START:

all_paths = make_path_list(source_path)

indexesList = make_list_to_cut(all_paths)

cuted_paths = list_paths(indexesList, all_paths)

target_paths = sort_files_del_from_dict(cuted_paths, all_paths, indexesList)


with open(csv_path, mode="r", encoding="utf-8") as file:
    csv_file = csv.reader(file)
    data = [tuple(row) for row in csv_file]


cat_list, cat_list02, cat_list_LM, cat_list_OBI  = add_lines_to_lists(data, all_programs, kitchen_pro)


#trzeba dopisać do add_lines_to_lists że jak 1 czy 2 lista jest pusta to nic się nie ma dziać dalej
make_add_heading(f"Aktualizacja z dnia {datetime.today().strftime('%d.%m.%Y')} ", doc)
make_add_paragraph("", doc)

heading_0 = H(outlinelevel=1, stylename=heading02_style, text = "")
underlinedpart = Span(stylename=underline_style, text="Zmiany wspólne dla programów CAD Decor PRO i CAD Decor oraz CAD Kuchnie")
heading_0.addElement(underlinedpart)
doc.text.addElement(heading_0)
make_add_paragraph("", doc)


for category in cat_list:
    make_bug_dict(data, category)
    make_lines(category, bug_dict.items(), doc)


# spróbować usunąć powtarzający się kod


if len(cat_list02) > 0:
    make_add_paragraph("", doc)
    heading_1 = H(outlinelevel=1, stylename=heading02_style, text = "")
    underlinedpart = Span(stylename=underline_style, text="Zmiany wspólne dla programów CAD Decor PRO oraz CAD Kuchnie")
    heading_1.addElement(underlinedpart)
    doc.text.addElement(heading_1)
    make_add_paragraph("", doc)

    for category in cat_list02:
        make_bug_dict(data, category)
        make_lines(category, bug_dict.items(), doc)


if len(cat_list_LM) > 0:
    make_add_paragraph("", doc)
    heading_2 = H(outlinelevel=1, stylename=heading02_style, text = "")
    underlinedpart = Span(stylename=underline_style, text="Zmiany dla programu w wersji LM")
    heading_2.addElement(underlinedpart)
    doc.text.addElement(heading_2)
    make_add_paragraph("", doc)

    for category in cat_list_LM:
        make_bug_dict(data, category)
        make_lines(category, bug_dict.items(), doc)


if len(cat_list_OBI) > 0:
    make_add_paragraph("", doc)
    heading_3 = H(outlinelevel=1, stylename=heading02_style, text = "")
    underlinedpart = Span(stylename=underline_style, text="Zmiany dla programu w wersji OBI")
    heading_3.addElement(underlinedpart)
    doc.text.addElement(heading_3)
    make_add_paragraph("", doc)

    for category in cat_list_OBI:
        make_bug_dict(data, category)
        make_lines(category, bug_dict.items(), doc)


make_add_paragraph("", doc)
make_add_paragraph("ZMIENIONE PLIKI:", doc)


list_of_lines = write_changed_files(target_paths)
sorted_list_of_lines = sorted(list_of_lines, key=str.casefold)

for line in sorted_list_of_lines:
    doc.text.addElement(P(stylename=paragraph_style00, text = line))

save_with_current_day(doc)


# dobrze byłoby zrobić otwieranie listy z nazwami programów z json file
# trzeba ogarnąć jak dodawać puste linie





