import os
from os import walk
from win32api import *
from datetime import datetime
import time
import csv
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties
from odf.text import H, P, Span



# laptop:
# source_path = "C:\\Users\\Czerwiec\\Desktop\\setup ini test"

source_path = "C:\\Users\\Czerwiec\\Desktop\\test_folder"
csv_path = "C:\\Users\\Czerwiec\\Desktop\\test_folder\\csv\\tomasz.czerwinski.csv"

bug_dict = {}
bug_dict02 = {}


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
        return time.strftime("%Y-%m-%d")


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
    return dict


def make_list_to_cut(dict):
    indexesToCut = []
    for i, v in enumerate(dict.values()):
        if list(dict.values()).count(v) > 1 and i not in indexesToCut:
            indexesToCut.append(i)
    return indexesToCut


def list_paths(i_list):
    cuted_pathList = []
    for index in i_list:
        pathList = list(paths.keys())
        cuted_pathList += [pathList[index]]
    return cuted_pathList


def make_lines(m_name, data_dict, document):
    for key, value in data_dict:
        if key == m_name:
            headline = H(outlinelevel=1, text=key.upper())
            document.text.addElement(headline)
            for line in value:
                line_ = P(text= line[0] + line[1] + line[2])
                document.text.addElement(line_)



def make_bug_dict(file, m_name): 
    
    list_00 = []
    for line in file:
        if line[3] == m_name:
            # print(line[3])
            var_01 = (line[0][3:], " ", line[1])
            
            list_00.append(var_01)
    bug_dict.update({m_name : list_00}) 
    # print(list_00)
    # bug_dict.update({m_name : list_00})  
    # bug_dict[m_name] = list_00
    return bug_dict


def save_with_current_day(document):
    doc_name = datetime.today().strftime('%d.%m.%Y')
    document.save(f"{doc_name}.odt")

# przywoływanie nazwy pliku:
# var_01 = os.path.basename(path)



all_programs = ['CAD', 'CAD Licencja', 'Dystrybutor baz', 'Instalator', 'Konwerter mdb', 'Marketing', 'Panel aktualizacji', 'Podesty', 'Tłumaczenia', 'aktualizator internetowy', 'analytics', 'asystent pobierania', 'bazy danych', 'deinstalator', 'dokumentacja', 'dot4cad', 'drzwi i okna', 'edytor bazy płytek', 'edytor szafek bazy kuchennej', 'edytor szafek użytkownika', 'edytor ścian', 'elementy dowolne', 'export3D', 'instalator baz danych', 'instalator programu', 'konwerter', 'kreator ścian', 'launcher', 'listwy', 'manager', 'obrót 3d / przesuń', 'obserVeR', 'przeglądarka PDF', 'render', 'szafy wnękowe', 'ukrywacz', 'wersja Leroy Merlin', 'wersja Obi', 'wizualizacja', 'wstawianie elementów wnętrzarskich', 'wstawianie elementów wnętrzarskich w wizualizacji', 'zestawienie płytek, farb, fug, klejów']

kitchen_pro = ['blaty','wstawianie elementów agd', 'wstawianie szafek kuchennych', 'wycena']



# App START:

paths = make_path_list(source_path)

indexesList = make_list_to_cut(paths)

cuted_paths = list_paths(indexesList)

doc = OpenDocumentText()


with open(csv_path, mode="r", encoding="utf-8") as file:
    csv_file = csv.reader(file)

    data = [tuple(row) for row in csv_file]

    cat_list_with_duplicates = []
    cat_list_with_duplicates02 = []


    for i in data:

        
        if i[3].lower() != "projekt" and i[3] in all_programs:
            cat_list_with_duplicates.append(i[3])
            cat_list = []
            [cat_list.append(x) for x in cat_list_with_duplicates if x not in cat_list]

        elif i[3] in kitchen_pro:
            cat_list_with_duplicates02.append(i[3])
            cat_list02 = []
            [cat_list02.append(x) for x in cat_list_with_duplicates02 if x not in cat_list02]




    for category in cat_list:
        make_bug_dict(data, category)
        make_lines(category, bug_dict.items(), doc)

    for category in cat_list02:
        make_bug_dict(data, category)
        make_lines(category, bug_dict.items(), doc)
   
 
save_with_current_day(doc)






