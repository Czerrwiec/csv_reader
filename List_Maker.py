import os
from os import walk
from win32api import *
import time
import csv
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties
from odf.text import H, P, Span


#laptop:
# source_path = "C:\\Users\\Czerwiec\\Desktop\\setup ini test"  

source_path = "C:\\Users\\Czerwiec\\Desktop\\test_folder"

# dir_list = []
# filesNames_list = []
# # pathName_dictionary = {}
# cuted = []
# # indexes = []


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

            # filesNames_list.append(file_name)
            # dir_list.append(path)
            # próba z słownikiem
            pathName_dictionary[path] = file_name
    return pathName_dictionary
        


def sort_files_del_from_dict(filesList, dict, indexes):
    list_of_creation_time = []
    # test_list = []
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


# przywoływanie nazwy pliku:
# var_01 = os.path.basename(path)




#App START:

paths = make_path_list(source_path)

indexesList = make_list_to_cut(paths)

cuted_paths = list_paths(indexesList)



for a, b in sort_files_del_from_dict(cuted_paths, paths, indexesList).items():
    print(a)



