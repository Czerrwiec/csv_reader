import os
from os import walk
from win32api import *
import time
import csv
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties
from odf.text import H, P, Span


# laptop:
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


# App START:

paths = make_path_list(source_path)

indexesList = make_list_to_cut(paths)

cuted_paths = list_paths(indexesList)


# for a, b in sort_files_del_from_dict(cuted_paths, paths, indexesList).items():
#     print(a)


# csv próba


def add_lines(list, module_name):
    for i, u in enumerate(list):
        # print(i)
        # print(u)
        if i == 0:
            f = H(outlinelevel=1, text=module_name)
            doc.text.addElement(f)
        # print(u)
        i = P(text=u[0] + u[1] + u[2])
        doc.text.addElement(i)


doc = OpenDocumentText()

list_00 = []
list_01 = []
list_02 = []

bug_dict = {}

with open(
    "C:\\Users\\Czerwiec\\Desktop\\test_folder\\csv\\tomasz.czerwinski.csv", mode="r", encoding="utf-8"
) as file:
    csv_file = csv.reader(file)
    print()

    # list_01 = []

    for line in csv_file:
        print(line[0][3:] + "\t" + line[1])
        # print(line[3])
        # print(line)

        if line[3] == "dokumentacja":
            x = (line[0][3:], " ", line[1])
            list_00.append(x)
            bug_dict["dokumentacja"] = list_00

        elif line[3] == "konwerter":
            u = (line[0][3:], " ", line[1])
            list_02.append(u)
            bug_dict["konwerter"] = list_02


for i, j in bug_dict.items():
    if i == "dokumentacja":
        f = H(outlinelevel=1, text=i.upper())
        doc.text.addElement(f)
        # for h in j:
        for l in j:
            # print(l[0])
            i = P(text=l[0] + l[1] + l[2])
            doc.text.addElement(i)

    elif i == "konwerter":
        f = H(outlinelevel=1, text=i.upper())
        doc.text.addElement(f)
        # for h in j:
        for l in j:
            # print(l[0])
            i = P(text=l[0] + l[1] + l[2])
            doc.text.addElement(i)

    # i = P(text = j[0][0] + j[0][1] + j[0][2])
    # doc.text.addElement(i)

# bugList = list(bug_dict.values())[0]

# bugList2 = list(bug_dict.values())[1]

# bugList = list(bug_dict.values())[0]

# bugList2 = list(bug_dict.values())[1]


# for i, u in enumerate(bugList):
#     # print(i)
#     # print(u)
#     if i == 0:
#         f = H(outlinelevel=1, text = "DOKUMENTACJA")
#         doc.text.addElement(f)
#     # print(u)
#     i = P(text = u[0] + u[1] + u[2])
#     doc.text.addElement(i)


# for i, u in enumerate(bugList2):
#     if i == 0:
#         f = H(outlinelevel=1, text = "WIZUALIZACJA")
#         doc.text.addElement(f)
#     # print(u)
#     i = P(text = u[0] + u[1] + u[2])
#     doc.text.addElement(i)


# add_lines(bugList, "DOKUMENTACJA")
# add_lines(bugList2, "WIZUALIZACJA")


doc.save("this is number 2.odt")


for a in bug_dict.items():
    print(a)


# for i, u in list(enumerate(bug_dict.values())):
#     print(u)
#     for j in u:
#         print(j)

# print()


# for a in bug_dict.items():
#     print(a)
