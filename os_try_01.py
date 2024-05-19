import os
from os import walk
from win32api import *
import time
import csv
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties
from odf.text import H, P, Span


source_path = "C:\\Users\\Czerwiec\\Desktop\\setup ini test"


# def get_path_of_files(folder_path):


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


def get_creation_date(path):
    time_created = time.ctime(os.path.getmtime(path))
    t_obj = time.strptime(time_created)
    return time.strftime("%Y-%m-%d %H:%M:%S", t_obj)


# def get_all_filles_paths(path):


dir_list = []
name_list = []
pathName_dictionary = {}


for dirpath, dirnames, filenames in walk(
    "C:\\Users\\Czerwiec\\Desktop\\setup ini test"
):
    for file_name in filenames:
        path = os.path.abspath(os.path.join(dirpath, file_name))

        name_list.append(file_name)

        # print(os.path.basename(path))
        dir_list.append(path)

        pathName_dictionary[path] = file_name


cuted = []
indexes = []
# new = []

for i, v in enumerate(pathName_dictionary.values()):
    if list(pathName_dictionary.values()).count(v) > 1 and i not in indexes:
        indexes.append(i)
        cuted.append(v)
        
# for i, v in enumerate(name_list):
#     # print(i)
#     # print(v)
#     if name_list.count(v) > 1 and i not in indexes:
#         indexes.append(i)
#         cuted.append(v)

print(indexes)
print(cuted)
print()

nList = []

# for index in indexes:
#     g = list(path_dictionary.keys())
#     nList += [g[index]]

for index in indexes:
    nList += [dir_list[index]]
print(nList)
print()


list_of_creation_time = []

for x in nList:

    print(get_creation_date(x))
    # o = get_creation_date(x)
    list_of_creation_time.append(get_creation_date(x))
    # list_of_creation_time.append(o)
    # for i in nList:
    #     print(x)
    #     list_of_creation_time.append(get_creation_date(x))


p = list_of_creation_time.index(max(list_of_creation_time))
# print(p)
print()





# for var_1,var_2 in enumerate(dir_list):
    # print(var_1, var_2)

list_of_indexes_to_delete = indexes[p + 1 :]
# print(list_of_indexes_to_delete)

test_list = []



for index in sorted(list_of_indexes_to_delete, reverse=True):
    # del dir_list[index]
    # dir_list.pop(index)
    test_list.append(dir_list.pop(index))
    test_list.reverse()

    for i, name in enumerate(list(pathName_dictionary.keys())):
        if index == i:
            del pathName_dictionary[name]
    
    
   

print()

for a, b in pathName_dictionary.items():
    print(a, b)
print()
for var_1,var_2 in enumerate(dir_list):
    print(var_1, var_2)






# for var_1,var_2 in enumerate(dir_list):
#     print(var_1, var_2)

# for a in test_list:
#     print(a)


# for j in dir_list[: h[0]]:
#     print(j)


# for y in h:
#     uu = dir_list[y:y+1]
#     for x in uu:
#         print(x)
#         for i in dir_list:
#             if i == uu[0]:
#                 dir_list.remove(i)
# # print(uu[0])
#     # dir_list.remove(uu[0])
    
# for x in dir_list:
#     print(x)

#     for u in dir_list[y:y+1]:
#         print(dir_list[y:y+1][u])
#     # dir_list.remove(dir_list[y:y+1])
    # if dir_list[y:y+1]
# print(dir_list)








# csv próba

# with open ('C:\\Users\\Czerwiec\\Desktop\\VS Code work\\csv\\tomasz.czerwinski.csv', mode = 'r', encoding='utf-8') as file:
#     csv_file = csv.reader(file)
#     print()
#     list_00 = []
#     for line in csv_file:
#         # if lines[3] == "dokumentacja":
#         #     print(lines[0][3:] + "\t" + lines[1])

#         if line[3] == "wizualizacja":
#             # print(line[0][3:] + "\t" + line[1])
#             list_00.append(line[0][3:] + " " + line[1])
#             # print(list_00)
#             doc = OpenDocumentText()
#             f = H(outlinelevel=1, text = "Wizualizacja")
#             doc.text.addElement(f)
#             # y = line[0][3:] + " " + line[1]
#             # print(lines[0][3:] + lines[1])
#             # print(y)
#     for x in list_00:
#         print(x)
#         i = P(text = x)


#         doc.text.addElement(i)
#         # doc.text.addElement(o)


#         doc.save("this is number 2.odt")
