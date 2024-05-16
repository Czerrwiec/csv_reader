import os
from os import walk
from win32api import *
import time
import csv
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties
from odf.text import H, P, Span




def get_version_number(file_path): 
    try:
        File_information = GetFileVersionInfo(file_path, "\\") 
    
        ms_file_version = File_information['FileVersionMS'] 
        ls_file_version = File_information['FileVersionLS'] 

        list_of_versions = [str(HIWORD(ms_file_version)), str(LOWORD(ms_file_version)), 
                str(HIWORD(ls_file_version)), str(LOWORD(ls_file_version))]
        
        return ".".join(list_of_versions)
    
    except: return "-"
  
                

def get_creation_date(path):
    time_created = time.ctime(os.path.getmtime(path))
    t_obj = time.strptime(time_created)
    return  time.strftime("%d-%m-%Y", t_obj)




# file_path = r'Z:\Testy\2024-05-14 4.0.7 hotfix\MainFiles\V4_I10x64\DoorsWindows.dll'
  
# version = ".".join(get_version_number(file_path)) 
  
# print(version)


path_name_dict = {}
name_version_dict = {}
version_data_dict = {}



list_of_data = []

dir_list = []
name_list = []






for (dirpath, dirnames, filenames) in walk("C:\\Users\\Czerwiec\\Desktop\\test_folder"):

    for f in filenames:

        path = (os.path.abspath(os.path.join(dirpath, f)))
        
        # print(path)
        
        # print(f) 
        # print(get_version_number(path))
        # print(get_creation_date(path))

        x = f + " " + get_version_number(path) + " " + get_creation_date(path)


       
        list_of_data.append((f, get_version_number(path), get_creation_date(path)))
    
        dir_list.append(path)


path_name_dict = dict(zip(dir_list, list_of_data))

# print(path_name_dict)

# for i in path_name_dict.keys():
    
#     print(path_name_dict[i][0])




list = (list(path_name_dict.values()))

for i in list:
    print(i[0])
    
    # print(list[i][0])

    # if i[0] == list[i][0]:
    #     print("lol")

# print(list[1][0])

# print(type(list[i][0]))

# print(y[0][0])
    # print(path_name_dict[i][0])
# print(dir_list[0])
# print(list_of_data[0])

    #     version_data_dict[get_version_number(path)] = get_creation_date(path)


    #     # for y in version_data_dict:
    #     #     print(y)
    # print(version_data_dict)
    
    # for y, t in version_data_dict.items():
    #     print(y, t)
 




# print(list_of_data)
# print()
# print()

# print(path_name_dict)

# print(version_data_dict)
# print(name_version_dict)

# for k, v in path_name_dict.items():
#     print(k, v)
# for i in dir_list:
#     print(".".join(get_version_number(i)))
    




#csv próba 

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



