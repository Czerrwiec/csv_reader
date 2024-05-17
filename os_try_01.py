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
    return  time.strftime("%Y-%m-%d %H:%M:%S", t_obj)


# file_path = r'Z:\Testy\2024-05-14 4.0.7 hotfix\MainFiles\V4_I10x64\DoorsWindows.dll'
  
# version = ".".join(get_version_number(file_path)) 
  
# print(version)


dir_list = []
name_list = []


for (dirpath, dirnames, filenames) in walk("C:\\Users\\Czerwiec\\Desktop\\test_folder"):

    # print(dirpath)
    # print()
    # print(dirnames)
    # print()
    # print(filenames)


    for f in filenames:
        path = (os.path.abspath(os.path.join(dirpath, f)))

        # print(path)
        # print()
        # print(f)
        
        name_list.append(f)
        
        
        # print(os.path.basename(path))
        dir_list.append(path)




# for u in dir_list:
#     print(u)




# print(name_list)




    #     version_data_dict[get_version_number(path)] = get_creation_date(path)

# for i in name_list:
#     print(i)

 
# a= [1, 2, 3, 2, 1, 5, 6, 5, 5, 5]
# result=[d for d, item in enumerate(name_list) if item in name_list[:d]]

# print(result) 



# my_list = [1, 2, 3, 4, 2, 5, 6, 4]
cuted = []
indexes = []
new = []
 
for i, v in enumerate(name_list):
    # print(i)
    # print(v)
    if name_list.count(v) > 1 and i not in indexes:
        indexes.append(i)
        cuted.append(v)
    
   
print(indexes)
print(cuted)
print()

nList = []
for index in indexes:
    nList += [dir_list[index]]
# print(nList)
# print()
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

h = indexes[p+1:]

print(h)

# for t, h in enumerate(dir_list):
#     print(t)
#     print(h)



for j in dir_list[:h[0]]:
    print(j)





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



