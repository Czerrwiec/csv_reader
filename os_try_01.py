import os
from os import walk
from win32api import *
import numpy as np

# arr = os.listdir("Z:\\Testy\\2024-05-14 4.0.7 hotfix")

# for i in arr:
#     print(i)
#     if i == "MainFiles":
#         u = os.listdir("Z:\\Testy\\2024-05-14 4.0.7 hotfix\\MainFiles\\V4_I10x64")
#         print(u)
#         print()
#         print()
# # print(arr)

dictionary_01 = {}

dir_list = []
name_list = []


for (dirpath, dirnames, filenames) in walk("Z:\\Testy\\2024-05-14 4.0.7 hotfix"):
    print(f"{dirpath}{filenames}")
   


    # print(filenames)
    # dir_list.append(dirpath)


    # name_list.extend(filenames)
    
# print(dir_list[0])
# print(name_list[0])
# print()
# print(dir_list[1])
# print(name_list[1])
# print()
# print(dir_list[2])
# print(name_list[2])
# print()
# print(dir_list[3])
# print(name_list[3])

    # dir_list.extend(dirpath)

# print(name_list)
# print(dir_list)




# new_list = np.unique(x)
# print(new_list)
    



# def get_version_number(file_path): 
  
#     File_information = GetFileVersionInfo(file_path, "\\") 
  
#     ms_file_version = File_information['FileVersionMS'] 
#     ls_file_version = File_information['FileVersionLS'] 
  
#     return [str(HIWORD(ms_file_version)), str(LOWORD(ms_file_version)), 
#             str(HIWORD(ls_file_version)), str(LOWORD(ls_file_version))] 
  
# file_path = r'Z:\Testy\2024-05-14 4.0.7 hotfix\MainFiles\V4_I10x64\DoorsWindows.dll'
  
# version = ".".join(get_version_number(file_path)) 
  
# print(version)

# path = "Z:\\Testy\\2024-05-14 4.0.7 hotfix\\MainFiles\\V4_I10x64\\DoorsWindows.dll"

# file_info = GetFileVersionInfo(path, "\\") 
# print(file_info)
# print(file_info['FileVersionMS'])
# print(file_info['FileVersionLS'])
# print(str(HIWORD(file_info['FileVersionMS'])))
# print(str(LOWORD(file_info['FileVersionMS'])))