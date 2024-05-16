import os
from os import walk
from win32api import *
import time




def get_version_number(file_path): 
  
    File_information = GetFileVersionInfo(file_path, "\\") 
  
    ms_file_version = File_information['FileVersionMS'] 
    ls_file_version = File_information['FileVersionLS'] 

    list_of_versions = [str(HIWORD(ms_file_version)), str(LOWORD(ms_file_version)), 
            str(HIWORD(ls_file_version)), str(LOWORD(ls_file_version))]
    
    return ".".join(list_of_versions)

    # return [str(HIWORD(ms_file_version)), str(LOWORD(ms_file_version)), 
    #         str(HIWORD(ls_file_version)), str(LOWORD(ls_file_version))] 
                

def get_creation_date(path):
    time_created = time.ctime(os.path.getmtime(path))
    t_obj = time.strptime(time_created)
    return  time.strftime("%d-%m-%Y", t_obj)




# file_path = r'Z:\Testy\2024-05-14 4.0.7 hotfix\MainFiles\V4_I10x64\DoorsWindows.dll'
  
# version = ".".join(get_version_number(file_path)) 
  
# print(version)




dictionary_01 = {}

list_of_data = []

dir_list = []
name_list = []






for (dirpath, dirnames, filenames) in walk("C:\\Users\\Czerwiec\\Desktop\\test_folder"):

    for f in filenames:

        path = (os.path.abspath(os.path.join(dirpath, f)))
        
        print(path)
        
        print(f) 
        print(get_version_number(path))
        print(get_creation_date(path))

        x = f + " " + get_version_number(path) + " " + get_creation_date(path)

        list_of_data.append(x)
     
       

        dir_list.append(path)

print(list_of_data)


# for i in dir_list:
#     print(".".join(get_version_number(i)))
    







# new_list = np.unique(x)
# print(new_list)
    





# path = "Z:\\Testy\\2024-05-14 4.0.7 hotfix\\MainFiles\\V4_I10x64\\DoorsWindows.dll"

# file_info = GetFileVersionInfo(path, "\\") 
# print(file_info)
# print(file_info['FileVersionMS'])
# print(file_info['FileVersionLS'])
# print(str(HIWORD(file_info['FileVersionMS'])))
# print(str(LOWORD(file_info['FileVersionMS'])))