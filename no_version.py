from win32api import *




def get_version_number(file_path): 
  
    try:
        File_information = GetFileVersionInfo(file_path, "\\") 

        print(File_information)
    
        ms_file_version = File_information['FileVersionMS'] 
        ls_file_version = File_information['FileVersionLS'] 

        list_of_versions = [str(HIWORD(ms_file_version)), str(LOWORD(ms_file_version)), 
                str(HIWORD(ls_file_version)), str(LOWORD(ls_file_version))]
        
        return ".".join(list_of_versions)
    
    except: return "-"




print(get_version_number("C:\\Users\\Czerwiec\\Desktop\\test_folder\\translation.xgz"))

# try:
#     File_information = GetFileVersionInfo("C:\\Users\\Czerwiec\\Desktop\\test_folder\\translation.xgz", "\\") 

#     print(File_information)
# except: 
#     print("error")