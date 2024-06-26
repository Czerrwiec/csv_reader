import os
import shutil
import sys
import time
from tkinter import *
from tkinter import filedialog
from customtkinter import *
from CustomTkinterMessagebox import CTkMessagebox
from os import walk
from win32api import *
from datetime import datetime
import time
import csv
from odf.opendocument import OpenDocumentText
from odf.style import Style, TextProperties, ParagraphProperties
from odf.text import H, P, Span


path = None
csv_file = None
folder_path = None
csv_list = []

all_programs = [
    "CAD",
    "CAD Licencja",
    "Dystrybutor baz",
    "Instalator",
    "Konwerter mdb",
    "Marketing",
    "Panel aktualizacji",
    "Podesty",
    "Tłumaczenia",
    "aktualizator internetowy",
    "analytics",
    "asystent pobierania",
    "bazy danych",
    "deinstalator",
    "dokumentacja",
    "dot4cad",
    "drzwi i okna",
    "edytor bazy płytek",
    "edytor szafek bazy kuchennej",
    "edytor szafek użytkownika",
    "edytor ścian",
    "elementy dowolne",
    "export3D",
    "instalator baz danych",
    "instalator programu",
    "konwerter",
    "kreator ścian",
    "launcher",
    "listwy",
    "manager",
    "obrót 3d / przesuń",
    "obserVeR",
    "przeglądarka PDF",
    "render",
    "szafy wnękowe",
    "ukrywacz",
    "wersja Leroy Merlin",
    "wersja Obi",
    "wizualizacja",
    "wstawianie elementów wnętrzarskich",
    "wstawianie elementów wnętrzarskich w wizualizacji",
    "zestawienie płytek, farb, fug, klejów",
    "CAD Rozkrój"
]

kitchen_pro = [
    "blaty",
    "wstawianie elementów agd",
    "wstawianie szafek kuchennych",
    "wycena",
]

bug_dict = {}

def load_data():
    global path
    data_file = f"{get_script_path()}\\path.txt"
    try:
        with open(data_file, "r") as f:
            for line in f:
                path = line
    except:
        FileNotFoundError

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

def make_path_list(folder_p):
    pathName_dictionary = {}
    for dirpath, dirnames, filenames in walk(folder_p):
        for file_name in filenames:
            path = os.path.abspath(os.path.join(dirpath, file_name))
            pathName_dictionary[path] = file_name
    return pathName_dictionary



def sort_files_del_from_dict(file_list, bug_d):
    list_of_names = []
    list_without_duplicates = []

    for i in file_list:
        list_of_names.append(os.path.basename(i))

    [list_without_duplicates.append(x) for x in list_of_names if x not in list_without_duplicates]

    for n in range(len(list_without_duplicates)):
        condition = lambda x: os.path.basename(x) == list_without_duplicates[n]

        filtered_list = [x for x in file_list if condition(x)]
        time_list = []

        for item in filtered_list:
            time_list.append(os.path.getmtime(item))
            x = time_list.index(max(time_list))

        del filtered_list[x]

        for a in filtered_list:
            del bug_d[a]

    listed_dict = [x for x in bug_d.keys()]

    for x in listed_dict:
        if x.endswith(".txt"):
            del bug_d[x]

    return bug_d

def make_list_to_cut(dict):
    indexesToCut = []
    v_list = [v for v in dict.values()]

    for i, v in enumerate(dict.values()):
        if v_list.count(v) > 1 and i not in indexesToCut:
            indexesToCut.append(i)
    return indexesToCut

def list_paths(i_list, paths):

    cuted_pathList = []
    k_list = [k for k in paths.keys()]
    for index in i_list:
        pathList = k_list
        cuted_pathList += [pathList[index]]
    return cuted_pathList

def add_lines_to_lists(data_file, all_list, rest_list):
    cat_list_with_duplicates = []
    cat_list_with_duplicates02 = []
    cat_list_LM_with_duplicates = []
    cat_list_OBI_with_duplicates = []
    cat_list_rozkroj_with_duplicates = []
    cat_list_OBI = []
    cat_list_LM = []
    cat_list02 = []
    cat_list = []
    cat_list_rozkroj = []

    for line in data_file:
        if line[3].lower() == "wersja obi":
            cat_list_OBI_with_duplicates.append(line[3])
            cat_list_OBI = []
            [
                cat_list_OBI.append(x)
                for x in cat_list_OBI_with_duplicates
                if x not in cat_list_OBI
            ]

        elif line[3].lower() == "wersja leroy merlin":
            cat_list_LM_with_duplicates.append(line[3])
            cat_list_LM = []
            [
                cat_list_LM.append(x)
                for x in cat_list_LM_with_duplicates
                if x not in cat_list_LM
            ]

        elif line[3].lower() == "cad rozkrój":
            cat_list_rozkroj_with_duplicates.append(line[3])
            cat_list_rozkroj = []
            [
                cat_list_rozkroj.append(x)
                for x in cat_list_rozkroj_with_duplicates
                if x not in cat_list_rozkroj
            ]

        elif line[3].lower() != "projekt" and line[3] in all_list:
            cat_list_with_duplicates.append(line[3])
            cat_list = []
            [cat_list.append(x) for x in cat_list_with_duplicates if x not in cat_list]

        elif line[3] in rest_list:
            cat_list_with_duplicates02.append(line[3])
            cat_list02 = []
            [
                cat_list02.append(x)
                for x in cat_list_with_duplicates02
                if x not in cat_list02
            ]

    return cat_list, cat_list02, cat_list_LM, cat_list_OBI, cat_list_rozkroj

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
    bug_dict.update({m_name: list_00})
    return bug_dict

def save_with_current_day(document):
    doc_name = datetime.today().strftime("%d.%m.%Y")
    file = os.path.join(get_script_path(), doc_name + ".odt")
    document.save(file)

def write_changed_files(dict_of_files):
    list_of_lines_to_write = []
    for key, value in dict_of_files.items():
        file_creation_date = get_creation_date(key, long=False)
        file_version = get_version_number(key)
        file_name = value
        list_of_lines_to_write.append(
            f"{file_name} {file_version} {file_creation_date}"
        )
    return list_of_lines_to_write

def make_add_paragraph(text, document):
    var_01 = P(stylename=paragraph_style00, text=text)
    document.text.addElement(var_01)

def make_add_heading(text, document):
    var_01 = H(outlinelevel=1, stylename=heading01_style, text=text)
    document.text.addElement(var_01)

def get_and_display_path():
    choice = var_0.get()
    global folder_path

    if choice == 1:
        for n in paths_dir:
            if n.endswith("hotfix") == True and n.endswith("_op") == False:
                hotfix_cat = n
                hotfix_path = os.path.join(path + "\\" + hotfix_cat)
        label1.configure(text=hotfix_cat)
        folder_path = hotfix_path
        button3.configure(state="normal")

    elif choice == 2:
        for n in paths_dir:
            if (
                "feat" not in n
                and n.endswith("_op") == False
                and n.endswith("hotfix") == False
                and n != "_dok"
            ):
                new_version_cat = n
                new_version_path = os.path.join(path + "\\" + new_version_cat)
        label1.configure(text=new_version_cat)
        folder_path = new_version_path
        button3.configure(state="normal")

def get_script_path():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

def ask_for_dir(value):
    global folder_path
    global csv_file
    if value == 1:
        csv_dir = filedialog.askopenfilename()
        if csv_dir.endswith(".csv"):
            s_dir = csv_dir.split("/")
            label0.configure(text=s_dir[-1])
            csv_file = csv_dir
        else:
            csv_file = ""
            label0.configure(text="Wybierz prawidłowy plik")

    elif value == 2:
        pack_dir = filedialog.askdirectory()
        s_pack_dir = pack_dir.split("/")

        label1.configure(text=s_pack_dir[-1])
        folder_path = pack_dir
        button3.configure(state="normal")

def make_list(folder, csv_f):

    doc = OpenDocumentText()
    doc_s = doc.styles

    heading01_style = Style(name="Heading 1", family="paragraph")
    heading01_style.addElement(
        TextProperties(
            attributes={"fontsize": "16pt", "fontweight": "bold", "fontfamily": "Calibri"}
        )
    )
    heading01_style.addElement(ParagraphProperties(attributes={"textalign": "center"}))
    doc_s.addElement(heading01_style)

    heading02_style = Style(name="Heading 2", family="paragraph")
    heading02_style.addElement(
        TextProperties(
            attributes={"fontsize": "14pt", "fontweight": "bold", "fontfamily": "Calibri"}
        )
    )
    heading02_style.addElement(ParagraphProperties(attributes={"textalign": "center"}))
    doc_s.addElement(heading02_style)

    heading03_style = Style(name="Heading 3", family="paragraph")
    heading03_style.addElement(
        TextProperties(
            attributes={"fontsize": "11pt", "fontweight": "bold", "fontfamily": "Calibri"}
        )
    )
    heading03_style.addElement(ParagraphProperties(lineheight="145%"))
    doc_s.addElement(heading03_style)


    underline_style = Style(name="Underline", family="text")
    u_prop = TextProperties(
        attributes={
            "textunderlinestyle": "solid",
            "textunderlinewidth": "auto",
            "textunderlinecolor": "font-color",
        }
    )

    underline_style.addElement(u_prop)
    doc_s.addElement(underline_style)

    boldstyle = Style(name="Bold", family="text")
    boldstyle.addElement(TextProperties(attributes={"fontweight": "bold"}))
    doc_s.addElement(boldstyle)

    paragraph_style00 = Style(
        name="paragraph",
        family="paragraph",
    )
    paragraph_style00.addElement(
        TextProperties(attributes={"fontsize": "11pt", "fontfamily": "Calibri"})
    )
    paragraph_style00.addElement(ParagraphProperties(lineheight="135%"))
    doc_s.addElement(paragraph_style00)

    all_paths = make_path_list(folder)

    indexesList = make_list_to_cut(all_paths)

    cuted_paths = list_paths(indexesList, all_paths)

    target_paths = sort_files_del_from_dict(cuted_paths, all_paths)


    with open(csv_f, mode="r", encoding="utf-8") as file:
        csv_file = csv.reader(file)
        data = [tuple(row) for row in csv_file]

    cat_list, cat_list02, cat_list_LM, cat_list_OBI, cat_list_rozkroj = add_lines_to_lists(
        data, all_programs, kitchen_pro
    )


    make_add_heading(
        f"Aktualizacja z dnia {datetime.today().strftime('%d.%m.%Y')} ", doc
    )
    make_add_paragraph("", doc)

    if len(cat_list) > 0:
        heading_0 = H(outlinelevel=1, stylename=heading02_style, text="")
        underlinedpart = Span(
            stylename=underline_style,
            text="Zmiany wspólne dla programów CAD Decor PRO i CAD Decor oraz CAD Kuchnie",
        )
        heading_0.addElement(underlinedpart)
        doc.text.addElement(heading_0)
        make_add_paragraph("", doc)

        for category in cat_list:
            make_bug_dict(data, category)
            make_lines(category, bug_dict.items(), doc)

    if len(cat_list02) > 0:
        make_add_paragraph("", doc)
        heading_1 = H(outlinelevel=1, stylename=heading02_style, text="")
        underlinedpart = Span(
            stylename=underline_style,
            text="Zmiany wspólne dla programów CAD Decor PRO oraz CAD Kuchnie",
        )
        heading_1.addElement(underlinedpart)
        doc.text.addElement(heading_1)
        make_add_paragraph("", doc)

        for category in cat_list02:
            make_bug_dict(data, category)
            make_lines(category, bug_dict.items(), doc)

    if len(cat_list_LM) > 0:
        make_add_paragraph("", doc)
        heading_2 = H(outlinelevel=1, stylename=heading02_style, text="")
        underlinedpart = Span(
            stylename=underline_style, text="Zmiany dla programu w wersji LM"
        )
        heading_2.addElement(underlinedpart)
        doc.text.addElement(heading_2)
        make_add_paragraph("", doc)

        for category in cat_list_LM:
            make_bug_dict(data, category)
            make_lines(category, bug_dict.items(), doc)

    if len(cat_list_OBI) > 0:
        make_add_paragraph("", doc)
        heading_3 = H(outlinelevel=1, stylename=heading02_style, text="")
        underlinedpart = Span(
            stylename=underline_style, text="Zmiany dla programu w wersji OBI"
        )
        heading_3.addElement(underlinedpart)
        doc.text.addElement(heading_3)
        make_add_paragraph("", doc)

        for category in cat_list_OBI:
            make_bug_dict(data, category)
            make_lines(category, bug_dict.items(), doc)


    if len(cat_list_rozkroj) > 0:
        make_add_paragraph("", doc)
        heading_4 = H(outlinelevel=1, stylename=heading02_style, text="")
        underlinedpart=Span(stylename=underline_style, text="Zmiany dla programu CAD Rozkrój")
        heading_4.addElement(underlinedpart)
        doc.text.addElement(heading_4)
        make_add_paragraph("", doc)

        for category in cat_list_rozkroj:
            make_bug_dict(data, category)
            make_lines(category, bug_dict.items(), doc)
    

    make_add_paragraph("", doc)
    make_add_paragraph("ZMIENIONE PLIKI:", doc)

    list_of_lines = write_changed_files(target_paths)
    sorted_list_of_lines = sorted(list_of_lines, key=str.casefold)

    for line in sorted_list_of_lines:
        doc.text.addElement(P(stylename=paragraph_style00, text=line))

    save_with_current_day(doc)

def copy_pack(folder): 
    global a_dirs_to_delete
    a_dirs_to_delete = []
    
    try:
        new_dir = get_script_path() + "/" + datetime.today().strftime("%d-%m-%Y") + " x64"
        shutil.copytree(folder, new_dir)
    except: 
        FileExistsError
        new_dir = get_script_path() + "/" + datetime.today().strftime("%d-%m-%Y") + " x64(1)"
        shutil.copytree(folder, new_dir)

    directory_list =  make_path_list(new_dir)

    i_list = make_list_to_cut(directory_list)

    if len(i_list) > 0:
        cuted_p = list_paths(i_list, directory_list)
       
        t_paths = sort_files_del_from_dict(cuted_p, directory_list)

        del_files_and_dirs(new_dir, t_paths)
        
    move_odt_file(new_dir)
    

def del_files_and_dirs(dir, paths):

    for p in make_path_list(dir):
        dir_to_remove = []
        dir_to_remove_02 = []

        if p not in paths:
            os.remove(p)
            if p.endswith('.txt') == False:
                dir_to_remove.append(os.path.dirname(p))
           
        for i in dir_to_remove:
            dir_to_remove_02.append(os.path.dirname(i))
            os.rmdir(i)
        
        for p in dir_to_remove_02:
            
            if p not in a_dirs_to_delete:
                a_dirs_to_delete.append(p)

    for i in a_dirs_to_delete:
        shutil.rmtree(i)

def move_odt_file(new_dir):
    dir = os.listdir(get_script_path())
    odt_list = []
    try:
        for file in dir:
            if file.endswith(".odt") == True:
                file_p = get_script_path() + "\\" + file
                odt_list.append(file_p)
                odt_list.sort(key=os.path.getmtime, reverse=True)
                shutil.move(odt_list[0], new_dir)
    except:
        print("error")

def move_csv_files():
    file_p = ''
    files_list = []
    data_file = f"{get_script_path()}\\csv_path.txt"
    csv_list02 = []
    try:
        with open(data_file, "r") as f:
            for line in f:
                file_p = line
    except:
        FileNotFoundError
    
    try:
        files_list = os.listdir(file_p)
    except:
        FileNotFoundError


    if len(files_list) > 0:
        try:
            for item in files_list:
                if item.endswith(".csv") == True:
                    file = file_p + "\\" + item
                    csv_list02.append(file)
                    csv_list02.sort(key=os.path.getmtime, reverse=True)
            shutil.move(csv_list02[0], get_script_path())
        except:
            print("error in move csv")
    
    get_default_csv()


def get_default_csv():
    global csv_file
    list = os.listdir(get_script_path())
    try:
        paths_dir = os.listdir(path)
    except:
        FileNotFoundError
        print("error in get default csv")
    try:
        for file in list:
            if file.endswith(".csv") == True:
                file_p = get_script_path() + "\\" + file
                csv_list.append(file_p)
                csv_list.sort(key=os.path.getmtime, reverse=True)
                splited_csv = csv_list[0].split("\\")
        label0.configure(text=splited_csv[-1]) 
        csv_file = csv_list[0]
    except:
        label0.configure(text="Wybierz plik csv")
    return csv_file, paths_dir


doc = OpenDocumentText()
doc_s = doc.styles

heading01_style = Style(name="Heading 1", family="paragraph")
heading01_style.addElement(
    TextProperties(
        attributes={"fontsize": "16pt", "fontweight": "bold", "fontfamily": "Calibri"}
    )
)
heading01_style.addElement(ParagraphProperties(attributes={"textalign": "center"}))
doc_s.addElement(heading01_style)

heading02_style = Style(name="Heading 2", family="paragraph")
heading02_style.addElement(
    TextProperties(
        attributes={"fontsize": "14pt", "fontweight": "bold", "fontfamily": "Calibri"}
    )
)
heading02_style.addElement(ParagraphProperties(attributes={"textalign": "center"}))
doc_s.addElement(heading02_style)

heading03_style = Style(name="Heading 3", family="paragraph")
heading03_style.addElement(
    TextProperties(
        attributes={"fontsize": "11pt", "fontweight": "bold", "fontfamily": "Calibri"}
    )
)
heading03_style.addElement(ParagraphProperties(lineheight="145%"))
doc_s.addElement(heading03_style)


underline_style = Style(name="Underline", family="text")
u_prop = TextProperties(
    attributes={
        "textunderlinestyle": "solid",
        "textunderlinewidth": "auto",
        "textunderlinecolor": "font-color",
    }
)

underline_style.addElement(u_prop)
doc_s.addElement(underline_style)

boldstyle = Style(name="Bold", family="text")
boldstyle.addElement(TextProperties(attributes={"fontweight": "bold"}))
doc_s.addElement(boldstyle)

paragraph_style00 = Style(
    name="paragraph",
    family="paragraph",
)
paragraph_style00.addElement(
    TextProperties(attributes={"fontsize": "11pt", "fontfamily": "Calibri"})
)
paragraph_style00.addElement(ParagraphProperties(lineheight="135%"))
doc_s.addElement(paragraph_style00)


gui = CTk()

gui.geometry("350x500")
gui.title("Generator")
gui.resizable(False, False)

gui.eval('tk::PlaceWindow . center')

var_0 = IntVar()

button00 = CTkButton(
    gui,
    text="*",
    width=20,
    height=8,
    font=("Consolas", 14),
    command=move_csv_files
)
button00.pack(pady=(10, 0), padx=(10,0), anchor=W)
button00.configure(border_width=1)

label0 = CTkLabel(gui, height=20, font=("Consolas", 14))
label0.pack(pady=(0,10), anchor=CENTER)


load_data()

if path == None:
    result = CTkMessagebox.messagebox(
        title="Błąd!",
        text='Dla automatycznego wykrywania paczki \n dodaj plik "path.txt" ze ścieżką.',
        button_text="OK",
    )

csv_file, paths_dir = get_default_csv()

button0 = CTkButton(
    gui,
    text="Wybierz csv",
    width=160,
    height=40,
    font=("Consolas", 16),
    command=lambda v=1: ask_for_dir(v),
)
button0.pack(pady=(5, 25), anchor=CENTER)

frame = CTkFrame(gui)
frame.configure(border_width=2, fg_color="transparent")
frame.pack()

r_check0 = CTkRadioButton(
    frame,
    text="Hotfix",
    variable=var_0,
    value=1,
    font=("Consolas", 14),
    command=get_and_display_path,
)
r_check0.pack(pady=20, padx=(15, 5), side="left", anchor=N)

r_check1 = CTkRadioButton(
    frame,
    text="Nowa wersja",
    variable=var_0,
    value=2,
    font=("Consolas", 14),
    command=get_and_display_path,
)
r_check1.pack(pady=20, padx=(5, 15), side="left", anchor=N)

label1 = CTkLabel(gui, text="", height=40, font=("Consolas", 14))
label1.pack(pady=15, anchor=CENTER)

button1 = CTkButton(
    gui,
    text="Wybierz paczkę",
    width=160,
    height=40,
    font=("Consolas", 16),
    command=lambda v=2: ask_for_dir(v),
)
button1.pack(pady=(5, 25), anchor=CENTER)

button2 = CTkButton(
    gui,
    text="Generuj listę",
    width=160,
    height=40,
    font=("Consolas", 16),
    command=lambda: make_list(folder_path, csv_file),
)
button2.pack(pady=(5, 25), anchor=CENTER)

button3 = CTkButton(
    gui,
    text="Utwórz paczkę",
    width=160,
    height=40,
    font=("Consolas", 16),
    command=lambda: copy_pack(folder_path)
)
button3.pack(pady=(5, 25), anchor=CENTER)
button3.configure(state="disabled", hover=True, border_width=1, border_color="gold")

gui.mainloop()

