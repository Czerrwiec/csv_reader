import os
import sys
import time
from tkinter import *
from tkinter import filedialog

path = "F:\\Testy"
paths_dir = os.listdir(path)


def get_creation_date(path, long=True):
    time_created = time.ctime(os.path.getmtime(path))
    t_obj = time.strptime(time_created)
    if long == True:
        return time.strftime("%Y-%m-%d %H:%M:%S", t_obj)
    elif long == False:
        return time.strftime("%d.%m.%Y", t_obj)

def get_and_display_path():
    choice = var_0.get()

    if choice == 1:
        for n in paths_dir:
            if n.endswith("hotfix") == True and n.endswith("_op") == False:
                hotfix_cat = n
                hotfix_path = os.path.join(path + "\\" + hotfix_cat)
        label1.config(text = f"ścieżka: {hotfix_path}")
    elif choice == 2:
        for n in paths_dir:
             if "feat" not in n and n.endswith("_op") == False and n.endswith("hotfix") == False and n != "_dok":
                new_version_cat = n
                new_version_path = os.path.join(path + "\\" + new_version_cat)
        label1.config(text = f"ścieżka: {new_version_path}")
    

def get_script_path():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

def ask_for_dir(value):
    if value == 2:
        directory = filedialog.askdirectory()
        label1.config(text = f"ścieżka: {directory}")
    elif value == 1:
        directory = filedialog.askopenfilename()
        label0.config(text = f"ścieżka: {directory}")



gui = Tk()
gui.geometry("300x200")
gui.title("Generator")

var_0 = IntVar()

label0 = Label(gui)
label0.pack()


button0 = Button(gui, text='zmień csv', width=15, command=lambda v=1 : ask_for_dir(v))
button0.pack( anchor = W ) 

r_check0 = Radiobutton(gui, text="hotfix", variable=var_0, value=1, command=get_and_display_path)
r_check0.pack( anchor = W )

r_check1 = Radiobutton(gui, text="nowa wersja", variable=var_0, value=2, command=get_and_display_path)
r_check1.pack( anchor = W )

label1 = Label(gui)
label1.pack()

button1 = Button(gui, text='zmień paczkę', width=15, command=lambda v=2 : ask_for_dir(v))
button1.pack( anchor = W ) 


list = os.listdir(get_script_path())
csv_list = []

try:
    for file in list:
        if file.endswith(".csv") == True:
            h = get_script_path() + "\\" + file
            csv_list.append(h)
            csv_list.sort(key=os.path.getctime, reverse=True)
    label0.config(text = "Plik csv: " + get_script_path() + "\\" + csv_list[0])
except:
    label0.config(text = "Wybierz plik csv", bg="yellow")

gui.mainloop()

