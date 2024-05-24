import csv
import numpy as np

all_programs = ['CAD', 'CAD Licencja', 'Dystrybutor baz', 'Instalator', 'Konwerter mdb', 'Marketing', 'Panel aktualizacji', 'Podesty', 'Tłumaczenia', 'aktualizator internetowy', 'analytics', 'asystent pobierania', 'bazy danych', 'deinstalator', 'dokumentacja', 'dot4cad', 'drzwi i okna', 'edytor bazy płytek', 'edytor szafek bazy kuchennej', 'edytor szafek użytkownika', 'edytor ścian', 'elementy dowolne', 'export3D', 'instalator baz danych', 'instalator programu', 'konwerter', 'kreator ścian', 'launcher', 'listwy', 'manager', 'obrót 3d / przesuń', 'obserVeR', 'przeglądarka PDF', 'render', 'szafy wnękowe', 'ukrywacz', 'wersja Leroy Merlin', 'wersja Obi', 'wizualizacja', 'wstawianie elementów wnętrzarskich', 'wstawianie elementów wnętrzarskich w wizualizacji', 'zestawienie płytek, farb, fug, klejów']

list_0 = []

dict_01 = {}

new_list = []

# with open(
#     "C:\\Users\\Czerwiec\\Desktop\\test_folder\\csv\\tomasz.czerwinski.csv", mode="r", encoding="utf-8") as file:
#     csv_file = csv.reader(file)

    
#     for line in csv_file:
#         if line[3] in all_programs:
#             list_0.append(line[3])
#             new_list = []
#             [new_list.append(x) for x in list_0 if x not in new_list]      
#             for key in range(len(new_list)):
#                 dict_01.update({new_list[key] : []})





with open(
    "C:\\Users\\Czerwiec\\Desktop\\test_folder\\csv\\tomasz.czerwinski.csv", mode="r", encoding="utf-8") as file:
    csv_file = csv.reader(file)
    data = [tuple(row)[3:] for row in csv_file]

    # for line in csv_file:
        
    #     if line[3] in new_list:
    #         hhh = []
    #         x = line[0][3:], " ", line[1]
    #         hhh.append(x)

    
    for i in data:
        print(i)
    