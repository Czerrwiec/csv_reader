import csv

list = []

with open(
    "C:\\Users\\Czerwiec\\Desktop\\test_folder\\csv\\tomasz.czerwinski_all.csv", mode="r", encoding="utf-8") as file:
    csv_file = csv.reader(file)

    for line in csv_file:
        list.append(line[3])

deleted_d = []
[deleted_d.append(x) for x in list if x not in deleted_d]
 
sorted_list = sorted(deleted_d)

for i, b in enumerate(sorted_list):
    print(f"{i} - {b}")
    