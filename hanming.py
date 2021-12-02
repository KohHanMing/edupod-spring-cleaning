
import csv

filepath = "./.data/source/testinfo.csv"



with open(filepath, 'r', encoding="utf-8-sig") as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
        print(row)
    