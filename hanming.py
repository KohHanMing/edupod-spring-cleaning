import csv
import os

#check if dir exist if not create it
def check_dir(file_name):
    directory = os.path.dirname(file_name)
    if not os.path.exists(directory):
        os.makedirs(directory)

filepath = "./.data/source/testinfo.csv"
valid_filepath = "./.data/target/valid.csv"
invalid_filepath = "./.data/target/invalid.csv"

check_dir(valid_filepath)
check_dir(invalid_filepath)

with open(filepath, 'r', encoding="utf-8-sig") as f:
    reader = csv.reader(f, delimiter=',')
    valid_file = open(valid_filepath, 'w', encoding='utf-8')
    valid_writer = csv.writer(valid_file, delimiter=',')
    invalid_file = open(invalid_filepath, 'w', encoding='utf-8')
    invalid_writer = csv.writer(invalid_file, delimiter=',')
    for row in reader:
        model = row[1]
        if '-' in model:
            print(model.upper() + " is valid")
        else:
            print(model + " is invalid")

f.close()
valid_file.close()
invalid_file.close()


    