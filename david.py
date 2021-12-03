import csv
import os

filepath = "./.data/source/rawinfo.csv"
valid_filepath = "./.data/target/valid.csv"
invalid_filepath = "./.data/target/invalid.csv"

def main():
    make_csv(valid_filepath)
    make_csv(invalid_filepath)

    in_file = open(filepath, "r", encoding="utf-8-sig")
    valid_out_file = open(valid_filepath, "w",newline='', encoding="utf-8-sig")
    invalid_out_file = open(invalid_filepath, "w",newline='', encoding="utf-8-sig")

    in_reader = csv.reader(in_file)
    valid_out_writer = csv.writer(valid_out_file)
    invalid_out_writer = csv.writer(invalid_out_file)

    edge_cases = []
    for row in in_reader:
        if "-" not in row[1]: #invalid case
            print(row[1] + " is invalid")
            edge_cases.append(row[1])
            invalid_out_writer.writerow(row)
            continue
        
        print(row[1] + " is valid")
        row[1] = row[1].upper()
        row[1] = row[1].replace(" ","")
        row[1] = row[1].replace("&AMP;", "AND")
        position = row[1].rfind("-")
        num = len(row[1]) - position - 1

        #valid cases
        if num == 1: 
            row[1] = row[1][:position+1] +"00" + row[1][position+1:]
            valid_out_writer.writerow(row)

        elif num == 2:
            row[1] = row[1][:position+1] +"0" + row[1][position+1:]
            valid_out_writer.writerow(row)
        else:
            valid_out_writer.writerow(row)

#create csv if it doesnt exist.
def make_csv(file_name):
    directory = os.path.dirname(file_name)
    if not os.path.exists(directory):
        os.makedirs(directory)

if __name__ == "__main__":
    main()