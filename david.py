import csv
import pandas as pd

filepth = "./.data/source/testinfo.csv"

df = pd.read_csv(filepth, header=None, usecols=[1])
df.to_csv("AllDetails.csv", index=False)
inFile = open("./.data/source/testinfo.csv")
outFile = open('AllDetails.csv', "w",newline='')
outFile.truncate()
inReader =csv.reader(inFile)
outwriter = csv.writer(outFile)

edge_cases = []
for row in inReader:
    row[1] = row[1].upper()
    row[1] = row[1].replace(" ","")
    if "-" not in row[1]:
        edge_cases.append(row[1])
        outwriter.writerow(row)
        continue
    position = row[1].rfind("-")
    num = len(row[1]) - position - 1

    if num == 1:
        row[1] = row[1][:position+1] +"00" + row[1][position+1:]
        outwriter.writerow(row)

    elif num == 2:
        row[1] = row[1][:position+1] +"0" + row[1][position+1:]
        outwriter.writerow(row)
    else:
        outwriter.writerow(row)