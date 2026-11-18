import csv

with open("RosterCSVs/Baseball.csv", newline="", encoding="utf-8-sig") as f:
    for row in csv.DictReader(f):
        print(f'{row["Last Name"]}, {row["First Name"]} ({row["Person Populi ID"]})')