import openpyxl
import os

xlsx_path = str(input("please input xlsx file path:\n"))

workbook = openpyxl.load_workbook(xlsx_path)
worksheet = workbook.active

outputPath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output.txt")

table = []

for row in worksheet.iter_rows(min_row=2):
    l = []
    if row[0].value is None:
        continue
    for cell in row:
        l.append(str(cell.value))
    table.append(l)

with open(outputPath, "w") as f:
    for l in table:
        f.write(
            "\t".join(
                [
                    l[1] + "." + worksheet.title + ".",
                    "1",
                    "IN",
                    l[0],
                    l[3] + ("." if l[0] == "CNAME" else ""),
                ]
            )
            + "\n"
        )

print("Done! output file is in " + outputPath)
