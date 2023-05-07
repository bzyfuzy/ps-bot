import openpyxl
import json
# load the workbook
workbook = openpyxl.load_workbook('../data/MAY.xlsx')

# select the worksheet by name
worksheet = workbook['Sheet1']

# select the column to read (e.g. column A)
fname = worksheet['D']
lname = worksheet['E']
data = worksheet['C']
# iterate over the cells in the column

ids = ["blah"]
names = []

for id, cell in enumerate(fname):
    if cell.value is not None and cell.value != "First name":
        names.append({
            "name": f"{cell.value}\r{lname[id].value}",
            "id": ids[-1],
            "custom": data[id].value
        })
        del ids[-1]
# json.dumps(ids)

result = [names[i:i+40] for i in range(0, len(names), 40)]


with open("data_455.json", "w", encoding='utf8') as json_file:
    json.dump(result, json_file, ensure_ascii=False)

per_page = []

for index, dat in enumerate(result):
    index_data = [dat[i:i+5] for i in range(0, len(dat), 5)]
    per_page.append(index_data)

with open("per_page_455.json", "w", encoding='utf8') as json_file:
    json.dump(per_page, json_file, ensure_ascii=False)
