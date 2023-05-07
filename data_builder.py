import openpyxl
import json
# load the workbook
workbook = openpyxl.load_workbook('../data/2023_MAY.xlsx')

# select the worksheet by name
worksheet = workbook['Sheet1']

# select the column to read (e.g. column A)
fname = worksheet['D']
lname = worksheet['E']
data = worksheet['C']
# iterate over the cells in the column

ids = [
    "xkt5r9q51fn90b8102ef",
    "ci463fh15bc10b8102ef",
    "pyq6rnfolw7z0b8102ef",
    "i4pfi7cpeeda0b8102ef",
    "8ssn36uj44ws0b8102ef",
    "9pqhd7mpartm0b8102ef",
    "bz92kncsr1ar0b8102ef",
    "d9wtrunmebkz0b8102ef",
    "b1suc9c7e6dx0b8102ef",
    "lcnwih0rmr1t0b8102ef",
    "tp6aexchh0tq0b8102ef",
    "tooomau9a8c00b8102ef",
    "0smq76fsn72n0b8102ef",
    "mxur6jszl2uj0b8102ef",
    "94rc2r0cfnfk0b8102ef",
    "j1kik46qrjls0b8102ef",
    "17t4rtz6zs920b8102ef",
    "cy2mmoz2fepe0b8102ef",
    "dkonjhyb957n0b8102ef",
    "74yoq7m9bqqk0b8102ef",
    "ldc957tzkjsy0b8102ef",
    "icarnho76zyx0b8102ef",
    "lff7ytyslpfj0b8102ef",
    "0n26krajp6840b8102ef",
    "to1m9za3lkoq0b8102ef",
    "akeqttb64he80b8102ef",
    "ws6inft29wwh0b8102ef",
    "8sybey54ym2f0b8102ef",
    "cmw2ucz6y3x30b8102ef",
    "zaw2ywttonhz0b8102ef",
    "0at7c0rnxkat0b8102ef",
    "w9wp8d0mie6e0b8102ef",
    "nbndkzt4l8da0b8102ef",
    "zk89nplw0iat0b8102ef",
    "12eui41iqawj0b8102ef",
    "9xjkfdj5xsmy0b8102ef",
    "89z7mb4fjz0i0b8102ef",
    "kpp2rjahk1m60b8102ef",
    "qured1ls2jts0b8102ef",
    "5aojkada31xf0b8102ef",
    "m9dxczwz8c030b8102ef",
    "uxzhc12ceck50b8102ef",
]
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
