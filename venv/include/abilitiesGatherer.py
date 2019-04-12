import coreapi
import xlwt
import sys

from datetime import datetime

# Initialize a client & load the schema document
client = coreapi.Client()
schema = client.get("http://swgoh.gg/api/")

# Interact with the API endpoint
action = ["abilities", "list"]
result = client.action(schema, action)

style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

wb = xlwt.Workbook()

currentCharacter = ''

row = 0
col = 0

for item in result:

    try:
        if currentCharacter != item['character_base_id']:
            ws = wb.add_sheet(item['character_base_id'])
            ws.write(0, 0, 'name')
            ws.write(0, 1, 'image')
            ws.write(0, 2, 'url')
            ws.write(0, 3, 'combat_type')
            currentCharacter = item['character_base_id']
            row = 1
            col = 0

        if currentCharacter == item['character_base_id']:
            ws.write(row, col, item['name'])
            col += 1
            ws.write(row, col, item['image'])
            col += 1
            ws.write(row, col, item['url'])
            col += 1
            ws.write(row, col, item['combat_type'])

        row +=1
        col = 0
    except Exception:
        col = 0

wb.save('abilities.xls')