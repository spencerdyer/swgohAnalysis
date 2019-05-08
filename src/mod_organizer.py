import coreapi
import xlwt
import sys

from datetime import datetime

allyCode = 455127936
modTypes = {
    1: "Transmitter", 
    2: "Receiver", 
    3: "Processor", 
    4: "Holo-array", 
    5: "Data-bus", 
    6: "Multiplexer"
    }

setTypes = {
    1: "Health", 
    2: "Defense", 
    3: "Crit Chance", 
    4: "Crit Damage", 
    5: "Tenacity", 
    6: "Potency",
    7: "Offense",
    8: "Speed"
    }

client = coreapi.Client()
schema = client.get("http://swgoh.gg/api/")

action = ["players", "mods", "list"]
params = {"ally_code": allyCode}

result = client.action(schema, action, params)

style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

wb = xlwt.Workbook()

ws = wb.add_sheet('mod list')
ws.write(0, 0, 'slot')
ws.write(0, 1, 'set')
ws.write(0, 2, 'level')
ws.write(0, 3, 'tier')
ws.write(0, 4, 'primary_stat_type')
ws.write(0, 5, 'primary_stat_value')
ws.write(0, 6, 'character')
ws.write(0, 7, 'secondary_stats')

row = 1
col = 0

for item in result["mods"]:
    try:
        ws.write(row, col, modTypes[item['slot']])
        col += 1
        ws.write(row, col, setTypes[item['set']])
        col += 1
        ws.write(row, col, item['level'])
        col += 1
        ws.write(row, col, item['tier'])
        col += 1
        ws.write(row, col, item['primary_stat']['name'])
        col += 1
        ws.write(row, col, item['primary_stat']['display_value'])
        col += 1
        ws.write(row, col, item['character'])
        col += 1

        row +=1
        col = 0
    except Exception:
        col = 0
        
wb.save('mods.xls')