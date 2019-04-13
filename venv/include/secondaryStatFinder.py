import coreapi
import xlwt
import sys

from datetime import datetime

def findStat(stat, secondaries):
    for secondary in secondaries:
        if secondary['name'] == stat:
            return secondary['display_value']
    return 0

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
ws.write(0, 2, 'character')
ws.write(0, 3, 'speed')

row = 1
col = 0

for item in result["mods"]:
    try:
        ws.write(row, col, modTypes[item['slot']])
        col += 1
        ws.write(row, col, setTypes[item['set']])
        col += 1
        ws.write(row, col, item['character'])
        col += 1
        secondary = findStat('Speed', item['secondary_stats'])
        if secondary > 0:
            ws.write(row, col, secondary)

        row +=1
        col = 0
    except Exception:
        col = 0
        
wb.save('mods2.xls')