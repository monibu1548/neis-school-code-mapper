import json
import xlrd
import xlwt
import requests
import urllib.parse

loc = "incheon"
_url = "https://stu.ice.go.kr/spr_ccm_cm01_100.ws"
cookie = "Cookie: WMONID=PY4eXA4KY4i; Veraport20Use=N; JSESSIONID=RH2C7QE1FnNaax0mup5tT9Fi424cwA9F0Pf1t4hAGcUArRs14aNFg5tOhRzd3fpn.ice-pacwas2_servlet_stuwas"


excel_filename = loc + '.xlsx'
result_filename = 'after_' + loc + '.xlsx'
wb = xlrd.open_workbook(filename = excel_filename)
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('result')

ws = wb.sheet_by_name('Sheet1')

def code(name):
    headers = {'Content-Type': 'application/json; charset=utf-8', 'Cookie': cookie}
    url = _url
    data = '{"kraOrgNm": "' + name + '"}'
    response = requests.post(url, headers=headers, data=data.encode('utf-8'))
    data = json.loads(response.text)
    try:
        result = data["resultSVO"]["orgDVOList"][0]["orgCode"]
    except:
        result = "None"
    return result


for i in range(ws.nrows):
    name = ws.cell_value(i, 0)
    kind = ws.cell_value(i, 1)
    location = ws.cell_value(i,2)
    neis = code(name)

    worksheet.write(i, 0, neis)
    worksheet.write(i, 1, name)
    worksheet.write(i, 2, kind)
    worksheet.write(i, 3, location)
    print(i)

workbook.save(result_filename)
