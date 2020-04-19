import json
import xlrd
import xlwt
import requests
import urllib.parse

loc = "chungnam"
_url = "https://stu.cne.go.kr/spr_ccm_cm01_100.ws"
cookie = "Cookie: WMONID=s4YE1Z8QdvU; JSESSIONID=OTciadGGS8KeoKoJNiW6ocUbX4Nn49xw0VCEU8s5tldP01OR4j2KQjZ6650Ykclz.cne-pacwas2_servlet_stuwas; Veraport20Use=N"


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
