import json
import xlrd
import xlwt
import requests
import urllib.parse

excel_filename = 'gyeonggi.xlsx'
result_filename = 'after_gyeonggi.xlsx'
wb = xlrd.open_workbook(filename = excel_filename)
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('result')

ws = wb.sheet_by_name('Sheet1')

def code(name):
    headers = {'Content-Type': 'application/json; charset=utf-8', 'Cookie':'Cookie: WMONID=QRzuzNBmwvr; Veraport20Use=N; insttNm=%uC6D0%uACE1%uC911%uD559%uAD50; schulCode=J100002709; schulKndScCode=03; schulCrseScCode=3; JSESSIONID=XXrAIN8L60zDdGba7Y64X8iptnO9UaMX0IHGHctrJXfu7PzC4lEpa6otmZq2uFSq.goe-pacwas2_servlet_stuwas'}
    url = "https://stu.goe.go.kr/spr_ccm_cm01_100.ws"
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
