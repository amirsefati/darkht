import schedule
import time
import requests
import json
from datetime import datetime
import xlsxwriter

workbook = xlsxwriter.Workbook('report.xlsx')
worksheet = workbook.add_worksheet()

header_mehr = {
    'Host': 'mehr.exirbroker.com',
    'Connection': 'keep-alive',
    'Content-Length': '544',
    'Accept': 'application/json, text/plain, */*',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36',
    'Content-Type': 'application/json',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Dest': 'empty',
    'Referer': 'https://mehr.exirbroker.com/mainNew',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'en-GB,en;q=0.9,fa-IR;q=0.8,fa;q=0.7,en-US;q=0.6',
    'Cookie': 'PLAY_LANG=fa; cookiesession1=0B4E32E6KB5IQECHY0UHNWWTNXMB6BBF; PLAY_SESSION=f920564f4c96f33be5f0446764f790f5fe85e051-client_login_id=7f6c3e18b72749d9ba36d720b1cf92e5&client_id=122c3661eb8144e2965216355ec3c06a&authToken=4d51ed68fe6c414a8f7042d1adfec1d0'
    }
pyload = {"id":"","version":1,"hon":"","bankAccountId":-1,"insMaxLcode":"IRO1SIPA0001","abbreviation":"","latinAbbreviation":"","side":"SIDE_BUY","quantity":1,"quantityStr":"","remainingQuantity":0,"price":1,"priceStr":"","tradedQuantity":0,"averageTradedPrice":0,"disclosedQuantity":0,"orderType":"ORDER_TYPE_LIMIT","validityType":"VALIDITY_TYPE_DAY","validityDate":"","validityDateHidden":"hidden","orderStatusId":0,"queueIndex":-1,"searchedWord":"","coreType":"c","marketType":"","hasUnderCautionAgreement":"false","dividedOrder":"false","clientUUID":""}

#فقط این دو مقدار را تغییر دهید
#############

delay = .1

def job():
    init = 1
    final = 5
    data = []
    while init < final :
        report_start = datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        url = "https://mehr.exirbroker.com/api/v1/order"
        res = requests.post(url,headers=header_mehr,data=json.dumps(pyload))
        report_end = datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
        report_data = res.text
        init += 1
        data.append([report_start,report_end,report_data])
        time.sleep(delay)
    else:
        row = 1
        worksheet.write(0,1,'send request')
        worksheet.write(0,2,'get resopnse')
        worksheet.write(0,3,'data')
        for item in data:
            worksheet.write(row,1,item[0])
            worksheet.write(row,2,item[1])
            worksheet.write(row,3,item[2])
            row += 1
        workbook.close()

schedule.every().day.at("22:25:00").do(job)

while True:
    schedule.run_pending()
    time.sleep(0.0001)