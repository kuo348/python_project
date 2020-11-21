#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      cooper.huang
#
# Created:     19/11/2020
# Copyright:   (c) cooper.huang 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from os import system, name
import requests
import urllib.parse
from bs4 import BeautifulSoup
import xlwt
import pandas as pd
import re
import sys
__viewstate=""
__eventvaildation=""
__viewstategenerator=""
year=2020
outputfile=""
title="" #
subject="" #科目
level=""  #等級
category="" #
exam_name="" #考試名稱
exam_link=""   #題目連結
ans_link="" #答案連結
exam_list=[]
url="https://wwwq.moex.gov.tw/exam/"
def get_viewstate():
    r = requests.get("https://wwwq.moex.gov.tw/exam/wFrmExamQandASearch.aspx")

    if r.status_code== requests.codes.ok:
        soup = BeautifulSoup(r.text, 'html.parser')
        #print(r.text)
        #取出url_history的動態參數值
        """__viewstate = soup.find("input",{"id":"__VIEWSTATE","type":"hidden"})["value"]
        __viewstategenerator = soup.find("input",{"id":"__VIEWSTATEGENERATOR"})["value"]
        __eventvaildation = soup.find("input",{"id":"__EVENTVALIDATION"})["value"]"""
        hidden_tags = soup.find_all("input", type="hidden")
        #print(hidden_tags)
        for tag in hidden_tags:
            if tag.get('value') is not None:
                if tag.get('id')=="__VIEWSTATE":
                    __viewstate = tag.get('value')
                if tag.get('id')=="__VIEWSTATEGENERATOR":
                    __viewstategenerator = tag.get('value')
                if tag.get('id')=="__EVENTVALIDATION":
                    __eventvaildation = tag.get('value')


def query():
    #print(msg)
    headers = {
          #"Authorization": "Bearer " + token,c
          "Content-Type" : "application/x-www-form-urlencoded",
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.119 Safari/537.36"

    }
    data = {
          "__VIEWSTATE":"/wEPDwULLTExNjUxNDA4NzEPFggeClNlYXJjaE1vZGULKXZNb2V4UWFuZEFTZWFyY2gud0ZybUV4YW1RYW5kQVNlYXJjaCtTZWFyY2hNb2RlLCBNb2V4UWFuZEFTZWFyY2gsIFZlcnNpb249MS4wLjAuMCwgQ3VsdHVyZT1uZXV0cmFsLCBQdWJsaWNLZXlUb2tlbj1udWxsAB4MU2ltcGxlU2VhcmNoaB4OUXVlcnlMaW5xT3JkZXIFJkV4YW1TZXRDb2RlLCBDYXRlZ29yeUNvZGUsIFN1YmplY3RDb2RlHgdDaGtOYW1lBYsGY2hrXzEwOTA1MF8xMDRfMDEwMSxjaGtfMTA5MDUwXzEwNF8wMjAxLGNoa18xMDkwNTBfMTA0XzAzMDgsY2hrXzEwOTA1MF8xMDRfMTIwMixjaGtfMTA5MDUwXzEwNF8xMjA0LGNoa18xMDkwNTBfMTA0XzEyMDUsY2hrXzEwOTA1MF8xMDRfMTIwOSxjaGtfMTA5MDUwXzE0M18wMTAzLGNoa18xMDkwNTBfMTQzXzAyMDUsY2hrXzEwOTA1MF8xNDNfMDMwOSxjaGtfMTA5MDUwXzE0M18xMjExLGNoa18xMDkwNTBfMTQzXzEyMTMsY2hrXzEwOTA1MF8zMDhfMDEwMixjaGtfMTA5MDUwXzMwOF8wMzA4LGNoa18xMDkwNTBfMzA4XzEyMDMsY2hrXzEwOTA1MF8zMDhfMTIwNixjaGtfMTA5MDUwXzMwOF8xMjA3LGNoa18xMDkwNTBfMzA4XzEyMDgsY2hrXzEwOTA1MF8zMDhfMTIxMCxjaGtfMTA5MDkwXzM4NF8wMTAxLGNoa18xMDkwOTBfMzg0XzAyMTYsY2hrXzEwOTA5MF8zODRfMTUwNCxjaGtfMTA5MDkwXzM4NF8xNTA1LGNoa18xMDkwOTBfMzg0XzE1MDcsY2hrXzEwOTA5MF8zODRfMTUwOCxjaGtfMTA5MDkwXzM4NF8xNTA5LGNoa18xMDkwOTBfMzg0XzE1MTAsY2hrXzEwOTA5MF80NTJfMDEwMixjaGtfMTA5MDkwXzQ1Ml8wMjE3LGNoa18xMDkwOTBfNDUyXzE1MDIsY2hrXzEwOTA5MF80NTJfMTUwMyxjaGtfMTA5MDkwXzQ1Ml8xNTA2LGNoa18xMDkwOTBfNDUyXzE1MTEsY2hrXzEwOTE2MF8yMjJfMDEwMixjaGtfMTA5MTYwXzIyMl8wMTAzLGNoa18xMDkxNjBfMjIyXzEyMDEsY2hrXzEwOTE2MF8yMjJfMTIwMixjaGtfMTA5MTYwXzIyMl8xMjAzLGNoa18xMDkxNjBfMjIyXzEyMDQWAmYPZBYCAgMPZBYEAgUPZBYCAgEPFgIeBFRleHQFEuiAg+eVouippumhjOafpeipomQCBw9kFgICAQ9kFhICAQ9kFgJmDxAPFgYeDURhdGFUZXh0RmllbGQFDEV4YW1ZZWFyVGV4dB4ORGF0YVZhbHVlRmllbGQFDUV4YW1ZZWFyVmFsdWUeC18hRGF0YUJvdW5kZ2QQFR0DMTA5AzEwOAMxMDcDMTA2AzEwNQMxMDQDMTAzAzEwMgMxMDEDMTAwAjk5Ajk4Ajk3Ajk2Ajk1Ajk0AjkzAjkyAjkxAjkwAjg5Ajg4Ajg3Ajg2Ajg1Ajg0AjgzAjgyAjgxFR0EMjAyMAQyMDE5BDIwMTgEMjAxNwQyMDE2BDIwMTUEMjAxNAQyMDEzBDIwMTIEMjAxMQQyMDEwBDIwMDkEMjAwOAQyMDA3BDIwMDYEMjAwNQQyMDA0BDIwMDMEMjAwMgQyMDAxBDIwMDAEMTk5OQQxOTk4BDE5OTcEMTk5NgQxOTk1BDE5OTQEMTk5MwQxOTkyFCsDHWdnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnFgFmZAICD2QWAmYPEA8WBh8FBQxFeGFtWWVhclRleHQfBgUNRXhhbVllYXJWYWx1ZR8HZ2QQFR0DMTA5AzEwOAMxMDcDMTA2AzEwNQMxMDQDMTAzAzEwMgMxMDEDMTAwAjk5Ajk4Ajk3Ajk2Ajk1Ajk0AjkzAjkyAjkxAjkwAjg5Ajg4Ajg3Ajg2Ajg1Ajg0AjgzAjgyAjgxFR0EMjAyMAQyMDE5BDIwMTgEMjAxNwQyMDE2BDIwMTUEMjAxNAQyMDEzBDIwMTIEMjAxMQQyMDEwBDIwMDkEMjAwOAQyMDA3BDIwMDYEMjAwNQQyMDA0BDIwMDMEMjAwMgQyMDAxBDIwMDAEMTk5OQQxOTk4BDE5OTcEMTk5NgQxOTk1BDE5OTQEMTk5MwQxOTkyFCsDHWdnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnZ2dnFgFmZAIFDxAPFgYfBQUEdGV4dB8GBQV2YWx1ZR8HZ2QQFRQV5omA5pyJ6ICD6Kmm57Ch56ixLi4uTjEwOeW5tOitpuWvn+S6uuWToeWNh+WumOetieiAg+ippuOAgTEwOeW5tOS6pOmAmuS6i+alremQtei3r+S6uuWToeWNh+izh+iAg+ippiQxMDflubTmtojpmLLorablr5/nibnogIPph43mlrDogIPoqaaHATEwOeW5tOWFrOWLmeS6uuWToeeJueeoruiAg+ippuWPuOazleWumOiAg+ippu+8iOesrOS6jOippu+8ieOAgTEwOeW5tOWwiOmWgOiBt+alreWPiuaKgOihk+S6uuWToemrmOetieiAg+ippuW+i+W4q+iAg+ippu+8iOesrOS6jOippu+8iTMxMDnlubTlhazli5nkurrlk6Hpq5jnrYnogIPoqabkuIDntJrmmqjkuozntJrogIPoqaaWATEwOeW5tOWFrOWLmeS6uuWToeeJueeoruiAg+ippuWkluS6pOmgmOS6i+S6uuWToeWPiuWkluS6pOihjOaUv+S6uuWToeiAg+ippuOAgeWci+mam+e2k+a/n+WVhuWLmeS6uuWToeiAg+ippuOAgeawkeiIquS6uuWToeiAg+ippuOAgeWOn+S9j+awkeaXj+iAg+ippnIxMDnlubTlsIjploDogbfmpa3lj4rmioDooZPkurrlk6Hpq5jnrYnogIPoqabmnIPoqIjluKvjgIHkuI3li5XnlKLkvLDlg7nluKvjgIHlsIjliKnluKvjgIHmsJHplpPkuYvlhazorYnkurrogIPoqaa9ATEwOeW5tOWFrOWLmeS6uuWToeeJueeoruiAg+ippuWPuOazleS6uuWToeiAg+ippuOAgeazleWLmemDqOiqv+afpeWxgOiqv+afpeS6uuWToeiAg+ippuOAgeWci+WutuWuieWFqOWxgOWci+WutuWuieWFqOaDheWgseS6uuWToeiAg+ippuOAgea1t+WyuOW3oemYsuS6uuWToeiAg+ippuOAgeenu+awkeihjOaUv+S6uuWToeiAg+ippocBMTA55bm05YWs5YuZ5Lq65ZOh54m556iu6ICD6Kmm5Y+45rOV5a6Y6ICD6Kmm77yI56ys5LiA6Kmm77yJ44CBMTA55bm05bCI6ZaA6IG35qWt5Y+K5oqA6KGT5Lq65ZOh6auY562J6ICD6Kmm5b6L5bir6ICD6Kmm77yI56ys5LiA6Kmm77yJyQExMDnlubTnrKzkuozmrKHlsIjmioDkurrlk6Hpq5jnrYnogIPoqabkuK3phqvluKvogIPoqabliIbpmo7mrrXogIPoqabjgIHnh5/ppIrluKvjgIHlv4PnkIbluKvjgIHorbfnkIbluKvjgIHnpL7mnIPlt6XkvZzluKvogIPoqabjgIExMDnlubTlsIjmioDkurrlk6Hpq5jnrYnogIPoqabms5XphqvluKvjgIHoqp7oqIDmsrvnmYLluKvjgIHogb3lipsuLi7PATEwOeW5tOesrOS6jOasoeWwiOaKgOS6uuWToemrmOetieiAg+ippumGq+W4q+iAg+ippuWIhumajuauteiAg+ippu+8iOesrOS4gOmajuaute+8ieOAgeeJmemGq+W4q+iXpeW4q+iAg+ippuWIhumajuauteiAg+ippuOAgemGq+S6i+aqoumpl+W4q+OAgemGq+S6i+aUvuWwhOW4q+OAgeeJqeeQhuayu+eZguW4q+OAgeiBt+iDveayu+eZguW4q+OAgeWRvOWQuC4uLjkxMDnlubTlhazli5nkurrlk6Hpq5jnrYnogIPoqabkuInntJrogIPoqabmmqjmma7pgJrogIPoqaZaMTA55bm056ys5LqM5qyh5bCI5oqA5Lq65ZOh6auY562J6ICD6Kmm6Yar5bir6ICD6Kmm5YiG6ZqO5q616ICD6Kmm77yI56ys5LqM6ZqO5q616ICD6Kmm77yJbDEwOeW5tOWFrOWLmeS6uuWToeeJueeoruiAg+ippuitpuWvn+S6uuWToeiAg+ippuOAgeS4gOiIrOitpuWvn+S6uuWToeiAg+ippuOAgeS6pOmAmuS6i+alremQtei3r+S6uuWToeiAg+ipps8BMTA55bm05bCI5oqA5Lq65ZOh6auY562J6ICD6Kmm5aSn5Zyw5bel56iL5oqA5bir6ICD6Kmm5YiG6ZqO5q616ICD6Kmm77yI56ys5LiA6ZqO5q616ICD6Kmm77yJ44CB6amX6Ii55bir44CB56ys5LiA5qyh6aOf5ZOB5oqA5bir6ICD6Kmm44CB6auY562J5pqo5pmu6YCa6ICD6Kmm5raI6Ziy6Kit5YKZ5Lq65ZOh6ICD6Kmm44CB5pmu6YCa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b2xkZXJDb250ZW50JGNoa18xMDkwNTBfMTA0XzAyMDEFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA1MF8xMDRfMDMwOAUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MDUwXzEwNF8xMjAyBSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkwNTBfMTA0XzEyMDQFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA1MF8xMDRfMTIwNQUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MDUwXzEwNF8xMjA5BSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkwNTBfMTQzXzAxMDMFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA1MF8xNDNfMDIwNQUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MDUwXzE0M18wMzA5BSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkwNTBfMTQzXzEyMTEFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA1MF8xNDNfMTIxMwUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MDUwXzMwOF8wMTAyBSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkwNTBfMzA4XzAzMDgFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA1MF8zMDhfMTIwMwUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MDUwXzMwOF8xMjA2BSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkwNTBfMzA4XzEyMDcFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA1MF8zMDhfMTIwOAUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MDUwXzMwOF8xMjEwBSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkwOTBfMzg0XzAxMDEFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA5MF8zODRfMDIxNgUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MDkwXzM4NF8xNTA0BSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkwOTBfMzg0XzE1MDUFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA5MF8zODRfMTUwNwUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MDkwXzM4NF8xNTA4BSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkwOTBfMzg0XzE1MDkFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA5MF8zODRfMTUxMAUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MDkwXzQ1Ml8wMTAyBSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkwOTBfNDUyXzAyMTcFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA5MF80NTJfMTUwMgUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MDkwXzQ1Ml8xNTAzBSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkwOTBfNDUyXzE1MDYFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTA5MF80NTJfMTUxMQUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MTYwXzIyMl8wMTAyBSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkxNjBfMjIyXzAxMDMFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTE2MF8yMjJfMTIwMQUnY3RsMDAkaG9sZGVyQ29udGVudCRjaGtfMTA5MTYwXzIyMl8xMjAyBSdjdGwwMCRob2xkZXJDb250ZW50JGNoa18xMDkxNjBfMjIyXzEyMDMFJ2N0bDAwJGhvbGRlckNvbnRlbnQkY2hrXzEwOTE2MF8yMjJfMTIwNMirVT9U5T0OgtExS+jGMQpraFjR",
          "__EVENTVALIDATION":"/wEWiwECsYHJmwYCodKArQkCmuPntQ4CxOPuhwoC2fSQrwkC2fTswwEC2fS4qwMC2fSUzAsC2fTgYALZ9PyFCQLZ9Mi+DgLZ9KTTBgLZ9LD0DwLZ9IypBAKyzb6wAwKyzYrVCwKyzaa8DQKyzbLRBQKyzY6KCgKyzZqvAwKyzfbDCwKyzcJkArLN3pkJArLNqrIOAvauw8QLAvau33kC9q7rwAUC9q7H5QoC9q7TngMC9q6vswgC9q67VAL2rpeJCQL4s93PBwKms9T9AwK7pKpVAruk1rkIArukgtEKArukrrYCAruk2poJArukxn8Cu6TyxAcCu6SeqQ8Cu6SKjgYCu6S20w0C0J2EygoC0J2wrwIC0J2cxgQC0J2IqwwC0J208AMC0J2g1QoC0J3MuQIC0J34ngkC0J3kYwLQnZDIBwKU/vm+AgKU/uWDCQKU/tG6DAKU/v2fAwKU/unkCgKU/pXJAQKU/oGuCQKU/q1zArbFi94MAtXj7zEChNDsgAsCh9CA5AICgdDogAsChdDsgAsCmtDsgAsCm9DsgAsCgNDsgAsCgdDsgAsChtDsgAsCh9DsgAsCntD4vQQCn9D4vQQChND4vQQChdD4vQQCmtD4vQQCm9D4vQQCgND4vQQCgdD4vQQChtD4vQQCr67tpAIC3NfnpQQCqtnW6Q0Cl4mHQQKH5q2vDAKY5q2vDAKZ5q2vDALL4f+MCgKYiYdBApfmra8MApbmra8MAsvhi40KApfmra8MAsbFr/YFAumwv7YOAoOppfoPArD839sJAp/LpvAMAp7LpvAMApD+sq8IAoPi9JoDAs2PufAOArKmm9sEAvaUgdoOAoX43sQIAs6lo5oEApGUiZkOArrKiq8NAoT4zoQJAoTiwOMBApD+6rcHAuj4jo4IApe9tc4JAvzTl7kPApH+1vcHArm05M0GArvnzfULArPZvIsJAuar4PUNAsvCwuADApXwhrYPAqqaxvQHAo+xqN8NAtLQ28oGAr39u60BArPvqsMOAsH9z20CppSy2AYC1djYmAgC3ObxggsCxu+KwgsC4dio1wUClKvQwQoCw+/2gQwC3tiUlwYCjZ271wfMG69Gv8waOCJWGjcNu0cB7Ps8tg==",
          #"__VIEWSTATE":__viewstate,
          #"__EVENTVALIDATION":__eventvaildation,
          #"__VIEWSTATEGENERATOR":__viewstategenerator,
          "ctl00$holderContent$hidStatus": "",
          #"ctl00$holderContent$txtExamSetName": "關務人員考試",
          #"ctl00$holderContent$txtSubjectName": "請輸入科目關鍵字...",
          "ctl00$holderContent$wUctlExamYearStart$ddlExamYear": year ,
          "ctl00$holderContent$wUctlExamYearEnd$ddlExamYear": year  ,
          "ctl00$holderContent$ddlExamCode" :"",
          "ctl00$holderContent$txtCategoryName": "資訊處理",
          "ctl00$holderContent$btnQuery": "查詢",
          "ctl00$holderContent$wUctlExamDisplayModeForSearch$ddlExamDisplayMode1": 0,
          "ctl00$holderContent$wUctlExamDisplayModeForSearch$ddlExamDisplayMode2": 1,
          "ctl00$holderContent$wUctlExamDisplayModeForSearch$ddlExamDisplayMode3": 2}
    payload={}
    """payload = { "ctl00%24holderContent%24hidStatus":"",
                "ctl00%24holderContent%24wUctlExamYearStart%24ddlExamYear": 2020,
                "ctl00%24holderContent%24wUctlExamYearEnd%24ddlExamYear": 2020,
                "ctl00%24holderContent%24ddlExamCode":"",
                "ctl00%24holderContent%24txtExamSetName":"%E8%AB%8B%E8%BC%B8%E5%85%A5%E8%80%83%E8%A9%A6%E5%90%8D%E7%A8%B1%E9%97%9C%E9%8D%B5%E5%AD%97...",
                "ctl00%24holderContent%24txtCategoryName": "%E8%B3%87%E8%A8%8A%E8%99%95%E7%90%86",
                "ctl00%24holderContent%24txtSubjectName": "%E8%AB%8B%E8%BC%B8%E5%85%A5%E7%A7%91%E7%9B%AE%E9%97%9C%E9%8D%B5%E5%AD%97...",
                "ctl00%24holderContent%24wUctlExamDisplayModeForSearch%24ddlExamDisplayMode1": 0,
                "ctl00%24holderContent%24wUctlExamDisplayModeForSearch%24ddlExamDisplayMode2": 1,
                "ctl00%24holderContent%24wUctlExamDisplayModeForSearch%24ddlExamDisplayMode3": 2,
                "ctl00%24holderContent%24btnQuery": "%E6%9F%A5%E8%A9%A2"
                #"__EVENTVALIDATION": "%2FwEWiwECsYHJmwYCodKArQkCmuPntQ4CxOPuhwoC2fSQrwkC2fTswwEC2fS4qwMC2fSUzAsC2fTgYALZ9PyFCQLZ9Mi%2BDgLZ9KTTBgLZ9LD0DwLZ9IypBAKyzb6wAwKyzYrVCwKyzaa8DQKyzbLRBQKyzY6KCgKyzZqvAwKyzfbDCwKyzcJkArLN3pkJArLNqrIOAvauw8QLAvau33kC9q7rwAUC9q7H5QoC9q7TngMC9q6vswgC9q67VAL2rpeJCQL4s93PBwKms9T9AwK7pKpVAruk1rkIArukgtEKArukrrYCAruk2poJArukxn8Cu6TyxAcCu6SeqQ8Cu6SKjgYCu6S20w0C0J2EygoC0J2wrwIC0J2cxgQC0J2IqwwC0J208AMC0J2g1QoC0J3MuQIC0J34ngkC0J3kYwLQnZDIBwKU%2Fvm%2BAgKU%2FuWDCQKU%2FtG6DAKU%2Fv2fAwKU%2FunkCgKU%2FpXJAQKU%2FoGuCQKU%2Fq1zArbFi94MAtXj7zEChNDsgAsCh9CA5AICgdDogAsChdDsgAsCmtDsgAsCm9DsgAsCgNDsgAsCgdDsgAsChtDsgAsCh9DsgAsCntD4vQQCn9D4vQQChND4vQQChdD4vQQCmtD4vQQCm9D4vQQCgND4vQQCgdD4vQQChtD4vQQCr67tpAIC3NfnpQQCqtnW6Q0Cl4mHQQKH5q2vDAKY5q2vDAKZ5q2vDALL4f%2BMCgKYiYdBApfmra8MApbmra8MAsvhi40KApfmra8MAsbFr%2FYFAumwv7YOAoOppfoPArD839sJAp%2FLpvAMAp7LpvAMApD%2Bsq8IAoPi9JoDAs2PufAOArKmm9sEAvaUgdoOAoX43sQIAs6lo5oEApGUiZkOArrKiq8NAoT4zoQJAoTiwOMBApD%2B6rcHAuj4jo4IApe9tc4JAvzTl7kPApH%2B1vcHArm05M0GArvnzfULArPZvIsJAuar4PUNAsvCwuADApXwhrYPAqqaxvQHAo%2BxqN8NAtLQ28oGAr39u60BArPvqsMOAsH9z20CppSy2AYC1djYmAgC3ObxggsCxu%2BKwgsC4dio1wUClKvQwQoCw%2B%2F2gQwC3tiUlwYCjZ271wfMG69Gv8waOCJWGjcNu0cB7Ps8tg%3D%3D"
                }"""
    #p= urllib.parse.quote(payload)
    r = requests.post("https://wwwq.moex.gov.tw/exam/wFrmExamQandASearch.aspx", headers = headers,data=data)
    return r

def printTable(t):
    #print(t)
    global subject,exam_name,level,exam_link,ans_link,url,category,title,year
    global exam_list
    for row in t.find_all('tr',recursive=False) :
        i=0
        for col in row.find_all('td',recursive=False):
            #print(i)
            i=i+1
            if col.find('table')!=None :
                if col.get('class')!=None  and col['class'][0]=="level1":
                    if col.text!="": 
                        #exams = col.text.replace('\n\n','')
                        exam_name =col.text.replace('\n\n','')
                        #print(exam_name)
                        subject=""
                        exam_link=""
                        ans_link=""
                        level =""
                else :
                    if col.get('class')!=None and col['class'][0]=="level2":
                        print(col.text)
                        m = re.match('(^[\u4e00-\u9fa5]+)_([\u4e00-\u9fa5()]+)$',col.text.replace('\n\n',''))
                        #print(m.group(1))
                        #print(m.group(2))
                        if m!=None :
                            category=m.group(2)
                            level=m.group(1)
                        else :
                            category=""
                            l
                        print(level)
                        if level.find("考試")>0 :
                            m = re.match('(^[^考試$]+考試)|',level).group(1)
                            title=m
                            if m!=None:
                                level=level[len(m):]
                                title=""
                            if level=="" :
                                level=m

                    else:
                        printTable(col.find('table',recursive=False))
            else :
                label=col.find('label',recursive=False)
                links=col.find_all('a',recursive=False)
                if label!= None :
                    subject = label.text
                    #print(subject)
                if links!= None :
                    for link in links :
                        if link.get('class')!=None and link['class'][0]=="exam-question-ans":
                            if link.text=="試題":
                              exam_link= url+link.get('href')
                            else:
                                ans_link= url+link.get('href')
                            #print(ans_link)
    if exam_link!="":
        data=[]
        exams= exam_name.split('、')
        for e in exams :
            if title!=None and e.find(title) > 0:
                exam_name=e
                exit
        if exam_name.find('本考試所有測驗題標準答案')>0:
            exam_name=exam_name.replace('本考試所有測驗題標準答案','')
        data.append(exam_name)       
        data.append(level)
        data.append(subject)
        data.append(exam_link) 
        data.append(ans_link)
        exam_list.append(data)
                
                    




def save(filename,txt):
    #print(tbs)
    #print(content)
    # 開啟檔案
    fp = open(filename, "w+",encoding="utf-8")
    #fp.writelines('開啟檔案')
    fp.writelines(txt)
    fp.close()


def main():
    get_viewstate()
    r = query()
    #print(r.status_code)
    if r.status_code== requests.codes.ok:
        #print(r.text)
        #save("d:\\1.html",r.text)
        soup = BeautifulSoup(r.text, 'html.parser')
        t = soup.find("table",{"id":"ctl00_holderContent_tblExamQand"})
        #t = soup.find_all("table")
        printTable(t)
        #print(t)
    #print("viewstate="+__viewstate)
    #print("eventvaildation="+__eventvaildation)
if __name__ == '__main__':
    if len(sys.argv)>=2:
        year=sys.argv[1]
        outputfile=sys.argv[2]
    #print(year)
    #print(outputfile)
    exam_link=""
    main()
    df = pd.DataFrame(data=exam_list, columns=['examname','level','subject','link1','link2'])
    #print(df)
    df.to_excel(outputfile,na_rep=False)
    print('output complete....')
    #^[\u4e00-\u9fa5]+
