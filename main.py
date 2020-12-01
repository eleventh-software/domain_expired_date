import whois
from openpyxl import *


def sorgula(domain):
    try:
        tarih = whois.whois(domain)
    except:
        print("Olmadı")
    return tarih.expiration_date


book = load_workbook("demo-icerik.xlsx")
sheet = book.active

i = 1

while True:

    if sheet['A' + str(i)].value:
        expired_date = sorgula(sheet['A' + str(i)].value)
        sheet['B' + str(i)] = str(expired_date)
        
        if type(expired_date) == list:
            sheet['B' + str(i)] = expired_date[0]
        
        i += 1

    else:
        break

book.save("yeni.xlsx")
book.close()
