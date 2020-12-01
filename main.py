import whois
from  openpyxl import *
 
# sorgula fonksiyonu parametre olarak verilen domaine sorgu atar 
# ve geriye domain son kullanma tarihini döndürür
def sorgula(domain):
    try:
        tarih = whois.whois(domain)
    except:
        print("Olmadı")
    return tarih.expiration_date

# Burda excel dosyasını açıp aktif ettik
book = load_workbook("test.xlsx")
sheet=book.active

i = 1
# Sonsuz bir döngü açıyoruz
while(True):
    # eğer A sütunu boş değilse işlem yap
    if sheet['A' + str(i)].value:
        # sorgula fonksiyonumuzu çalıştırıyoruz ve domain olarak excelden aldığımız veriyi veriyoruz
        # sorgunun sonucunu da expired_date isimli bir değişkene atıyoruz
        expired_date = sorgula(sheet['A' + str(i)].value)

        
        # whois kütüphanesi bazı domainlerde çalışmıyor. Bu tip durumlarda başka bir şey kullanmamız lazım
        # expired_date None ise alternatif yöntem ile expired_date i elde etmemiz lazım. Şu an bu yöntem ne bilmiyorum
        # sen halledersin :)
        if expired_date == None:
            pass

        # expired_date değerini excel dosyasında B sütununa yazıyoruz
        sheet['B' + str(i)] = str(expired_date)

        #domainleri gezmek için i değerini her seferinde 1 arttırıyoruz
        i += 1

    # eğer A sütunu boşsa döngüyü bitir
    else:
        break
        
book.save("yeni.xlsx")
book.close()