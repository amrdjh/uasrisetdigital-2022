# Kelompok 3
# Amir Musa                         (2006592032)
# Arifah MsRobin Emhas	            (2006592272)
# Nur Rahma Halimah Hadi            (2006592045)
# Nyak Cut Nadira 	                (2006522240)
# Rasnur Ellyani 		            (2006532840)
# Wulan Steviani Putri              (2006591912)


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import time

opsi = webdriver.ChromeOptions()
opsi.add_argument('--headless')
servis = Service('chromedriver.exe')
driver = webdriver.Chrome(service=servis, options=opsi)

shopee_link = "https://shopee.co.id/search?keyword=apple%20airpods&trackingId=searchhint-1671972405-3064bba7-8452-11ed-9133-f4ee081030ef"
driver.set_window_size(1300,800)
driver.get(shopee_link)
time.sleep(5)

rentang = 500
for i in range(1,7):
    akhir = rentang * i
    perintah="window.scrollTo(0,"+str(akhir)+")"
    driver.execute_script(perintah)
    print("loading ke-"+str(i))
    time.sleep(1)

driver.save_screenshot("homeshopeeappleairpods.png")
content = driver.page_source
driver.quit()

data = BeautifulSoup(content, 'html.parser')


i = 1
base_url = "https://shopee.co.id"

list_namaproduk,list_link,list_harga,list_terjual,list_kota=[],[],[],[],[]

for area in data.find_all('div',class_="col-xs-2-4 shopee-search-item-result__item"):
    print('proses data ke-'+str(i))
    namaproduk = area.find('div',class_="ie3A+n bM+7UW Cve6sh").get_text()
    link = base_url + area.find('a')['href']
    harga = area.find('span',class_="ZEgDH9").get_text()
    kota = area.find('div',class_="zGGwiV").get_text()
    terjual = area.find('div',class_="ZnrnMl").get_text()

    list_namaproduk.append(namaproduk)
    list_link.append(link)
    list_harga.append(harga)
    list_kota.append(kota)
    list_terjual.append(terjual)
    i+=1
    print("------")

df = pd.DataFrame({'A':list_namaproduk,'B':list_link,'C':list_harga,'D':list_kota,'E':list_terjual})
writer = pd.ExcelWriter('appleairpods1.xlsx')
df.to_excel(writer, 'Sheet1',index=False)
writer.save()
