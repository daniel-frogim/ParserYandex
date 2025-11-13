from bs4 import BeautifulSoup
from selenium import webdriver
import json
import openpyxl


with open('data/input.json', 'r', encoding='utf-8') as file:
    input_data = json.load(file)


city = input_data['Geoposition']
city = city.replace(" ", "%20")
Mtype = input_data['Type']

url = "https://yandex.ru/maps/1/moscow-and-moscow-oblast/search/"
url += city + "%20" + Mtype
driver = webdriver.Chrome()

driver.get(url)

check = input("Введи Enter когда долистаешь")


html = driver.page_source
bs = BeautifulSoup(html,"lxml")


ul_li = bs.find_all('div', 'search-business-snippet-view__content')

data = {
    "name": [],
    "rating": [],
    "href": []
    }
for i in range(len(ul_li)):
    EAX = ul_li[i].find('div', 'search-business-snippet-view__title')
    EDX = ul_li[i].find('a', 'search-business-snippet-view__rating')
    if EAX is not None and EDX is not None:
        data['name'].append(EAX.text)
        data['rating'].append(EDX.text)
        data['href'].append('https://yandex.ru' + EDX.get('href'))


with open('data/output.json', 'w', encoding='utf-8') as file:
    json.dump(data, file, indent=4, ensure_ascii=False)


workbook = openpyxl.Workbook()
sheet = workbook.active

count = len(ul_li)
if len(data['rating']) < count:
    count = len(data['rating'])


for i in range(count):
    rating_split = data['rating'][i].split(" ")
    sheet['A' + str(i + 1)] = data['name'][i]
    sheet['B' + str(i + 1)] = city.replace('%20', ' ')
    sheet['C' + str(i + 1)] = rating_split[0]
    sheet['D' + str(i + 1)] = rating_split[1]
    sheet['E' + str(i + 1)] = data['href'][i]


workbook.save("Общее/" + city.replace('%20', ' ') + " " + Mtype + ".xlsx")

print("\n\nВСЕ СДЕЛАНО\n\n")