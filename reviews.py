from bs4 import BeautifulSoup
from selenium import webdriver
import json
import openpyxl


with open('data/input.json', 'r', encoding='utf-8') as file:
    input_data = json.load(file)


url = input_data['reviews']

driver = webdriver.Chrome()
driver.get(url)

check = input("Введи Enter когда долистаешь")


html = driver.page_source
bs = BeautifulSoup(html,"lxml")

Name = bs.find('h1', 'orgpage-header-view__header')
reviews = bs.find_all('div', 'business-reviews-card-view__review')

data = {
    "author": [],
    "text": [],
    "date": [],
    "rating": []
    }


for i in range(len(reviews)):
    author = reviews[i].find('div', 'business-review-view__author-name')
    text = reviews[i].find('span', 'spoiler-view__text-container')
    date = reviews[i].find('span', 'business-review-view__date')
    rating = reviews[i].find('meta', itemprop='ratingValue')
    data['author'].append(author.text)
    data['text'].append(text.text)
    data['date'].append(date.text)
    data['rating'].append(rating.get('content'))


with open('data/output_reviews.json', 'w', encoding='utf-8') as file:
    json.dump(data, file, indent=4, ensure_ascii=False)


workbook = openpyxl.Workbook()
sheet = workbook.active

for i in range(len(reviews)):
    rating_split = data['rating'][i].split(" ")
    sheet['A' + str(i + 1)] = data['author'][i]
    sheet['B' + str(i + 1)] = data['rating'][i]
    sheet['C' + str(i + 1)] = data['date'][i]
    sheet['D' + str(i + 1)] = data['text'][i]

workbook.save("Отзывы/" + Name.text + " reviews.xlsx")



print("\n\nВСЕ СДЕЛАНО\n\n")
