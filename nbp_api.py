from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import json
import csv
from datetime import date

def run_chrome(url):
    driver = webdriver.Chrome()
    driver.get(url)
    elem = driver.find_element(By.TAG_NAME, 'body').text
    driver.close()
    parse_json(elem)
    
def parse_json(data):
    nbp_json = json.loads(data)
    nbp_json = nbp_json['rates']
    nbp_data = []
    offset = pd.tseries.offsets.BusinessDay(n=1)
    for i in range(len(nbp_json)):
        exchange_date = nbp_json[i]['effectiveDate']
        ts = pd.Timestamp(str(exchange_date))
        business_date = ts - offset
        business_date = business_date.strftime('%Y-%m-%d')
        raw_data = {
            'data_kursu' : exchange_date,
            'wartosc_kursu' : nbp_json[i]['mid'],
            'ostatni_dzien_roboczy' : business_date
            }
        nbp_data.append(raw_data)
    save_file(nbp_data) 

def save_file(file):
    fields = ['data_kursu', 'wartosc_kursu', 'ostatni_dzien_roboczy']
    filename = 'nbp_kurs.csv'
    with open(filename, 'w', encoding='utf8', newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames = fields, delimiter = ';')
        writer.writeheader()
        writer.writerows(file)
    
if __name__ == "__main__":
    today = date.today()
    d1 = today.strftime('%Y-%m-%d')
    url = 'http://api.nbp.pl/api/exchangerates/rates/a/EUR/2022-10-01/'+d1+'?format=json'
    run_chrome(url)
