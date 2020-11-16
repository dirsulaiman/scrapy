# library Selenium
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time

# library Beautifulsoup
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup as bs
from tqdm import tqdm
import re
import pandas as pd
import xlsxwriter

import os
import sys
import dotenv
from pathlib import Path

def check_yesorno (value):
    if value.upper() == "YES":
        return True
    return False 

url = 'https://www.olx.co.id/jakarta-dki_g2000007/mobil-bekas_c198'
output_file_name = "output"
page_limit = 2
count_all = False
multiple_scrap = False
app_loop = 1

try:
    env_path = Path(".") / (".env")
    dotenv.load_dotenv(dotenv_path = env_path)
    output_file_name = os.getenv("OUTPUT_FILENAME")
    page_limit = int(os.getenv("PAGE_SCRAP"))
    count_all = check_yesorno(os.getenv("LOAD_ALL"))
    multiple_scrap = check_yesorno(os.getenv("MULTIPLE_SCRAP"))
    app_loop = int(os.getenv("APP_LOOP"))
    print("Loaded .env file")
except Exception:
    print("An rror occurred when read .env file")
    print("Load default configuration")

# pattern for regex search
pattern_link = re.compile("href=[\"\'](.+?)[\"\']")
pattern_phone = re.compile("\"phone\",\"value\":\"(.*?)\"")
pattern_name = re.compile(":{\"id\":\"\d*\",\"name\":\"(.*?)\"")
pattern_about = re.compile("\":true,\"about\":\"(.*?)\"")
pattern_deallertype = re.compile("-type-diler\",\"value_name\":\"(.*?)\"")
pattern_brand = re.compile("\d\",\"brand\":\"(.*?)\"")
pattern_model = re.compile("\"model\":\"(.*?)\",\"modelDate") #0
pattern_mileage = re.compile("\"mileageFromOdometer\":{\"@type\":\"QuantitativeValue\",\"value\":\"(.*?)\"")
pattern_unitcode = re.compile("unitCode\":\"(.*?)\"") #0
pattern_fueltype = re.compile("fuelType\":\"(.*?)\"")
pattern_color = re.compile("fuelType\":\"[\w\s]*\",\"color\":\"(.*?)\"")
pattern_bodytype = re.compile("bodyType\":\"(.*?)\"")
pattern_transmission = re.compile("vehicleTransmission\":\"(.*?)\"")
pattern_enggine = re.compile(r"engineDisplacement\":{\"@type\":\"QuantitativeValue\",\"value\":\"[\\]*(\w){0,5}(.*?)\"")
pattern_int = re.compile(r"[^\d]")


# open new url and scrap profile detail
def add_detail (data):
    req = Request(data["link_detail"])            
    try:
        uopen = urlopen(req)
    except Exception:
        return
    page = uopen.read()
    str_page = str(page)
    phone = pattern_phone.findall(str_page)
    name = pattern_name.findall(str_page)
    about = pattern_about.findall(str_page)
    deallertype = pattern_deallertype.findall(str_page)
    brand = pattern_brand.findall(str_page)
    model = pattern_model.findall(str_page)
    mileage = pattern_mileage.findall(str_page)
    unitcode = pattern_unitcode.findall(str_page)
    fueltype = pattern_fueltype.findall(str_page)
    color = pattern_color.findall(str_page)
    bodytype = pattern_bodytype.findall(str_page)
    transmission = pattern_transmission.findall(str_page)
    enggine = pattern_enggine.findall(str_page)
    if brand:
        data["brand"] = brand[0]
    if model:
        data["model"] = model[0]
    if color:
        data["color"] = color[0]
    if bodytype:
        data["bodytype"] = bodytype[0]
    if transmission:
        data["transmission"] = transmission[0]
    if mileage:
        data["mileage"] = f"{mileage[0]} {unitcode[0]}"
    if fueltype:
        data["fueltype"] = fueltype[0]
    if enggine:
        try:
            data["enggine"] = enggine[0][1]
        except Exception:
            data["enggine"] = ""
    if phone:
        data["phone"] = phone[0]
    if name:
        data["name"] = name[0]
    if deallertype:
        data["deallertype"] = deallertype[0]
    if about:
        data["about"] = about[0]
    return data


def write_data_to_excel(data, file_name=output_file_name):
    df = pd.DataFrame(data)
    df = df.rename(columns={
        "title":"Judul", 
        "year":"Tahun", 
        "price":"Harga", 
        "brand": "Merek", 
        "model": "Model", 
        "mileage": "Jarak Tempuh", 
        "fueltype": "Tipe Bahan Bakar", 
        "color": "Warna", 
        "bodytype": "Tipe Bodi", 
        "transmission": "Transmisi",
        "enggine": "Kapasitas Mesin",
        "location":"Lokasi", 
        "link_detail":"url", 
        "phone":"Phone", 
        "name":"Nama", 
        "deallertype": "Tipe Penjual", 
        "about":"Info"})
    file_number = 0
    try:
        file_number = file_number + 1
        df.to_excel(f'{file_name}.xlsx', engine='xlsxwriter', index=False)
        print(f"\nComplete: write {len(data)} records to {file_name}.xlsx")
    except xlsxwriter.exceptions.FileCreateError:
        df.to_excel(f'{file_name}({file_number}).xlsx', engine='xlsxwriter', index=False)
        print(f"\nComplete: write {len(data)} records to {file_name}({file_number}).xlsx")
    except Exception as e:
        print(e)


def run_app(i, data):
    # Set Selenium webdriver
    print("Opening the web browser, do not disturb the browser and wait for the browser to close itself!")
    browser = webdriver.Chrome()
    browser.get(url)
    browser.implicitly_wait(3)
    
    page = ""
    cards = []
    page_count = 0

    # loop to Click "load more" with Selenium
    while True :
        if (not count_all) and (page_count >= page_limit):
            print("Load page complete")
            browser.close()
            break
        elif page_limit == 1:
            time.sleep(3)
            page = browser.page_source
            soup = bs(page, 'html.parser')
            cards = soup.find_all('li', class_="EIR5N")
            print("Finished loading page")
            print(f"Result: {len(cards)} products found")
            browser.close()
            break
        else:
            try:
                time.sleep(1)
                browser.find_element_by_xpath("//*[@data-aut-id='btnLoadMore']").click()
                page = browser.page_source
                page_count = page_count + 1
                soup = bs(page, 'html.parser')
                cards = soup.find_all('li', class_="EIR5N")
            except NoSuchElementException:
                page = browser.page_source
                # Parse html to BeautifulSoup
                soup = bs(page, 'html.parser')
                cards = soup.find_all('li', class_="EIR5N")
                print("Finished loading page")
                print(f"Result: {len(cards)} products found")
                browser.close()
                break
            except Exception as e:
                page = browser.page_source
                # Parse html to BeautifulSoup
                soup = bs(page, 'html.parser')
                cards = soup.find_all('li', class_="EIR5N")
                print("Something wrong")
                print(e)
                browser.close()
                break

    print(f"\nStart scraping {len(cards)} products from {url}")
    for card in tqdm(cards):
        row = {
            "title": "", 
            "year": "", 
            "price": "", 
            "brand": "", 
            "model": "", 
            "mileage": "", 
            "fueltype": "", 
            "color": "", 
            "bodytype": "", 
            "transmission": "",
            "enggine": " ",
            "location":"", 
            "link_detail":"", 
            "phone":"", 
            "name":"", 
            "deallertype": "", 
            "about":""}
        title = card.find('span', class_='_2tW1I')
        year = card.find('span', class_='_2TVI3')
        location = card.find('span', class_='tjgMj')
        price = card.find('span', class_='_89yzn')
        if title:
            row["title"] = title.text
        if year:
            row["year"] = int(pattern_int.sub('', year.text))
        if location:
            row["location"] = location.text
        if price:
            row["price"] = int(pattern_int.sub('', price.text))
        string = str(card.find('a'))
        link = pattern_link.findall(string)
        if link:
            link = "https://www.olx.co.id" + str(link[0])
            row["link_detail"] = link
            add_detail(row)
        data.append(row)
    
    file_name = f"{output_file_name}"
    if i >= 0:
        file_name = f"{output_file_name}({i})"
    write_data_to_excel(data, file_name=file_name)
    return data



if multiple_scrap:
    data_result = []
    for i in range(1, app_loop+1):
        print(f"\nRunning App {i}")
        data = []
        data = run_app(i, data)
        data_result = data_result + data
        time.sleep(1)
    write_data_to_excel(data_result, file_name=f"Result-{output_file_name}")
else:
    data = []
    run_app(-1, data)

os.system('pause')  
print("You can close this window now")
# sys.exit()
