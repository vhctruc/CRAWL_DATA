from selenium import webdriver
from selenium.common.exceptions import *
from openpyxl.workbook import Workbook
import sqlalchemy as sa
import urllib
import pandas as pd
import pyodbc
import time
import glob
from datetime import datetime, timedelta
from os import listdir
from os.path import isfile, join
import urllib.parse
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains



pd.set_option('display.max_columns',40)

options = webdriver.ChromeOptions()
options.add_argument('start-maximized')
options.add_argument('disable-infobars')
options.add_argument('--disable-extensions')
options.add_argument("--disable-gpu")
# options.add_argument("--headless")


driver = webdriver.Chrome(r'D:\chromedriver_win32\chromedriver.exe',options=options)
url = "https://www.nhathuocankhang.com/"
driver.get(url)
driver.implicitly_wait(0.5)
df = pd.DataFrame()
CATE_ = []
NAME_SUB_CATE_ = []
PRODUCT_ = []
RANK_ = []
PRODUCT_PRICES_ = []
PRODUCT_DETAILS_ =[]
CATE = driver.find_elements_by_xpath("//div[@class='cate-menu']//ul//li[1]//a")
cate = [i.get_attribute('href') for i in CATE]
for i_cate in range(len(cate)):
    driver.get(cate[i_cate])
    driver.implicitly_wait(5)
    SUB_CATE = driver.find_elements_by_xpath("//div[@class='group-medic']//ul//li//a")
    sub_cate = [i.get_attribute('href') for i in SUB_CATE]
    NAME_SUB_CATE = driver.find_elements_by_xpath("//div[@class='group-medic']//ul//li//a/h3[1]")
    name_sub_cate = [i.text for i in SUB_CATE]
    print(cate[i_cate])
    if (cate[i_cate] == "https://www.nhathuocankhang.com/dung-cu-y-te") or (cate[i_cate] == "https://www.nhathuocankhang.com/my-pham"):
        driver.get(cate[i_cate] + "#&o=7&pi=20")
        driver.implicitly_wait(5)
        print(driver.get(cate[i_cate] + "#&o=7&pi=20"))
        time.sleep(15)
        LINK_PRODUCT = driver.find_elements_by_xpath("//ul[@class = 'listing-prod']//li//a[@href]")
        link_product = [i.get_attribute('href') for i in LINK_PRODUCT]
        print(link_product)
        for i_link_product in range(len(link_product)):
            driver.get(link_product[i_link_product])
            driver.implicitly_wait(5)
            try:
                PRODUCT_PRICES = driver.find_elements_by_xpath("//div[@class='box-price']//span")
                product_prices = [i.text for i in PRODUCT_PRICES]
            except NoSuchElementException:
                pass
            try:
                PRODUCT = driver.find_element_by_xpath("//div[@class = 'detail-title']//h3").text
            except NoSuchElementException:
                pass
            try:
                action = driver.find_element_by_xpath("//div//div[@class='detail-wrapper normal']//a[@class = 'btn-seemore']//span").click()
            except:
                pass
            time.sleep(2)
            try:
                PRODUCT_DETAILS = driver.find_elements_by_xpath("//div//div[@class='text-info']//div[@class='box-textdt']")
                product_details = [i.text for i in PRODUCT_DETAILS]
            except:
                try:
                    PRODUCT_DETAILS = driver.find_elements_by_xpath("//div[@class = 'article__content showall review-post']")
                    product_details = [i.text for i in PRODUCT_DETAILS]
                except:
                    pass
            for i_product_prices in range(len(product_prices)):
                for i_product_details in range(len(product_details)):
                    CATE_.append(cate[i_cate])
                    NAME_SUB_CATE_.append(cate[i_cate])
                    PRODUCT_.append(PRODUCT)
                    RANK_.append(i_link_product + 1)
                    PRODUCT_PRICES_.append(product_prices[i_product_prices])
                    PRODUCT_DETAILS_.append(product_details[i_product_details])
                    print(link_product[i_link_product])
                    print(PRODUCT)
                    print(product_prices[i_product_prices])
                    print(product_details[i_product_details])
                    print('-----------------')
    elif (cate[i_cate] == "https://www.nhathuocankhang.com/cham-soc-tre-em"):
        driver.get(cate[i_cate] + "#&o=7&pi=2")
        print(driver.get(cate[i_cate] + "#&o=7&pi=2"))
        driver.implicitly_wait(6)
        LINK_PRODUCT = driver.find_elements_by_xpath("//ul[@class = 'listing-prod']//li//a[@href]")
        link_product = [i.get_attribute('href') for i in LINK_PRODUCT]
        print(link_product)
        for i_link_product in range(len(link_product)):
            driver.get(link_product[i_link_product])
            driver.implicitly_wait(5)
            try:
                PRODUCT_PRICES = driver.find_elements_by_xpath("//div[@class='box-price']//span")
                product_prices = [i.text for i in PRODUCT_PRICES]
            except NoSuchElementException:
                pass
            try:
                PRODUCT = driver.find_element_by_xpath("//div[@class = 'detail-title']//h3").text
            except NoSuchElementException:
                pass
            try:
                action = driver.find_element_by_xpath("//div//div[@class='detail-wrapper normal']//a[@class = 'btn-seemore']//span").click()
            except:
                pass
            time.sleep(2)
            try:
                PRODUCT_DETAILS = driver.find_elements_by_xpath("//div//div[@class='text-info']//div[@class='box-textdt']")
                product_details = [i.text for i in PRODUCT_DETAILS]
            except:
                try:
                    PRODUCT_DETAILS = driver.find_elements_by_xpath("//div[@class = 'article__content showall review-post']")
                    product_details = [i.text for i in PRODUCT_DETAILS]
                except:
                    pass
            for i_product_prices in range(len(product_prices)):
                for i_product_details in range(len(product_details)):
                    CATE_.append(cate[i_cate])
                    NAME_SUB_CATE_.append(cate[i_cate])
                    PRODUCT_.append(PRODUCT)
                    RANK_.append(i_link_product + 1)
                    PRODUCT_PRICES_.append(product_prices[i_product_prices])
                    PRODUCT_DETAILS_.append(product_details[i_product_details])
                    print(link_product[i_link_product])
                    print(PRODUCT)
                    print(product_prices[i_product_prices])
                    print('-----------------')
    else:
        for i_sub_cate in range(len(sub_cate)):
            if (cate[i_cate] == "https://www.nhathuocankhang.com/thuoc") and (sub_cate[i_sub_cate] != "https://www.nhathuocankhang.com/dau-cao-xoa-mieng-dan"):
                for i_type in range(1,2):
                    print(i_type)
                    if i_type == 2:
                        driver.get(sub_cate[i_sub_cate] + "#&protype=" + str(i_type) + "&o=7&pi=15")
                        driver.implicitly_wait(4)
                        print(name_sub_cate[i_sub_cate])
                        try:
                            LINK_PRODUCT = driver.find_elements_by_xpath("//ul[@class = 'listing-prod']//li//a[@href]")
                            link_product = [j.get_attribute('href') for j in LINK_PRODUCT]
                            print(link_product)
                        except NoSuchElementException:
                            break
                        for i_link_product in range(len(link_product)):
                            driver.get(link_product[i_link_product])
                            driver.implicitly_wait(5)
                            try:
                                PRODUCT_PRICES = driver.find_elements_by_xpath("//div[@class='box-price']//span")
                                product_prices = [i.text for i in PRODUCT_PRICES]
                            except NoSuchElementException:
                                pass
                            try:
                                PRODUCT = driver.find_element_by_xpath("//div[@class = 'detail-title']//h3").text
                            except NoSuchElementException:
                                pass
                            time.sleep(1)
                            try:
                                action = driver.find_element_by_xpath("//div//div[@class='detail-wrapper normal']//a[@class = 'btn-seemore']//span").click()
                            except:
                                pass
                            time.sleep(2)
                            try:
                                PRODUCT_DETAILS = driver.find_elements_by_xpath("//div//div[@class='text-info']//div[@class='box-textdt']")
                                product_details = [i.text for i in PRODUCT_DETAILS]
                            except:
                                pass
                            if product_prices != []:
                                for i_product_prices in range(len(product_prices)):
                                    for i_product_details in range(len(product_details)):
                                        CATE_.append(cate[i_cate])
                                        NAME_SUB_CATE_.append(name_sub_cate[i_sub_cate])
                                        PRODUCT_.append(PRODUCT)
                                        RANK_.append(i_link_product + 1)
                                        PRODUCT_PRICES_.append(product_prices[i_product_prices])
                                        PRODUCT_DETAILS_.append(product_details[i_product_details])
                                        print(link_product[i_link_product])
                                        print(PRODUCT)
                                        print(product_prices[i_product_prices])
                                        print(product_details[i_product_details])
                                        print('-----------------')
                            else:
                                for i_product_details in range(len(product_details)):
                                    CATE_.append(cate[i_cate])
                                    NAME_SUB_CATE_.append(name_sub_cate[i_sub_cate])
                                    PRODUCT_.append(PRODUCT)
                                    RANK_.append(i_link_product + 1)
                                    PRODUCT_PRICES_.append(product_prices)
                                    PRODUCT_DETAILS_.append(product_details[i_product_details])
                                    print(link_product[i_link_product])
                                    print(PRODUCT)
                                    print(product_prices)
                                    print(product_details[i_product_details])
                                    print('-----------------')
                    else:
                        driver.get(sub_cate[i_sub_cate] + "#&protype=" + str(i_type) + "&o=7&pi=15")
                        driver.implicitly_wait(4)
                        print(name_sub_cate[i_sub_cate])
                        time.sleep(3)
                        try:
                            LINK_PRODUCT = driver.find_elements_by_xpath("//ul[@class = 'listing-prod']//li//a//div[@class = 'text-prod']")
                            link_product = [j.text for j in LINK_PRODUCT]
                            print(link_product)
                        except NoSuchElementException:
                            pass
                        for i_link_product in range(len(link_product)):
                            print(i_link_product)
                            try:
                                PRODUCT_PRICES = driver.find_elements_by_xpath("//div[@class='box-price']//span")
                                product_prices = [i.text for i in PRODUCT_PRICES]
                            except NoSuchElementException:
                                product_prices = 'None'
                                pass
                            time.sleep(1)
                            try:
                                action = driver.find_element_by_xpath("//div//div[@class='detail-wrapper normal']//a[@class = 'btn-seemore']//span").click()
                            except:
                                pass
                            time.sleep(2)
                            try:
                                PRODUCT_DETAILS = driver.find_elements_by_xpath("//div//div[@class='text-info']//div[@class='box-textdt']")
                                product_details = [i.text for i in PRODUCT_DETAILS]
                            except:
                                product_details = 'None'
                                pass
                            if product_prices != []:
                                try:
                                    PRODUCT_PRICES = driver.find_elements_by_xpath("//div[@class='box-price']//span")
                                    product_prices = [i.text for i in PRODUCT_PRICES]
                                except NoSuchElementException:
                                    product_prices = 'None'
                                    pass
                                time.sleep(1)
                                try:
                                    action = driver.find_element_by_xpath(
                                        "//div//div[@class='detail-wrapper normal']//a[@class = 'btn-seemore']//span").click()
                                except:
                                    pass
                                time.sleep(2)
                                try:
                                    PRODUCT_DETAILS = driver.find_elements_by_xpath(
                                        "//div//div[@class='text-info']//div[@class='box-textdt']")
                                    product_details = [i.text for i in PRODUCT_DETAILS]
                                except:
                                    product_details = 'None'
                                    pass
                                for i_product_prices in range(len(product_prices)):
                                    print(i_product_prices)
                                    for i_product_details in range(len(product_details)):
                                        print(i_product_details)
                                        CATE_.append(cate[i_cate])
                                        NAME_SUB_CATE_.append(name_sub_cate[i_sub_cate])
                                        PRODUCT_.append(link_product[i_link_product])
                                        RANK_.append(i_link_product + 1)
                                        PRODUCT_PRICES_.append(product_prices[i_product_prices])
                                        PRODUCT_DETAILS_.append(product_details[i_product_details])
                                        print(link_product[i_link_product])
                                        print(product_prices[i_product_prices])
                                        print(product_details[i_product_details])
                                        print('-----------------')
                            else:
                                    CATE_.append(cate[i_cate])
                                    NAME_SUB_CATE_.append(name_sub_cate[i_sub_cate])
                                    PRODUCT_.append(link_product[i_link_product])
                                    RANK_.append(i_link_product + 1)
                                    PRODUCT_PRICES_.append(product_prices)
                                    PRODUCT_DETAILS_.append(product_details)
                                    print(link_product[i_link_product])
                                    print(product_prices)
                                    print(product_details)
                                    print('-----------------')
            else:
                driver.get(sub_cate[i_sub_cate] + "#&o=7&pi=20")
                driver.implicitly_wait(5)
                print(name_sub_cate[i_sub_cate])
                LINK_PRODUCT = driver.find_elements_by_xpath("//ul[@class = 'listing-prod']//li//a[@href]")
                link_product = [j.get_attribute('href') for j in LINK_PRODUCT]
                print(link_product)
                for i_link_product in range(len(link_product)):
                    try:
                        PRODUCT_PRICES = driver.find_elements_by_xpath("//div[@class='box-price']//span")
                        product_prices = [i.text for i in PRODUCT_PRICES]
                    except NoSuchElementException:
                        pass
                    try:
                        PRODUCT = driver.find_element_by_xpath("//div[@class = 'detail-title']//h3").text
                    except NoSuchElementException:
                        pass
                    try:
                        action = driver.find_element_by_xpath("//div//div[@class='detail-wrapper normal']//a[@class = 'btn-seemore']//span").click()
                    except:
                        pass
                    time.sleep(2)
                    try:
                        PRODUCT_DETAILS = driver.find_elements_by_xpath("//div//div[@class='text-info']//div[@class='box-textdt']")
                        product_details = [i.text for i in PRODUCT_DETAILS]
                    except:
                        try:
                            PRODUCT_DETAILS = driver.find_elements_by_xpath("//div[@class = 'article__content showall review-post']")
                            product_details = [i.text for i in PRODUCT_DETAILS]
                        except:
                            pass
                    for i_product_prices in range(len(product_prices)):
                        for i_product_details in range(len(product_details)):
                            CATE_.append(cate[i_cate])
                            NAME_SUB_CATE_.append(name_sub_cate[i_sub_cate])
                            PRODUCT_.append(PRODUCT)
                            RANK_.append(i_link_product + 1)
                            PRODUCT_PRICES_.append(product_prices[i_product_prices])
                            PRODUCT_DETAILS_.append(product_details[i_product_details])
                            print(link_product[i_link_product])
                            print(PRODUCT)
                            print(product_prices[i_product_prices])
                            print(product_details[i_product_details])
                            print('-----------------')
df['CATE'] = CATE_
df['SUB_CATE'] = NAME_SUB_CATE_
df['PRODUCT'] = PRODUCT_
df['RANK'] = RANK_
df['PRODUCT_PRICES'] = PRODUCT_PRICES_
df['PRODUCT_DETAILS'] = PRODUCT_DETAILS_

df.to_excel(r'C:\Users\User\Desktop\AnKhang-PRODUCT.xlsx', sheet_name='Your sheet name', index=False)