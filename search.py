
# -*- coding: utf-8 -*-
"""
Created on Thu Apr 21 10:59:17 2022
@author: TommyLo
"""

#################################################
## Google Keywords Search Result Scraper
##
## How to use:
## - input the list of keywords interested into the Gsheet (or xlsx file) 
## - run the script and it will return the result of first X pages (X to be defined on the script)
#################################################

from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import date
import time
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


def wait_until_async_results_available(driver):
    timeout = 10

    try:
        # Use WebDriverWait to wait for either the element or the timeout
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, 'div[id="botstuff"] div[decode-data-ved="1"] > div')
            )
        )
        # If the element is found, perform actions on it here

    except Exception as e:
        print("Timed out waiting for the element to be found", e)


def search_using_automation(keyword, site, device):
    # No. of top N result extracting 
    n_result = 100

    # initializing
    se_result = None
    n_keyword = 0

    print('searching', keyword)
    #Loop over keyword list

        #Start webdriver
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument("--test-type")

    mobile_emulation = { "deviceName": "Nexus 5" }
    # if device == 'Mobile':
        # options.add_experimental_option("mobileEmulation", mobile_emulation)
    # options.add_argument("--headless") #hidden browser
    #    options.add_experimental_option('excludeSwitches', ['enable-logging']) #To fix device event log error
    #    options.use_chromium = True  #To fix device event log error
    #    options.add_argument("user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 13_2_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.3 Mobile/15E148 Safari/604.1")
    #    options.binary_location = r"C:\Program Files\Google\Chrome\Application\Chrome.exe"
    #    driver = webdriver.Chrome(executable_path=r'C:\TommyLo\google-rank-tracker\chromedriver.exe',chrome_options=options)    
    driver = webdriver.Chrome(options=options)

    n_rank = 0 #reset rank

    n_keyword += 1
    url = 'https://www.google.com/search?num={}&q={}'.format(n_result,keyword)
    print('#{} --- {} --- {} ...'.format(n_keyword,keyword,url))
    driver.get(url)    
    time.sleep(2)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    #Organic class start with 'yuRUbf'  
    wait_until_async_results_available(driver)
    page = driver.execute_script('return document.body.innerHTML')
    soup = BeautifulSoup(''.join(page), 'html.parser')
    # soup = BeautifulSoup(page_source, 'lxml')
    results_selector = soup.select('div[id="rso"] > div, div[id="botstuff"] div[decode-data-ved="1"] > div')
    print(len(results_selector))
    #Loop over the results
    for result_selector in results_selector:             
                           
        people_also_ask_selector = result_selector.select('div[class*="cUnQKe"]')
        if len(people_also_ask_selector) > 0:
            print('Skipping people also ask')
            continue
        main_result_selector = result_selector.select('div[class*="yuRUbf"]')
        if len(main_result_selector) == 0:
            print('Skipping main result')
            continue
        link = main_result_selector[0].select('a')[0]['href']
        print(link, n_rank)
        n_rank += 1
        if site in link or n_rank > n_result:
            se_result = {
                'query_date' : date.today().strftime("%Y%m%d"),
                'keyword' : keyword,
                'rank' : n_rank,
                'link' : link
                }        
            break   
    
    driver.close()
    driver.quit()
    return se_result 

