import pyautogui as p
import selenium
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By

def main_function():
    try:
        options = get_options()
        open_driver(options)
    except Exception as erro:
        print(erro)

def open_driver(options):
    try:
        driver = webdriver.Chrome(options=options)
        driver.get("https://rpachallenge.com")
        get_file(driver)
        start(driver)
        get_information(driver)
        p.sleep(1)
        delete()
    except Exception as erro:
        print(erro)

def get_options():
    try:
        chrome_options = webdriver.ChromeOptions()
        prefs = {'download.default_directory' : r'C:\Users\user\Desktop\rpa'}
        chrome_options.add_experimental_option('prefs', prefs)
        return chrome_options
    except Exception as erro:
        print(erro)

def get_file(driver):
    try:
        driver.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a").click()
        flag = True
        arquive = r"C:\Users\user\Desktop\rpa\challenge.xlsx"
        while flag:
            if(os.path.exists(arquive)):
                return False
            p.sleep(1)
    except Exception as erro:
        print(erro)

def get_information(driver):
    try:
        df = pd.read_excel("challenge.xlsx")
        p.sleep(1)
        for row in range(len(df)):
            first_name = df["First Name"][row]
            last_name = df["Last Name "][row]
            company_name = df["Company Name"][row]
            role_in_company = df["Role in Company"][row]
            address = df["Address"][row]
            email = df["Email"][row]
            phone_number = df["Phone Number"][row]
            driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelFirstName"]').send_keys(first_name)
            driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelLastName"]').send_keys(last_name)
            driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelCompanyName"]').send_keys(company_name)
            driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelRole"]').send_keys(role_in_company)
            driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelAddress"]').send_keys(address)
            driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelEmail"]').send_keys(email)
            driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelPhone"]').send_keys(str(phone_number))
            p.sleep(1)
            driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()
    except Exception as erro:
        print(erro)
        
def start(driver):
    try:
        driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()
    except Exception as erro:
        print(erro)

def delete():
    try:
        path = r"C:\Users\user\Desktop\rpa\challenge.xlsx"
        if os.path.exists(path):
            os.remove(path)
        else:
            print("deleted")
    except Exception as erro:
        print(erro)

main_function()