ameimport os
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
import ExcelAnalysis

file = "path.xlsx"
username = "<username>"
password = "<password>"
company = "<company_id>"
user = "<name as saved organizer app>"

url = "https://app.shiftorganizer.com/login/?lang=he&previous=homepage&greeting=true"


def connectToShift():

    driver = webdriver.Chrome("path")

    driver.get(url)

    driver.find_element(By.NAME, "Company").send_keys(company)
    driver.find_element(By.NAME, "password").send_keys(password)
    driver.find_element(By.NAME, "user").send_keys(username)
    driver.find_element(By.ID, "log-in").click()
    time.sleep(2)
    driver.get("https://app.shiftorganizer.com/app/rota")
    time.sleep(2)

    if os.path.exists(file):
        os.remove(file)
        print("Old file was deleted\nDownloading new file\n")
    else:
        print("Old file was not found\nDownloading new file\n")
    time.sleep(5)
    try:
        driver.find_element(By.CSS_SELECTOR, ".btn.btn-success.btn-block").click()
        time.sleep(5)
        ExcelAnalysis.analyzeExcel(file,user)
    except:
        print("Can't download excel file from https://app.shiftorganizer.com/app/rota\n")
