import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
from datetime import datetime

# load the Excel file and turn it into a dataframe
file_path = 'PO_DATA.xlsx'
df = pd.read_excel(file_path, dtype=str)

# for Selenium WebDriver
headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,"
              "application/signed-exchange;v=b3;q=0.7",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "en-US,en;q=0.9",
    "Priority": "u=0, i",
    "Sec-Ch-Ua": "\"Not)A;Brand\";v=\"99\", \"Microsoft Edge\";v=\"127\", \"Chromium\";v=\"127\"",
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": "\"Windows\"",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "cross-site",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0",
}

# credentials stored in environment variables (another location) for security
USER_ID = os.environ["USER_ID"]
PASSWORD = os.environ["PASSWORD"]

# all user PO information must be entered into the PO_DATA.xlsx file
BUSINESS_UNIT = df.iloc[:, 0].tolist()
VENDOR_SHORT_NAME_LOOKUP = df.iloc[:, 1].tolist()
ITEM_CODE = df.iloc[:, 2].tolist()
PO_QTY = df.iloc[:, 3].tolist()
PRICE = df.iloc[:, 4].tolist()
DELIVERY_DATE = df.iloc[:, 5].tolist()
SHIP_TO_LOCATION = df.iloc[:, 6].tolist()
COMMENT_TYPE = df.iloc[:, 7].tolist()
COMMENT_ID = df.iloc[:, 8].tolist()


def sleep(y):
    time.sleep(y)


def create_po(x):
    sleep(.8)
    # instantiate Chrome with Selenium
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=chrome_options)
    action = ActionChains(driver)
    # driver.get('http://ccs.combe.com/Pages/Default.aspx')
    driver.get('http://peoplesoft.global.combe.com:11080/psp/fs88prod/?cmd=login&errorPg=ckreq&languageCd=ENG')

    # log in
    driver.find_element(By.XPATH, '//*[@id="login"]/table/tbody/tr/td/p[3]/a').click()
    user_id = driver.find_element(By.ID, 'userid')
    user_id.send_keys(USER_ID, Keys.TAB, PASSWORD, Keys.RETURN)
    sleep(1.2)

    # instantiate the new PO
    driver.get('http://peoplesoft.global.combe.com:11080/psp/fs88prod/EMPLOYEE/ERP/c/MANAGE_PURCHASE_ORDERS'
               '.PURCHASE_ORDER.GBL?PORTALPARAM_PTCNAV=EP_PURCHASE_ORDER_GBL&EOPP.SCNode=ERP&EOPP.SCPortal=EMPLOYEE&EOPP'
               '.SCName=EPPO_PURCHASING&EOPP.SCLabel=Purchasing&EOPP.SCPTfname=EPPO_PURCHASING&FolderPath'
               '=PORTAL_ROOT_OBJECT.EPPO_PURCHASING.EPCO_PURCHASE_ORDERS1.EP_PURCHASE_ORDER_GBL&IsFolder=false')
    sleep(1.5)
    action.send_keys(Keys.DELETE, Keys.DELETE, Keys.DELETE, Keys.DELETE, Keys.DELETE, BUSINESS_UNIT[x], Keys.RETURN)
    action.perform()
    sleep(2.5)

    # search supplier by short name
    action.send_keys(Keys.TAB, Keys.RETURN)
    action.perform()
    sleep(3)
    action.key_down(Keys.SHIFT)
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.TAB, Keys.TAB, VENDOR_SHORT_NAME_LOOKUP[x])
    action.perform()
    action.key_up(Keys.SHIFT)
    action.send_keys(Keys.RETURN)
    action.perform()
    sleep(2)
    action.send_keys(Keys.RETURN)
    action.perform()
    sleep(2)

    # enter first page item detail information
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, USER_ID)
    action.perform()
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     ITEM_CODE[x], Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, PO_QTY[x], Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.TAB, PRICE[x])
    action.perform()

    # navigate to the second page
    action.send_keys(Keys.TAB, Keys.TAB, Keys.RETURN)
    action.perform()
    sleep(1.8)

    # add date, which converts the string to a datetime, then reorder and strips it when going back to a string
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     datetime.strftime(datetime.strptime(DELIVERY_DATE[x], "%Y-%m-%d %H:%M:%S"), "%m/%d/%Y"),
                     Keys.RETURN, Keys.RETURN)
    action.perform()
    sleep(1)

    # out of range date check
    action.send_keys(Keys.RETURN)
    action.perform()
    sleep(.6)
    action.send_keys(Keys.RETURN)
    action.perform()
    sleep(.6)

    # add ship to location
    action.send_keys(Keys.TAB, Keys.TAB, SHIP_TO_LOCATION[x], Keys.RETURN)
    action.perform()
    sleep(1)

    # # CUSTOM PRICE FIX
    # action.send_keys(Keys.TAB, Keys.RETURN)
    # action.perform()
    # sleep(.5)

    # navigate to distributions/chart fields page
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.RETURN)
    action.perform()
    sleep(1.2)

    # navigate to tax details tab
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.RETURN)
    action.perform()
    sleep(1.2)

    # enter tax details
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     SHIP_TO_LOCATION[x], Keys.RETURN)
    action.perform()
    sleep(1.7)

    # return to main page
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.RETURN)
    action.perform()
    sleep(1)
    action.send_keys(Keys.RETURN)
    action.perform()
    sleep(1.8)

    # navigate to comment section
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB)
    action.perform()
    action.send_keys(Keys.RETURN)
    action.perform()
    sleep(.8)

    # add comments
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.RETURN)
    action.perform()
    sleep(.9)
    action.send_keys(Keys.TAB, COMMENT_TYPE[x], Keys.TAB, Keys.TAB, COMMENT_ID[x])
    action.perform()
    sleep(.6)

    # save comments
    action.send_keys(Keys.TAB, Keys.RETURN)
    action.perform()
    sleep(.5)

    # return to the main page
    action.send_keys(Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB, Keys.TAB,
                     Keys.TAB, Keys.RETURN)
    action.perform()


for z in range(0, len(ITEM_CODE)):
    create_po(z)
