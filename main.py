from selenium import webdriver
import os
from config import EMAIL,PASSWORD
import pandas as pd


excel = pd.read_excel('data.xlsx')


paths = os.path.join(os.getcwd(), "chromedriver_win32_2", 'chromedriver.exe')
browser = webdriver.Chrome(executable_path= paths)


browser.get('http://www.gmail.com')
browser.maximize_window()


email_input = browser.find_element_by_name("identifier")
email_input.send_keys(EMAIL)


next_btn =browser.find_element_by_xpath('//*[@id="identifierNext"]/div/button/span')
next_btn.click()


browser.implicitly_wait(4)


password_input =browser.find_element_by_name("password")
password_input.send_keys(PASSWORD)


next_btn =browser.find_element_by_xpath('//*[@id="passwordNext"]/div/button/span')
next_btn.click()


for x in excel.index:
    print(excel.loc[x], end='/n/n')

    browser.implicitly_wait(4)


    compose_btn = browser.find_element_by_css_selector('.T-I.T-I-KE.L3')
    compose_btn.click()


    browser.implicitly_wait(4)


    to_field = browser.find_element_by_name("to")
    to_field.send_keys(excel.loc[x]['Email'])


    subject_field = browser.find_element_by_name("subjectbox")
    subject_field.send_keys(excel.loc[x]['Subject'])


    body_field = browser.find_element_by_css_selector('.Am.Al.editable.LW-avf.tS-tW')
    body_field.send_keys(excel.loc[x]['Body'])


    send_btn = browser.find_element_by_css_selector('.T-I.J-J5-Ji.aoO.v7.T-I-atl.L3')
    send_btn.click()


    browser.implicitly_wait(4)



