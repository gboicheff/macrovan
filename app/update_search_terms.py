from utils import *
# from bs4 import BeautifulSoup
import time
import pickle
import random
import sys
import os
import re
from selenium.webdriver.support.select import Select
import tkinter as tk
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException

path = os.getcwd()
print(f"The current working directory is {path}")



def execute(driver, selected_folder, selected_row_type):
    get_page(driver)
    driver.implicitly_wait(60)
    expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_HyperLinkMenuSavedLists"]').click()


    folder_name = '//*[text()="{}"]'.format(selected_folder)
    expect_by_XPATH(driver, folder_name).click()

    # find num rows
    try:
        num_rows = int(expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_gvList"]/tfoot/tr/td/table/tbody/tr/td[1]/b[1]', 5).text.split()[0])
    except TimeoutException:
        num_rows = int(expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_gvList"]/tfoot/tr/td/b[1]', 5).text.split()[0])

    # set the number of rows per page to 999
    settings_button = expect_by_XPATH(driver, '//*[@id="HyperLinkSettings"]')
    settings_button.click()
    rows_p_page_input = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_VANDetailsItemDefaultRows_VANInputItemDetailsItemDefaultRows_DefaultRows"]')

    rows_p_page_input.clear()
    rows_p_page_input.send_keys("999")
   

    save_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_ButtonSave"]')
    save_button.click()


    # for index in range(1,num_rows+1):
    max_errors = 15
    total_errors = 0
    index = 1
    while index <= num_rows: 
        if total_errors == max_errors:
            print("Hit max amount of errors.  Something is probably wrong.  Shutting down to prevent infinite loop")
            break
        try:
            time.sleep(2 + random.randint(1,3))
            row_xpath = '//*[@id="ctl00_ContentPlaceHolderVANPage_gvList"]/tbody/tr[{index}]'.format(index=index)

            # check for map turf
            row_type = expect_by_XPATH(driver, row_xpath + '/td[3]/span').text

            if (row_type == "Map Region" and row_type == selected_row_type) or (row_type == "Map Turf" and row_type == selected_row_type):

                # get stuff on edit page
                button = expect_by_XPATH(driver, row_xpath + '/td[4]')

                edit_button = expect_by_tag(button, "a")
                edit_button.click()

                # accept alert if present
                try:
                    WebDriverWait(driver, 5).until(EC.alert_is_present())
                    alert = driver.switch_to.alert
                    alert.accept()
                    print("alert Exists in page")
                except TimeoutException:
                    print("alert does not Exist in page")
                
                time.sleep(1.5)
                search_button = expect_by_XPATH(driver, '//*[@id="addStep"]')
                search_button.click()

                edit_s_button = expect_by_XPATH(driver, '//*[@id="editSearchOption"]')
                edit_s_button.click()


                #click address dropdown

                a_button = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionLocationGroups"]')
                a_button.click()

                city_input = expect_by_XPATH(driver, '//*[@id="City"]')
                city_input.send_keys("Saint Petersburg") # make this a varialbe?

                save_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_SearchRunButton"]')
                save_button.click()


                time.sleep(5)
                driver.get("https://www.votebuilder.com/Default.aspx")

                expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_HyperLinkMenuSavedLists"]').click()

                folder_name = '//*[text()="{}"]'.format(selected_folder)
                expect_by_XPATH(driver, folder_name).click()
                print(index)
            index+=1

        except Exception as e:
            total_errors+=1
            print(e)
            driver.get("https://www.votebuilder.com/Default.aspx")
            expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_HyperLinkMenuSavedLists"]').click()
            folder_name = '//*[text()="{}"]'.format(selected_folder)
            expect_by_XPATH(driver, folder_name).click()






def run(root):
    teardown()
    driver = start_driver(os.path.join(path, "chrome-data"))
    driver.maximize_window()
    execute(driver, "*2021 Municipal St. Petersburg Master", "Map Region")
    root.destroy()

if __name__ == '__main__':
    root = tk.Tk()
    root.geometry("300x300")
    check_button_city = tk.Button(root, text="Start", command=lambda root=root:run(root), bg="green", width=100, height=50)
    check_button_city.pack()
    root.mainloop()
    print("done")

    