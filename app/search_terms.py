from utils import *
# from bs4 import BeautifulSoup
import time
import pickle
import random
import sys
import os
import re

path = os.getcwd()
print(f"The current working directory is {path}")


class MapRegion:
    ID = -1
    NAME = "NONE"
    LAST_REFRESH = "NONE"
    PEOPLE = -1
    HOME_PHONES = -1
    PREFERRED_PHONES = -1
    DOORS = -1
    MAILBOXES = -1

    def display(self):
        output = """
        ID:{id}
        NAME:{name}
        PEOPLE:{people}
        HOME_PHONES:{h_phones}
        PREFERRED_PHONES:{p_phones}
        DOORS:{doors}
        MAILBOXES:{m_boxes}
        """.format(id = self.ID, name=self.NAME, people=self.PEOPLE, h_phones = self.HOME_PHONES, p_phones = self.PREFERRED_PHONES, doors = self.DOORS, m_boxes = self.MAILBOXES)
        print(output)

    def get_turf(self):
        try:
            num = re.search("Turf [0-9]+", self.NAME).group()
            num = int(num.split()[-1])
        except:
            num = "ERROR"
        return num

    def get_precinct(self):
        try:
            num = re.search("(P|p)(_| )?[0-9]+", self.NAME).group()
            num = int("".join([digit for digit in num if digit.isnumeric()]))
        except:
            num = "ERROR"
        return num

    
    def flatten(self):
        flattened_region = dict([("ID", self.ID), ("Name", self.NAME), ("People", self.PEOPLE), ("Home Phone Ct", self.HOME_PHONES), ("Pref Phone Ct", self.PREFERRED_PHONES),
        ("Doors", self.DOORS), ("Mailboxes", self.MAILBOXES)])
        return flattened_region


def execute(driver, selected_folder, selected_row_type, selected_output_name):
    get_page(driver)
    driver.implicitly_wait(10)
    attempt_login(driver)  
    list_folders(driver)


    folder_name = '//*[text()="{}"]'.format(selected_folder)
    expect_by_XPATH(driver, folder_name).click()
    # table = expect_by_tag(expect_by_tag(driver, "table"), "tbody")

    map_regions = []

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


    for index in range(1,num_rows+1):


        time.sleep(random.randint(1,3))
        row_xpath = '//*[@id="ctl00_ContentPlaceHolderVANPage_gvList"]/tbody/tr[{index}]'.format(index=index)

        # check for map turf
        row_type = expect_by_XPATH(driver, row_xpath + '/td[3]/span').text
        print("ROW_TYPE:{}".format(row_type))
        if (row_type == "Map Region" and row_type == selected_row_type) or (row_type == "Map Turf" and row_type == selected_row_type):
                map_region = MapRegion()
                map_region.ID = expect_by_XPATH(driver, row_xpath + '/td[2]/span').text
                map_region.NAME = expect_by_XPATH(driver, row_xpath + '/td[4]/a/span').text
                
                # get stuff on edit page
                button = expect_by_XPATH(driver, row_xpath + '/td[4]')
                edit_button = expect_by_tag(button, "a")
                edit_button.click()

                time.sleep(0.25)
                driver.switch_to.alert.accept()

                map_region.PEOPLE = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_VoterCount"]').text
                map_region.HOME_PHONES = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_PhoneCount"]').text
                map_region.PREFERRED_PHONES = expect_by_XPATH(driver, '//*[@id="content"]/div/div[1]/div[3]/div[3]/ul/li[3]/div/div[1]').text
                map_region.DOORS = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_DoorCount"]').text
                map_region.MAILBOXES = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_MailingCount"]').text

                # back to table
                driver.back()


                map_region.display()
                map_regions.append(map_region)

    if os.path.isfile("{}.pkl".format(selected_output_name)):
        os.remove("{}.pkl".format(selected_output_name))

    with open("{}.pkl".format(selected_output_name), "wb") as file:
        pickle.dump(map_regions, file)


if __name__ == '__main__':
    execute("*2021 Municipal St. Petersburg Master", "Any", "test")

    