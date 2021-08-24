"""
This is the original that I started "functionalizing" and then moved functions as far
as I got into utils.py

It logs into webpage, pulls a list, modifies it, and prints using some options.
"""

import utils
import os
from secrets import user_name, pass_word
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

def start_driver():
    """Initialize Chrome WebDriver with option to save user data to local
     folder to handle cookies"""
    # todo: check for valid (up-to-date) webdriver
    # driver.get('chrome://settings/')
    # driver.set_window_size(1210, 720)

    # https://stackoverflow.com/questions/15058462/how-to-save-and-load-cookies-using-python-selenium-webdriver

    chrome_options = Options()

    # todo: following lines added 7/8 trying to make repo pretty. Gave up. Maybe later.
    # Or maybe someone else can tell what I was trying to do and make it work. :)
    #chrome_options.add_argument("--user-data-dir=" + os.path.join('..', 'io', 'chrome-data'))
    # chrome_options.add_argument("--user-data-dir= chrome-data")
    # #chrome_options.add_argument(r"--user-data-dir='..\io\chrome-data")
    # #driver = webdriver.Chrome('./chromedriver 83')
    # driver = webdriver.Chrome('./chromedriver 83', options=chrome_options)

    chrome_options.add_argument("--user-data-dir=chrome-data")
    #driver = webdriver.Chrome('./chromedriver 83', options=chrome_options)
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
    return driver


def get_page(driver):
    # Get webpage
    driver.get('https://www.votebuilder.com/Default.aspx')
    print("Driver title is: \n", driver.title)
    return  ## driver


def login_to_page(driver):
    ## login and initialize:
    # Click ActionID Button to open login
    driver.find_element_by_xpath("//a[@href='/OpenIdConnectLoginInitiator.ashx?ProviderID=4']").click()
    print("Driver title is: \n", driver.title)
    username = driver.find_element_by_id("username")
    username.send_keys(user_name)
    password = driver.find_element_by_id("password")
    password.send_keys(pass_word)
    driver.find_element_by_class_name("btn-blue").click()
    return


def remember_this():
    """Tries to check "remember this" checkbox. Might as well do manually."""
    driver.find_element_by_class_name("checkbox").click()
    # wait at least long enough to enter code and pin
    #driver.implicitly_wait(25)


def list_folders(driver):
    # List "My Folders" and select folder:
    wait_no_longer_than = 30
    element = WebDriverWait(driver, wait_no_longer_than).until(
                EC.presence_of_element_located((By.XPATH, '//a[@href="FolderList.aspx"]')))
    print(f'ELEMENT = {element}')
    driver.find_element_by_xpath('//a[@href="FolderList.aspx"]').click()


def select_folder(driver):
    print('Select Folder')
    driver.find_element_by_xpath('//*[text()="District 68 2020 3/17 Primary/Municipals"]').click()
    print("Driver title is: \n", driver.title)


def select_turf(driver, turf_name):
    print('Select Saved Search')
    driver.find_element_by_xpath('//*[text()="' + turf_name + '"]').click()


def handle_alert(driver):
    obj = driver.switch_to.alert
    obj.accept()


def edit_search(driver):
    # Edit Search:
    print('Edit Search')
    driver.find_element_by_xpath('//button[normalize-space()="Edit Search"]').click()


if __name__ == '__main__':
    driver = start_driver()
    get_page(driver)
    login_to_page(driver)
    #remember_this()
    list_folders(driver)
    # Then speed up (but tighten this up with EC WAITS
    driver.implicitly_wait(10)
    select_folder(driver)

    # Need to begin loop here to cycle through all turfs
    turf_name = 'P123 Turf 04'
    select_turf(driver, turf_name)
    handle_alert(driver)


    print("Driver title is: \n", driver.title)

    edit_search(driver)

    # Narrow search:
    print('...to narrow results')
    driver.find_element_by_id("stepTypeItem4").click()
    print("Driver title is: \n", driver.title)
    print('...to registered Democrats')
    while True:
        print('WHILE TRUE')
        print('Click Party Toggle')
        driver.find_element_by_id("ImageButtonSectionParty").click()
        print("Past Party Click")

        try:
            print("TRY Value=D")
            element = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.XPATH, '//input[@value="D"]'))
            )
            element_attribute_value = element.get_attribute('value')
            if element_attribute_value == None:
                print('NONE element.get_attribute(\'value\'): {0}'.format(element_attribute_value))
                raise ValueError("ValueError: None Attribute")
            else:
                print('ELSE: GOT element.get_attribute(\'value\'): {0}'.format(element_attribute_value))
                try:
                    print('Try Party Element Click')
                    print('Party Element Click')
                    element.click()
                    print('Past Party Element Click')
                    break
                except:
                    print('Party Click Exception try toggling Voting History')
                    #element = driver.find_element_by_xpath('//*[text()="Voting History"]').click()
                    driver.find_element_by_id("ImageButtonSectionVotingHistory").click()
                    print('Now try Party toggle again')
                    element = WebDriverWait(driver, 8).until(
                        EC.presence_of_element_located((By.XPATH, '//input[@value="D"]'))
                    )
        except:
            print("ValueError: Value='D' expected condition not met")

    print("Driver title is: \n", driver.title)
    print('Run Search to Narrow')
    element = driver.find_element_by_id("ctl00_ContentPlaceHolderVANPage_SearchRunButton").click()
    print("Driver title is: \n", driver.title)

    # #############
    # Done with narrow, now Edit to Remove
    # Edit Search:
    print('Edit Search to Remove People')
    driver.find_element_by_xpath('//button[normalize-space()="Edit Search"]').click()
    # Remove People
    driver.find_element_by_id("stepTypeItem3").click()
    print("Driver title is: \n", driver.title)
    # toggle history list
    print('Toggle Voting History')
    driver.find_element_by_id("ImageButtonSectionVotingHistory").click()

    print('Look for value ANY')
    element = driver.find_element_by_xpath("//select[@id='HistorySelection']/option[@value='ANY']").click()

    # 2016 Primary (NOT Pres)
    print('After ANY, before 2016 primary')
    element = driver.find_element_by_xpath("//input[@name='VotingInfo12696']").click()

    # 2016 Pres Primary
    print('After ANY, before 2016 primary')
    element = driver.find_element_by_xpath("//input[@name='VotingInfo12269']").click()

    print('After 2016 primary')

    # Select Notes="Voted"
    print('Find: ImageButtonSectionNotes')
    element = driver.find_element_by_id("ImageButtonSectionNotes").click()
    print('Find NoteText')
    element = driver.find_element_by_id('NoteText')
    print('Send Keys: Voted')
    element.send_keys('Voted')

    print('Run Search to Remove selected voters (Click Search Button)')
    element = driver.find_element_by_id("ctl00_ContentPlaceHolderVANPage_SearchRunButton").click()
    print("Driver title is: \n", driver.title)

    #####################################
    # Click Print Button to start Print Wizard
    driver.find_element_by_xpath('//a[@href="PrintReportWizard.aspx?RefererView=My%20List"]').click()

    # Select Report Format Option
    # Locate the Sector and create a Select object
    print('Select Print Format Option')
    select_element = Select(driver.find_element_by_id('ctl00_ContentPlaceHolderVANPage_VanDetailsItemReportFormatInfo_VANInputItemDetailsItemReportFormatInfo_ReportFormatInfo'
                                                      )
                            )
    element = select_element.select_by_visible_text("*2020 March 17 St. Petersburg Post 2/18_copy")

    # Select Script Option
    # Locate the Sector and create a Select object
    select_element = Select(driver.find_element_by_id("ctl00_ContentPlaceHolderVANPage_VanDetailsItemvdiScriptID_VANInputItemDetailsItemActiveScriptID_ActiveScriptID"
                                                      )
                            )
    print([o.text for o in select_element.options])

    element = select_element.select_by_visible_text('2020 3/17 St. Pete Post-2/18REVISED')

    # Select Contacted How Option
    # Locate the Sector and create a Select object
    select_element = Select(driver.find_element_by_id("ctl00_ContentPlaceHolderVANPage_VanDetailsItemVANDetailsItemScriptSource_ScriptSource_VANInputItemDetailsItemScriptSource_ScriptSource"
                                                      )
                            )
    print([o.text for o in select_element.options])

    element = select_element.select_by_visible_text("Walk")

    #####################################
    # Select List Name Input (Enter Print File Name)

    element = driver.find_element_by_id("ctl00_ContentPlaceHolderVANPage_VANDetailsItemReportTitle_VANInputItemDetailsItemReportTitle_ReportTitle")
    element.clear()
    element.send_keys('P123T04ThomasSignatureDemsUnvoted')

    #####################################
    # Exclude Absentee/Early Voters
    driver.find_element_by_id("ctl00_ContentPlaceHolderVANPage_VANDetailsItemExcludeVoted_VANInputItemDetailsItemExcludeVoted_ExcludeVoted_0").click()

    # Click Next for next form page (Submit print job)
    driver.find_element_by_id("ctl00_ContentPlaceHolderVANPage_ButtonSortOptionsSubmit").click()


    # if element is not None:
    #     print('Element: ', element)
    #     element_text = element.text
    #     element_attribute_value = element.get_attribute('value')
    #     print(element.get_attribute('innerHTML'))
    #     print('element.text: {0}'.format(element_text))
    #     print('element.get_attribute(\'value\'): {0}'.format(element_attribute_value))


    print("Driver title is: \n", driver.title)
    # for glob in glob.iglob(os.path.join('C:', 'Users', 'admin', 'AppData', 'Local', 'Temp', 'scoped_dir*')):
    #     print(glob)
    #     shutil.rmtree(glob)

    print('bye')

