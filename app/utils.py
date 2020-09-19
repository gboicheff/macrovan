from secrets import *
import os
import io
import sys
import PyPDF2
import glob
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.remote.command import Command
import ctypes  # for windows message pop-up
import pandas as pd


def get_os():
    # print(sys.platform)
    if "win" in sys.platform:
        # print("os = Windows")
        return "Windows"


def teardown():
    """Remove temp files from prior run before starting driver"""

    print('Start Teardown')
    if get_os() == "Windows":
        print("Do Windows")
        windowsUser = os.getlogin()
        for path in glob.iglob(os.path.join('C:\\', 'Users', windowsUser, 'AppData', 'Local', 'Temp', 'scoped_dir*')):
            print(path)
            shutil.rmtree(path)

        for path in glob.iglob(os.path.join('C:\\', 'Users', windowsUser, 'AppData', 'Local', 'Temp', 'chrome_BITS_*')):
            print(path)
            shutil.rmtree(path)
    print('Teardown complete')


def pause(message):
    ctypes.windll.user32.MessageBoxW(0, message, "Macrovan", 1)


def start_driver():
    """Initialize Chrome WebDriver with option that saves user-data-dir to local
     folder to handle cookies"""
    # driver.get('chrome://settings/')
    # driver.set_window_size(1210, 720)

    # https://stackoverflow.com/questions/15058462/how-to-save-and-load-cookies-using-python-selenium-webdriver

    chrome_options = Options()

    # todo: following lines added 7/8 trying to make repo pretty.
    #  Trying to save chrome-data elsewhere. Gave up. Maybe later.
    # chrome_options.add_argument(r"--user-data-dir='..\io\chrome-data'")
    # #chrome_options.add_argument("--enable-caret-browsing")

    # adding argument causes chrome to open with address bar highlighted and I can't figure out why!
    chrome_options.add_argument("--user-data-dir=chrome-data")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_experimental_option("excludeSwitches", ['enable-logging'])
    chrome_options.add_argument('disable-infobars')
    # display_to_console("Loading...")
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
    # display_to_console("Finished loading!")
    return driver


def print_title(driver):
    print("Driver title is: \n", driver.title)


def get_page(driver):
    # Get webpage
    driver.get('https://www.votebuilder.com/Default.aspx')
    print_title(driver)
    return


def login_to_page(driver):
    # login and initialize:
    # Click ActionID Button to open login
    element = expect_by_XPATH(driver, '//a[@href="/OpenIdConnectLoginInitiator.ashx?ProviderID=4"]')
    # print(f'ELEMENT = {element}')

    # driver.find_element_by_xpath("//a[@href='/OpenIdConnectLoginInitiator.ashx?ProviderID=4']").click()
    expect_by_XPATH(driver, "//a[@href='/OpenIdConnectLoginInitiator.ashx?ProviderID=4']").click()
    print('After ActionID Button')
    print_title(driver)
    expect_by_id(driver, 'username')
    username = expect_by_id(driver, "username")
    username.send_keys(user_name)
    password = expect_by_id(driver, "password")
    password.send_keys(pass_word)
    expect_by_class(driver, "btn-blue").click()
    return


def remember_this(driver):
    expect_by_class(driver, "checkbox").click()
    # wait at least long enough to enter code and pin
    # driver.implicitly_wait(25)


def list_folders(driver):
    # List "My Folders"
    expect_by_XPATH(driver, '//a[@href="FolderList.aspx"]').click()
    print('AFTER FOLDER LIST CLICK')


def select_folder(driver):
    """select folder"""
    print('Select Folder')
    # driver.find_element_by_xpath('//*[text()="District 68 2020 3/17 Primary/Municipals"]').click()
    # expect_by_XPATH(driver, '//*[text()="2020 District 68"]').click()
    expect_by_XPATH(driver, '//*[text()="2020 District 68 November"]').click()
    print_title(driver)


def select_turf(driver, turf_name):
    print('Select Saved Search')
    expect_by_id(driver,
                 "ctl00_ContentPlaceHolderVANPage_VanInputItemviiFilterName_VanInputItemviiFilterName").send_keys(
        turf_name)
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_RefreshFilterButton").click()
    expect_by_XPATH(driver, '//*[text()="' + turf_name + '"]').click()


def handle_alert(driver):
    """Are you sure you want to load this Map Turf
    and overwrite your current version of My List"""
    obj = driver.switch_to.alert
    obj.accept()
    print_title(driver)


def edit_search(driver):
    # Edit Search:
    print('Edit Search')
    expect_by_XPATH(driver, '//button[normalize-space()="Edit Search"]').click()


def early_voting_twisty(driver):
    # to click Early Voting Twisty
    element = expect_by_id(driver, 'ImageButtonSectionEarlyVoting')
    print(f'Early Voting Section element located = {element}')
    expect_by_id(driver, "ImageButtonSectionEarlyVoting").click()


def notes_twisty(driver):
    element = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionNotes"]')
    print('Try to click "Notes" twisty')
    expect_by_XPATH(driver, '//*[@id="ImageButtonSectionNotes"]').click()


# Returns list of turf name and last name pairs under a provided captain
def get_turfs_by_captain(captain, turf_dict):
    try:
        return turf_dict[captain]
    except KeyError:
        print("Turf captain doesn't exist: " + captain[0] + " " + captain[1])


# Returns list of all turf name and last name pairs
def get_all_turfs(turf_data):
    output = []
    for item in turf_data.values():
        output += item
    return output


# Return list of all block captains
def get_all_captains(turf_dict):
    output = []
    for item in turf_dict.keys():
        output += [item]
    return output


def turfselection_plus(driver, turf_name):
    # ORIGINAL (SIDE) Test name: from turf selection

    # SELECT TURF NAME
    # use turf name selection method from macrovan
    print(f'Select turf_name = {turf_name}')
    select_turf(driver, turf_name)
    print('Handle erase current list alert')
    handle_alert(driver)

    expect_by_id(driver, "addStep").click()
    expect_by_id(driver, "stepTypeItem4").click()
    early_voting_twisty(driver)
    print('Click anyone Who Requested a Ballot')
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_EarlyVoteCheckboxId_RequestReceived").click()
    print('Click Preview Button')
    expect_by_id(driver, "ResultsPreviewButton").click()
    print("Driver title is: \n", driver.title)
    print('Click #AddNewStepButton')
    pause('Click Add New Step: Remove, and wait\n for page to load to continue')
    print('Unclick early voting twisty?')
    early_voting_twisty(driver)

    # click notes twisty
    notes_twisty(driver)

    print('Click in note text field. Is this needed?')
    expect_by_id(driver, "NoteText").click()
    print('Send keys to NoteText "*moved')
    expect_by_id(driver, "NoteText").send_keys("*moved")
    print(f'Sent keys *moved for remove step')

    # unclick notes_twisty
    notes_twisty(driver)

    print('Run Search to Remove selected voters (Click Run Search Button)')
    element = expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_SearchRunButton").click()
    print("Driver title is: \n", driver.title)
    print("And done with turfselection_plus SIDE function")


def print_list(driver, listName):
    # Print a List
    print('in print_list waiting for print icon')
    element = expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_HyperLinkImagePrintReportsAndForms")
    print('in print_list trying to click')
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_HyperLinkImagePrintReportsAndForms").click()
    print('just clicked print icon might need another EC')

    # Select Report Format Option
    # Locate the Sector and create a Select object
    print('Select Print Format Option')
    select_element = Select(expect_by_id(driver,
                                         "ctl00_ContentPlaceHolderVANPage_VanDetailsItemReportFormatInfo_VANInputItemDetailsItemReportFormatInfo_ReportFormatInfo"))
    element = select_element.select_by_visible_text("*2020 D68 Aug Primary")

    # Select Script Option
    # Locate the Sector and create a Select object
    select_element = Select(expect_by_id(driver,
                                         "ctl00_ContentPlaceHolderVANPage_VanDetailsItemvdiScriptID_VANInputItemDetailsItemActiveScriptID_ActiveScriptID"))
    # print([o.text for o in select_element.options])
    element = select_element.select_by_visible_text('*2020 D68 Aug Primary')

    expect_by_id(driver,
                 "ctl00_ContentPlaceHolderVANPage_VanDetailsItemvdiScriptID_VANInputItemDetailsItemActiveScriptID_ActiveScriptID").click()

    # Script source selection (Walk)
    expect_by_id(driver,
                 "ctl00_ContentPlaceHolderVANPage_VanDetailsItemVANDetailsItemScriptSource_ScriptSource_VANInputItemDetailsItemScriptSource_ScriptSource").click()
    dropdown = expect_by_id(driver,
                            "ctl00_ContentPlaceHolderVANPage_VanDetailsItemVANDetailsItemScriptSource_ScriptSource_VANInputItemDetailsItemScriptSource_ScriptSource")
    expect_by_XPATH(driver, "//option[. = 'Walk']").click()
    expect_by_id(driver,
                 "ctl00_ContentPlaceHolderVANPage_VanDetailsItemVANDetailsItemScriptSource_ScriptSource_VANInputItemDetailsItemScriptSource_ScriptSource").click()
    element = expect_by_id(driver,
                           "ctl00_ContentPlaceHolderVANPage_VANDetailsItemReportTitle_VANInputItemDetailsItemReportTitle_ReportTitle")
    element.clear()
    element.send_keys(listName)

    # Deselect Headers amd Breaks
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder1_Header1").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder1_Break1").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder2_Header2").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder2_Break2").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder3_Header3").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder3_Break3").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_Header4").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_Break4").click()

    # Sort Order 4
    expect_by_id(driver,
                 "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_VANInputItemDetailsItemSortOrder4_SortOrder4").click()
    dropdown = Select(expect_by_id(driver,
                                   "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_VANInputItemDetailsItemSortOrder4_SortOrder4"))
    dropdown.select_by_index(4)

    # Sort Order 5
    expect_by_id(driver,
                 "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder5_VANInputItemDetailsItemSortOrder5_SortOrder5").click()
    dropdown = Select(expect_by_id(driver,
                                   "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder5_VANInputItemDetailsItemSortOrder5_SortOrder5"))
    dropdown.select_by_index(5)

    # Sort Order 6
    expect_by_id(driver,
                 "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder6_VANInputItemDetailsItemSortOrder6_SortOrder6").click()
    dropdown = Select(expect_by_id(driver,
                                   "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder6_VANInputItemDetailsItemSortOrder6_SortOrder6"))
    dropdown.select_by_index(0)

    # Submit
    expect_by_id(driver,
                 "ctl00_ContentPlaceHolderVANPage_VanDetailsItemPrintMapNew_VANInputItemDetailsItemPrintMapNew_PrintMapNew_0")
    pause("Double Check that selections are correct").click()  #todo: test removal of .click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_ButtonSortOptionsSubmit").click()
    expect_by_link_text(driver, "My PDF Files").click()


def return_to_home(driver):
    expect_by_link_text(driver, "Home").click()


# Close everything and cleanup
def exit_program(window, driver):
    try:
        window.destroy()
    except:
        print("Window does not exist!")
    else:
        print("Window closed!")

    try:
        driver.close
        driver.quit()
    except:
        print("Driver does not exist!")
    else:
        print("Driver closed!")

    try:
        teardown()
    except:
        print("Teardown failed!")
    # else:
    # print("Teardown successfully ran!")


# Checks if the chrome browser is open or not closes everything if the chrome browser closed.
def check_browser(window, driver):
    if len(driver.get_log('driver')) > 0:
        if driver.get_log('driver')[0]['message'] == \
                "Unable to evaluate script: disconnected: not connected to DevTools\n":
            exit_program(window, driver)
    else:
        window.after(1500, lambda: check_browser(window, driver))


def enable_print():
    sys.stdout = sys.__stdout__
    sys.stderr = sys.__stderr__


def disable_print():
    text_trap = io.StringIO()
    sys.stdout = text_trap
    sys.stderr = text_trap


def display_to_console(x):
    enable_print()
    print(x)
    disable_print()


def expect_by_id(driver, id_tag):
    # handle expected conditions by id
    wait_no_longer_than = 30
    print(f'Expecting {id_tag}')
    element = WebDriverWait(driver, wait_no_longer_than).until(
        EC.presence_of_element_located((By.ID, id_tag)))
    return element


def expect_by_XPATH(driver, XPATH):
    wait_no_longer_than = 30
    print(f'Expecting {XPATH}')
    element = WebDriverWait(driver, wait_no_longer_than).until(
        EC.presence_of_element_located((By.XPATH, XPATH)))
    return element


def expect_by_class(driver, class_tag):
    wait_no_longer_than = 30
    print(f'Expecting {class_tag}')
    element = WebDriverWait(driver, wait_no_longer_than).until(
        EC.presence_of_element_located((By.CLASS_NAME, class_tag)))
    return element


def expect_by_css(driver, css_tag):
    wait_no_longer_than = 30
    print(f'Expecting {css_tag}')
    element = WebDriverWait(driver, wait_no_longer_than).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css_tag)))
    return element


def expect_by_link_text(driver, link_text):
    wait_no_longer_than = 30
    print(f'Expecting {link_text}')
    element = WebDriverWait(driver, wait_no_longer_than).until(
        EC.presence_of_element_located((By.LINK_TEXT, link_text)))
    return element


def get_turfs():
    # Read data from excel file into tuples
    fname = r"..\io\Input\Turf List.xlsx"
    df = pd.read_excel(fname, sheet_name="Sheet1")
    turfs = []
    count = 0
    for turf in df['Turf Name'].values:
        building = df['Building Name'].values[count]
        turfs.append((turf, building))
        count += 1
    return turfs


def get_entries():
    type_dict = {
        'Full' : "This is a list of all the targeted voters in your turf.  They are high scoring Democrats and NPAs.",
        'VBM' : "This is a list of all voters in your turf who have already registered to Vote by Mail.  We want to encourage them to return their ballot as soon as possible.  We'd also like to encourage them to volunteer.",
        'non-VBM' : "This is a list of all voters in your turf who have NOT registered to Vote by Mail.  We want to encourage them to sign up for VBM as soon as possible.",
        "Inc" : "This is a list of Inconsistent voters in your turf.  They did not vote in August or in the 2018, or 2016 election.  We want to encourage them to vote."
    }
    # Had to use full path to get it to work for me.
    fname = r"C:\Users\Grant\Desktop\macrovan\io\Input\Nov 2020 -Tracking All Voters.xlsx"
    df = pd.read_excel(fname, sheet_name="Sheet1")
    turfs = []
    count = 0
    # todo: fix count and unused turf iterator
    for turf in df['Organizer'].values:
        organizer = df['Organizer'].values[count]
        if organizer == "STOP":
            break
        if not pd.isnull(organizer):
            name = df['suffix'].values[count]
            name_split = name.split(" ")
            bc_name = df['BC Name'].values[count]
            first_name = name_split[0]
            if(len(name_split) > 1):
                last_name = name_split[1]
            else:
                last_name = ""
            turf_name = df['Name in VAN'].values[count]
            pdf_type = df['suffix 2'].values[count]
            building = "test"
            # building = df['Bldg Name'].values[count]
            bc_email_address = df['BC Email'].values[count]
            email_address = df['Email to:'].values[count]
            if not pd.isnull(name) and not pd.isnull(email_address) and not pd.isnull(turf_name) and not pd.isnull(building) and not pd.isnull(
                    organizer) and not pd.isnull(bc_name) and not pd.isnull(bc_email_address) and not pd.isnull(pdf_type):
                turfs.append({
                    "first_name" : first_name,
                    "last_name" : last_name,
                    "email_address" : email_address,
                    "bc_name" : bc_name,
                    "bc_email_address" : bc_email_address,
                    "organizer_email_address" : organizer,
                    "turf_name" : turf_name,
                    "building_name" : building,
                    "message" : type_dict[pdf_type]
                })
        count += 1
    return turfs


def get_fnames(path):
    # Get all the PDF filenames.
    pdf_files = []
    for filename in os.listdir(path):
        #print(filename)
        if filename.endswith('.pdf'):
            pdf_files.append(filename)
    pdf_files.sort(key=str.lower)
    #print(pdf_files)
    return pdf_files


def write_excel(path, list_dict):
    """export to excel worksheet"""
    # df = pd.DataFrame({'List Name': [element[0] for element in lists],
    #                    'List Number': [element[1] for element in lists],
    #                    'Doors': [element[2] for element in lists]
    #                    })
    df = pd.DataFrame(list_dict).transpose()
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='List Numbers', index=False)
    writer.save()


def extract_list_info(path=r'io\Output'):
    # Loop through all the PDF files.
    #path = r'io\Output'
    print(path)
    pdf_files = get_fnames(path)
    list_dict = {}
    for filename in pdf_files:
        #pdfFileObj = open(r'io\Output\\' + filename, 'rb')
        pdfFileObj = open(path + '\\' + filename, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        page = pdfReader.getPage(0).extractText()
        first_part, doors = page.split("Doors:", 1)
        date, people = page.split("People:", 1)
        date = date.split("Generated")[1]
        date = date.split(" ")[1]
        doors = doors.split("Affiliation")[0]
        people = people.split("Affiliation")[0].split()[0]
        page = pdfReader.getPage(2).extractText()
        lname, lnum = page.split("List", 1)
        lnum = lnum.split(" ")[1]
        lname_s = lname.split(" ")
        lname = lname_s[0] + " " + lname_s[3] + " " + lname_s[4]
        list_dict[lname] = {
            'list_number' : lnum,
            'door_count' : doors,
            'person_count' : people,
            'date_generated' : date,
            'turf_name' : lname
        }
    return list_dict
