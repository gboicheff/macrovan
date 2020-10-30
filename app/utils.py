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
import re
import fnmatch

pd.set_option('display.max_rows', 1000)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

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
    wait_no_longer_than = 120
    print(f'Expecting {id_tag}')
    element = WebDriverWait(driver, wait_no_longer_than).until(
        EC.presence_of_element_located((By.ID, id_tag)))
    return element


def expect_by_XPATH(driver, XPATH):
    wait_no_longer_than = 120
    print(f'Expecting {XPATH}')
    element = WebDriverWait(driver, wait_no_longer_than).until(
        EC.presence_of_element_located((By.XPATH, XPATH)))
    return element


def expect_by_class(driver, class_tag):
    wait_no_longer_than = 120
    print(f'Expecting {class_tag}')
    element = WebDriverWait(driver, wait_no_longer_than).until(
        EC.presence_of_element_located((By.CLASS_NAME, class_tag)))
    return element


def expect_by_css(driver, css_tag):
    wait_no_longer_than = 120
    print(f'Expecting {css_tag}')
    element = WebDriverWait(driver, wait_no_longer_than).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css_tag)))
    return element


def expect_by_link_text(driver, link_text):
    wait_no_longer_than = 120
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


def get_volunteer_data(fname=r"Nov 2020 -Tracking All Voters.xlsx",
                       sheet_name="To Deliver - Reports"):
    # Had to use full path to get it to work for me.
    # #print('Path string in get_volunteer_data = ', path)
    print(f'get_volunteer_data: os.getcwd = {os.getcwd()}')
    fname = os.path.join(os.getcwd(), "io", "Input", fname)
    print(fname)
    df = pd.read_excel(fname, sheet_name)
    print(f'df.keys = {df.keys()}')
    #print(f"df['Organizer Email'] = \n {df['Organizer Email'}")
    df['Organizer Email'] = df['Organizer Email'].str.lower()
    df['Organizer Email'] = df['Organizer Email'].str.lstrip()
    # df = df.sort_values('Organizer Email')

    volunteer_data = []
    count = 0
    # todo: fix count and unused turf iterator
    for org_email in df['Organizer Email'].values:

        # Ugly but gets the job done. Would be cleaner with a function:
        if 'Send to BC' in df.columns:
            email_to_bc = str(df['Send to BC'].values[count]).lower()
            #print(f'email_to_bc[0] = {email_to_bc[0]}')
            email_to_bc = email_to_bc[0]
        else:
            email_to_bc = ''

        if 'Zip to Organizer' in df.columns:
            zip_to_org = str(df['Zip to Organizer'].values[count]).lower()
            zip_to_org = zip_to_org[0]
            #print(zip_to_org)
        else:
            zip_to_org = ''

        if 'Want door hangers' in df.columns:
            want_door_hangers = str(df['Want door hangers'].values[count]).lower()
            #print('email_to_bc[0] = ', email_to_bc[0])
            want_door_hangers = want_door_hangers[0]
        else:
            want_door_hangers = ''

        if 'Organizer Email' in df.columns:
            organizer_email = df['Organizer Email'].values[count]
            if organizer_email != organizer_email:
                print('organizer_email = nan')
                organizer_email = ''
        else:
            organizer_email = ''

        if 'Org Phone' in df.columns:
            organizer_phone = str(df['Org Phone'].values[count])
        else:
            organizer_phone = ''

        if 'Org Name' in df.columns:
            organizer_name = df['Org Name'].values[count]
        else:
            organizer_name = ''

        if 'BC First Name' in df.columns:
            first_name = df['BC First Name'].values[count]
        else:
            first_name = ''

        if 'BC Last Name' in df.columns:
            last_name = df['BC Last Name'].values[count]
        else:
            last_name = ''

        if 'Name in VAN' in df.columns:
            #print("df['Name in VAN'].values[count] = ", df['Name in VAN'].values[count])
            turf_name_in_van = df['Name in VAN'].values[count]
        else:
            print('No Name in VAN')
            turf_name_in_van = ''

        if 'Total Voters' in df.columns:
            total_voters = int(df['Total Voters'].values[count])
        else:
            total_voters = ''

        if 'BC Email' in df.columns:
            bc_email_address = df['BC Email'].values[count]
        else:
            bc_email_address = ''

        volunteer_data.append({
            "email_to_bc": email_to_bc,
            "zip_to_org": zip_to_org,
            "want_door_hangers": want_door_hangers,
            "first_name": str(first_name),
            "last_name": str(last_name),
            "email_address": bc_email_address,
            "organizer_email_address": organizer_email,
            "organizer_phone": organizer_phone,
            "organizer_name": str(organizer_name),
            "turf_name_in_van": turf_name_in_van,
            "total_voters": total_voters
        })
        count += 1
    return volunteer_data


def get_organizer_turfs_dict(fname):
    volunteer_data = get_volunteer_data(fname)
    volunteer_dict = {}
    for turf in volunteer_data:
        turf_name_in_van = turf["turf_name_in_van"]
        organizer_email = turf["organizer_email_address"]
        volunteer_dict[turf_name_in_van] = organizer_email
        #print('turf_name_in_van = ' + turf_name_in_van + ', organizer_email = ' + organizer_email)
    return volunteer_dict


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


def write_excel(path, df):
    """export to excel worksheet"""
    # todo: parameterize sheet_name
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='List Numbers', index=False)
    writer.save()


def extract_pdf_info(path=r'io\Output'):
    # Loop through all the PDF files.
    #path = r'io\Output'
    print('Path string = ', path)
    pdf_files = get_fnames(path)
    print('pdf_files = ', pdf_files)

    #volunteer_dict = get_organizer_turfs_dict()
    pdf_dict = {}

    for filename in pdf_files:
        #print('pdf filename = ', filename)
        #pdfFileObj = open(r'io\Output\\' + filename, 'rb')
        pdfFileObj = open(path + '\\' + filename, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        page = pdfReader.getPage(0).extractText()
        first_part, doors = page.split("Doors:", 1)
        date, people = page.split("People:", 1)
        date = date.split("Generated")[1]
        date = date.split(" ")[1]
        doors = int(doors.split("Affiliation")[0])
        people = int(people.split("Affiliation")[0].split()[0])
        page = pdfReader.getPage(2).extractText()
        # print('Page =', page)
        if people != 0:
            pdf_file_name, lnum = page.split("List", 1)
            lnum = lnum.split(" ")[1]
        else:
            lnum = '0-0'
            pdf_file_name, date_part = filename.split("_2020", 1)
            # print(filename, '\n', pdf_file_name, '\n', date_part, '\n', page, '\n')
            #exit(2)

        # print('pdf_file_name = ' + pdf_file_name)
        pdf_dict[pdf_file_name] = {
            'list_number': lnum,
            'door_count': doors,
            'person_count': people,
            'date_generated': date,
            'pdf_file_name': pdf_file_name,
            #'organizer_email': organizer_email
        }
    return pdf_dict


#iterate through folder_dict and create a subfolder copying the files over for each organizer
def create_folders(folder_dict, parent_folder_name):
    parent_path = r'D:\Stuff\Projects\Pol\macrovan\io\Output'
    print(f'create_folders: parent_path = {parent_path}')
    os.chdir(parent_path)
    print(f'create_folders: os.getcwd = {os.getcwd()}')
    if(os.path.isdir(parent_folder_name)):
        shutil.rmtree(parent_folder_name)
    os.mkdir(parent_folder_name)
    os.chdir(parent_folder_name)
    print(f'create_folders: os.getcwd = {os.getcwd()}')
    print(f'create_folders: parent_folder_name = {parent_folder_name}')
    print(f'create_folders: folder_dict keys =\n {folder_dict.keys()}')
    for key in folder_dict.keys():
        subfolder = key
        if(os.path.isdir(subfolder)):
            print('ISDIR: Subfolder Exists')
            print(f'os.getcwd() = {os.getcwd()}\n Subfolder = {subfolder} EXISTS')
            continue
        else:
            os.mkdir(subfolder)
        os.chdir(subfolder)
        #print(f'chdir(subfolder) os.getcwd() = {os.getcwd()}')
        for file in folder_dict[subfolder]:
            search_file = file + "*" + ".pdf"
            search_file = search_file.replace(" ", "")
            file_found = 'No'
            for file in os.listdir(parent_path+r"\tests"):
                found_file = file.replace(" ", "")
                #print('search_file = ', search_file)
                #print('found_file = ', found_file)
                if fnmatch.fnmatch(found_file, search_file):
                    #print('Matched files', found_file)
                    #print()
                    #shutil.copy(parent_path+r"\app\io\output\\tests\"+file, file)
                    shutil.copy(parent_path + r'\\tests\\' + file, file)
                    file_found = 'Yes'
                    break
            if file_found != 'Yes':
                print(f"For Organizer: {subfolder}\nWARNING SEARCH FILE {search_file} NOT FOUND!")
        os.chdir("..")
    os.chdir(parent_path)

def create_organizer_folders(fname, sheet_name):
    organizerFiles = {}
    turfs = get_volunteer_data(fname, sheet_name)
    for turf in turfs:
        turf_name = turf['turf_name_in_van']
        organizer_email = turf['organizer_email_address']
        # PRINT to show emails flagged TO SEND TO ORG
        print(f"create_organizer_folders: For Org = {organizer_email}, zip_to_org = {turf['zip_to_org']}"
              f"\n\tEMAIL SENT for turf_name = {turf_name}")
        if turf['zip_to_org'] == 'y':
            if organizer_email != organizer_email:
                organizer_email = 'no_org'
            filename = turf_name
            if organizer_email in organizerFiles:
                organizerFiles[organizer_email] += [filename]
            else:
                organizerFiles[organizer_email] = [filename]
        else:
            # PRINT to show emails flagged DO NOT SEND TO ORG
            print(f"create_organizer_folders: For Org = {organizer_email}, zip_to_org = {turf['zip_to_org']}"
                  f"\n\tSo EMAIL ABORTED for turf_name = {turf_name}!")
            continue
    create_folders(organizerFiles, "Organizers")
