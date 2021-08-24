from utils import *
from printing_steps import *
import time

path = os.getcwd()

def read_vbm_excel():
    # Read VBM turfs to be printed from excel file
    fname = r"..\io\Input\Nov 2020 Tracking Non voters.xlsx"
    # fname = r"..\io\Input\Nov 2020 -Tracking All Voters.xlsx"
    df = pd.read_excel(fname, sheet_name="Print Reports")
    # df = pd.read_excel(fname, sheet_name="Ready to run Reports")
    return df

def get_vbm_turfs(df):
    turfs = []
    count = 0
    for turf in df['Name in VAN'].values:
        suffix = df['suffix'].values[count]
        print_type = df['Label/List'].values[count]
        print_script = df['Script'].values[count]
        turfs.append((turf, suffix, print_type, print_script))
        count += 1
    return turfs

def print_labels(driver, list_name):
    #Click on Labels from My List page
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_PrintLabelsLabel").click()
    #Send Name
    element = expect_by_id(driver,
                           "ctl00_ContentPlaceHolderVANPage_VANDetailsItemReportTitle_VANInputItemDetailsItemReportTitle_ReportTitle")
    element.clear()
    element.send_keys(list_name)
    print()
    print(f'list_name: {list_name}')
    time.sleep(1)
    #Select format
    select_element = Select(expect_by_id(driver,
                                         "ctl00_ContentPlaceHolderVANPage_VanDetailsItemReportFormatInfo_VANInputItemDetailsItemReportFormatInfo_ReportFormatInfo"))
    select_element.select_by_visible_text("New Standard Mailing Labels (Avery 5160)")
    select_element = Select(expect_by_id(driver,
                                         "ctl00_ContentPlaceHolderVANPage_VanDetailsItemHouseholding_VANInputItemDetailsItemHouseholding_Householding"))
    select_element.select_by_visible_text("Print one label per Address")
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_ButtonSortOptionsSubmit").click()
    expect_by_link_text(driver, "My PDF Files").click()

def replace_characters(list_name):
    chars_to_remove = ['.', '\'']
    list_name_new = list_name
    for char in chars_to_remove:
        if char in list_name_new:
            list_name_new = list_name_new.replace(char, ' ')
    return list_name_new


if __name__ == '__main__':
    # teardown()
    driver = start_driver()
    get_page(driver)
    driver.implicitly_wait(10)
    login_to_page(driver)
    list_folders(driver)
    select_folder(driver)

    data = read_vbm_excel()
    turfs = get_vbm_turfs(data)

    for turf in turfs:
        turf_name = turf[0]

        if (type(turf_name) != type('str')):
            continue
        else:
            if (type(turf[1]) == type('str')):
                print_list_name = turf[0] + " " + turf[1]
            else:
                print_list_name = turf[0]
            print_list_name = print_list_name.translate(print_list_name.maketrans('', '', "'.!*\\"))
            script_name = turf[3]

            select_turf(driver, turf_name)
            handle_alert(driver)

            if (turf[2] == "List"):
                print_controller(driver, print_list_name, script_name)
                time.sleep(2)
                return_to_home(driver)
                time.sleep(1)
                list_folders(driver)
                select_folder(driver)
            else:
                print_labels(driver, print_list_name)
                driver.implicitly_wait(15)
                return_to_home(driver)
                driver.implicitly_wait(15)
                list_folders(driver)
                select_folder(driver)