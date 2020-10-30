from utils import *
import time

def open_print(driver):
    # Click the Print List Button
    #wait_no_longer_than = 30
    print('in print_list waiting for print icon')
    # element = WebDriverWait(driver, wait_no_longer_than).until(
    #     EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolderVANPage_HyperLinkImagePrintReportsAndForms")))
    element = expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_HyperLinkImagePrintReportsAndForms")
    print('in print_list trying to click')
    # driver.find_element(By.ID, "ctl00_ContentPlaceHolderVANPage_HyperLinkImagePrintReportsAndForms").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_HyperLinkImagePrintReportsAndForms").click()
    print('just clicked print icon might need another EC')

def top_selections(driver, listName, script_name):
    # Select Report Format Option
    # Locate the Sector and create a Select object
    print('Select Print Format Option')
    # select_element = Select(driver.find_element_by_id(
    #     'ctl00_ContentPlaceHolderVANPage_VanDetailsItemReportFormatInfo_VANInputItemDetailsItemReportFormatInfo_ReportFormatInfo'
    #     ))
    select_element = Select(expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemReportFormatInfo_VANInputItemDetailsItemReportFormatInfo_ReportFormatInfo"))
    element = select_element.select_by_visible_text(script_name)
    # Select Script Option
    # Locate the Sector and create a Select object
    # select_element = Select(driver.find_element_by_id(
    #     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemvdiScriptID_VANInputItemDetailsItemActiveScriptID_ActiveScriptID"
    #     )
    #                         )
    select_element = Select(expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemvdiScriptID_VANInputItemDetailsItemActiveScriptID_ActiveScriptID"))
    element = select_element.select_by_visible_text(script_name)

    # driver.find_element(By.ID,
    #                     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemvdiScriptID_VANInputItemDetailsItemActiveScriptID_ActiveScriptID").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemvdiScriptID_VANInputItemDetailsItemActiveScriptID_ActiveScriptID").click()

    # Script source selection (Walk)
    # driver.find_element(By.ID,
    #                     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemVANDetailsItemScriptSource_ScriptSource_VANInputItemDetailsItemScriptSource_ScriptSource").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemVANDetailsItemScriptSource_ScriptSource_VANInputItemDetailsItemScriptSource_ScriptSource").click()
    # dropdown = driver.find_element(By.ID,
    #                                "ctl00_ContentPlaceHolderVANPage_VanDetailsItemVANDetailsItemScriptSource_ScriptSource_VANInputItemDetailsItemScriptSource_ScriptSource")
    dropdown = expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemVANDetailsItemScriptSource_ScriptSource_VANInputItemDetailsItemScriptSource_ScriptSource")
    # dropdown.find_element(By.XPATH, "//option[. = 'Walk']").click()
    expect_by_XPATH(dropdown, "//option[. = 'Walk']").click()
    # driver.find_element(By.ID,
    #                     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemVANDetailsItemScriptSource_ScriptSource_VANInputItemDetailsItemScriptSource_ScriptSource").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemVANDetailsItemScriptSource_ScriptSource_VANInputItemDetailsItemScriptSource_ScriptSource").click()

    # Enter List Name
    # element = driver.find_element_by_id(
    #     "ctl00_ContentPlaceHolderVANPage_VANDetailsItemReportTitle_VANInputItemDetailsItemReportTitle_ReportTitle")
    element = expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VANDetailsItemReportTitle_VANInputItemDetailsItemReportTitle_ReportTitle")
    element.clear()
    element.send_keys(listName)

def headers_and_pagebreaks(driver):
    # Deselect headers and page breaks for sort order 1-4
    # driver.find_element(By.ID, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder1_Header1").click()
    # driver.find_element(By.ID, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder1_Break1").click()
    # driver.find_element(By.ID, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder2_Header2").click()
    # driver.find_element(By.ID, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder2_Break2").click()
    # driver.find_element(By.ID, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder3_Break3").click()
    # driver.find_element(By.ID, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder3_Header3").click()
    # driver.find_element(By.ID, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_Header4").click()
    # driver.find_element(By.ID, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_Break4").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder1_Header1").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder1_Break1").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder2_Header2").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder2_Break2").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder3_Header3").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder3_Break3").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_Header4").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_Break4").click()

def sort_orders(driver):
    # Sort Order 4
    # driver.find_element(By.ID,
    #                     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_VANInputItemDetailsItemSortOrder4_SortOrder4").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_VANInputItemDetailsItemSortOrder4_SortOrder4").click()
    # dropdown = Select(driver.find_element_by_id(
    #     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_VANInputItemDetailsItemSortOrder4_SortOrder4"))
    dropdown = Select(expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder4_VANInputItemDetailsItemSortOrder4_SortOrder4"))
    dropdown.select_by_index(4)

    # Sort Order 5
    # driver.find_element(By.ID,
    #                     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder5_VANInputItemDetailsItemSortOrder5_SortOrder5").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder5_VANInputItemDetailsItemSortOrder5_SortOrder5").click()
    # dropdown = Select(driver.find_element_by_id(
    #     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder5_VANInputItemDetailsItemSortOrder5_SortOrder5"))
    dropdown = Select(expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder5_VANInputItemDetailsItemSortOrder5_SortOrder5"))
    dropdown.select_by_index(5)

    # Sort Order 6
    # driver.find_element(By.ID,
    #                     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder6_VANInputItemDetailsItemSortOrder6_SortOrder6").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder6_VANInputItemDetailsItemSortOrder6_SortOrder6").click()
    # dropdown = Select(driver.find_element_by_id(
    #     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder6_VANInputItemDetailsItemSortOrder6_SortOrder6"))
    dropdown = Select(expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemSortOrder6_VANInputItemDetailsItemSortOrder6_SortOrder6"))
    dropdown.select_by_index(0)
    # driver.find_element(By.ID,
    #                     "ctl00_ContentPlaceHolderVANPage_VanDetailsItemPrintMapNew_VANInputItemDetailsItemPrintMapNew_PrintMapNew_0").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_VanDetailsItemPrintMapNew_VANInputItemDetailsItemPrintMapNew_PrintMapNew_0").click()

def final_selections_submit(driver):
    # pause("Double Check Selections. Then Press Okay.")
    # driver.find_element(By.ID, "ctl00_ContentPlaceHolderVANPage_ButtonSortOptionsSubmit").click()
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_ButtonSortOptionsSubmit").click()
    # driver.find_element(By.LINK_TEXT, "My PDF Files").click()
    expect_by_link_text(driver, "My PDF Files").click()

def print_controller(driver, listName, script_name):
    #Print a List
    open_print(driver)
    top_selections(driver, listName, script_name)
    # pause('click ok when done')
    headers_and_pagebreaks(driver)
    # pause('click ok when done')
    sort_orders(driver)
    driver.implicitly_wait(20)
    time.sleep(1.5)
    final_selections_submit(driver)

#     #ACTION ID BUTTON FAIL
