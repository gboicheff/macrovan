from utils import *
from bs4 import BeautifulSoup
import time
path = os.getcwd()
print(f"The current working directory is {path}")

if __name__ == '__main__':

    teardown()
    driver = start_driver(os.path.join(path, "chrome-data"))
    get_page(driver)
    driver.implicitly_wait(10)
    attempt_login(driver)  
    list_folders(driver)
    select_folder(driver, '//*[text()="*2021 Municipal St. Petersburg Master"]')
    table = expect_by_tag(driver, "table")


    row = select_row(driver, table)
    button = expect_by_XPATH(expect_by_tag(row, "tbody"), '//*[@id="ctl00_ContentPlaceHolderVANPage_gvList"]/tbody/tr[1]/td[4]')
    edit_button = expect_by_tag(button, "a")
    edit_button.click()

    # accept alert
    time.sleep(5)
    driver.switch_to.alert.accept()

    button = expect_by_XPATH(driver, '//*[@id="addStep"]')
    button.click()
    inner_button = expect_by_XPATH(button, '//*[@id="editSearchOption"]')
    inner_button.click()

    # select new table
    table = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_AllPageSectionsPanelHolder"]')

    rows = table.find_elements(By.CLASS_NAME, "OpenClosePSButton")
    for row in rows:
        src = row.get_attribute("src")
        if "right" in src:
            time.sleep(1)
            try:
                row.click()
            except Exception:
                pass
    
    # parse the selected options
    page_src = driver.page_source
    soup = BeautifulSoup(page_src, 'html.parser')

    with open("src.txt", "w", encoding="utf-8") as file:
        file.write(str(soup))

