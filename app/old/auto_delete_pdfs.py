from utils import *
import time

if __name__ == '__main__':
    # teardown()
    driver = start_driver()
    get_page(driver)
    driver.implicitly_wait(10)
    login_to_page(driver)
    driver.implicitly_wait(30)
    expect_by_id(driver, "ctl00_ContentPlaceHolderVANPage_pdfListHyperLink").click()

    #run while loop until no deletes are left
    deletes_left = True
    while deletes_left:
        if driver.find_element(By.LINK_TEXT, "Delete"):
            try:
                time.sleep(1.5)
                expect_by_link_text(driver, "Delete").click()
                handle_alert(driver)
            except:
                print("exception thrown. refreshing")
                driver.refresh()
                time.sleep(5)
        else:
            driver.refresh()
            if driver.find_element(By.LINK_TEXT, "Delete"):
                continue
            else:
                deletes_left = False