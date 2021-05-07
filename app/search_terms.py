from utils import *


path = os.getcwd()
print(f"The current working directory is {path}")

if __name__ == '__main__':

    teardown()
    driver = start_driver(os.path.join(path, "chrome-data"))
    get_page(driver)
    driver.implicitly_wait(10)
    attempt_login(driver)  
    list_folders(driver)
    select_folder(driver, '//*[text()="Testing"]')
    table = expect_by_tag(driver, "table")
    row = select_row(driver, table)
    tbody = expect_by_tag(row, "tbody")
    edit_button = expect_by_tag(row, "a")
    edit_button.click()