from utils import *


path = os.getcwd()
print(f"The current working directory is {path}")

if __name__ == '__main__':

    teardown()
    driver = start_driver(os.path.join(path, "chrome-data"))
    get_page(driver)
    driver.implicitly_wait(10)
    login_to_page(driver)  
    list_folders(driver)
    select_folder(driver)