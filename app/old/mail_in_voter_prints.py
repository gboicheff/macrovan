from utils import *
from printing_steps import *

path = os.getcwd()
print(f"The current working directory is {path}")

if __name__ == '__main__':

    teardown()
    driver = start_driver()
    get_page(driver)
    driver.implicitly_wait(10)
    login_to_page(driver)  # todo: check if really need to or it might hang?
    list_folders(driver)
    select_folder(driver)

    turfs = get_turfs()

    for turf in turfs:
        turf_name = turf[0]
        print_list_name = turf[0] + " " + turf[1]
        script_name = "*2020 D68 Aug Primary"
        select_turf(driver, turf_name)
        handle_alert(driver)
        print_controller(driver, print_list_name, script_name)
        driver.implicitly_wait(15)
        return_to_home(driver)
        driver.implicitly_wait(15)
        list_folders(driver)
        select_folder(driver)