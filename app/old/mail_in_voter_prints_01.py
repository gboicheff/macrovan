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

    test_dict = {
        ("Barbara", "Law"): [("P 144 Turf 02", "Prendergast")],

        ("Andy", "Bragg"): [("P130 Upper Downtown Turf 07", "O'Connor")]
    }

    """
    test_dict = {
        ("Barbara", "Law"): [("P 144 Turf 02", "Prendergast"),  ("P 602 Turf 02", "Baker"), ("P138Hunter", "Hunter"),
    ("P 150 Turf 02", "Thurmond"), ("P 156 Turf 02", "Fling"), ('P 121Pincus', 'Pincus'), ('P 150 Turf 03','Davis'),
    ('P 155Tur', 'Tur'), ('P155Tur2', 'Tur')],

        ("Andy", "Bragg"): [("P130 Upper Downtown Turf 07", "O'Connor"), ("P130 Upper Downtown Turf 08", "Gardner"), ("P130 Upper Downtown Turf 09", "Warren"),
    ("P130 Upper Downtown Turf 10", "Strickler"), ("P130 Upper Downtown Turf 11", "Roux"), ("P130 Upper Downtown Turf 12", "Taylor"),
    ("P130 Upper Downtown Turf 13", "Grebenschikoff"), ("P130 Upper Downtown Turf 14", "Singer")]
    }
    """

    captains = get_all_captains(test_dict)

    """
    Remaining problems:
    sometimes already logged in so login_to_page fails. Have login status check.
    note field misses click
    print icon click fail
    anyone who requested (so early voting twisty i think)
    """

    script_name = "*2020 D68 Aug Primary"

    for captain in captains:
        turfs = get_turfs_by_captain(captain, test_dict)
        captain_name = captain[1] + ", " + captain[0]
        for turf in turfs:
            turf_name = turf[0]
            list_name = turf[0] + " " + turf[1]
            turfselection_plus(driver, turf_name, captain_name)

            print_controller(driver, list_name, script_name)
            driver.implicitly_wait(15)
            return_to_home(driver)
            driver.implicitly_wait(15)
            list_folders(driver)
            select_folder(driver)
            # pause("Click Ok to continue")

            #Select Print Format Option??