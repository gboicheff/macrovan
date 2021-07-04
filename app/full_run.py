from search_terms import *
from turf_sheet_update import *
from precinct_sheet_update import *

if __name__ == "__main__":
    teardown()
    driver = start_driver(os.path.join(path, "chrome-data"))
    execute(driver, "Sheet Precincts", "Search", "precincts")
    execute(driver, "**2021 Municipal St. Petersburg", "Map Turf", "turfs")
    run_turfs()
    run_precincts()
    print("---Done---")
