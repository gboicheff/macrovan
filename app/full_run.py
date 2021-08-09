from search_terms import *
from turf_sheet_update import *
from precinct_sheet_update import *
from VAN_routines import *
if __name__ == "__main__":
    teardown()
    driver = start_driver(os.path.join(path, "chrome-data"))
    try:
        execute(driver, "Sheet VBM Precincts", "Search", "VBM_precincts")
    except Exception as e:
        print(e)

    time.sleep(672)
    
    try:
        execute(driver, "Sheet Precincts Voter Removed", "Search", "normal_precincts")
    except Exception as e:
        print(e)

    time.sleep(750)

    try:
        execute(driver, "**2021 Municipal St. Petersburg", "Map Turf", "turfs")
    except Exception as e:
        print(e)

    time.sleep(750)

    
    try:
        run_turfs()
    except Exception as e:
        print(e)

    # try:
    #     run_VBM_turfs()
    # except Exception as e:
    #     print(e)

    try:
        run_precincts_VBM()
    except Exception as e:
        print(e)

    try:
        run_precincts_normal()
    except Exception as e:
        print(e)
    print("---Done---")
