from utils import *
# from bs4 import BeautifulSoup
import time
import pickle
import random
import sys
import os
import re
from selenium.webdriver.support.select import Select
import tkinter as tk
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from VAN_utils import *
from search_terms import *
path = os.getcwd()
print(f"The current working directory is {path}")



def step_1_mun_2020(driver, precinct_num):
    home_districts_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelDistrictsNarrow"]/div[1]/div/label')
    home_districts_button.click()

    county_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_CountyID"]')
    count_dropdown_select = Select(county_dropdown)
    count_dropdown_select.select_by_visible_text('Pinellas')

    sleep_random_time(2)

    precinct_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_PrecinctID"]')
    precinct_dropdown_select = Select(precinct_dropdown)
    precinct_dropdown_select.select_by_visible_text(str(precinct_num))


    party_dropdown_button = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionParty"]')
    party_dropdown_button.click()
    democrat_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[1]')
    toggle(democrat_button, True)


def step_2_mun_2020(driver, precinct_num):
    add_step(driver, True)


    party_dropdown_button = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionParty"]')
    party_dropdown_button.click()

    ind_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[3]')
    toggle(ind_button, True)

    no_party_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[4]')
    toggle(no_party_button, True)

    other_party_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[5]')
    toggle(other_party_button, True)


    home_districts_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelDistrictsNarrow"]/div[1]/div/label')
    home_districts_button.click()

    county_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_CountyID"]')
    count_dropdown_select = Select(county_dropdown)
    count_dropdown_select.select_by_visible_text('Pinellas')

    sleep_random_time(2)

    precinct_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_PrecinctID"]')
    precinct_dropdown_select = Select(precinct_dropdown)
    precinct_dropdown_select.select_by_visible_text(str(precinct_num))


    scores_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelScoring"]/div[1]/div/label')
    scores_button.click()

    input1 = expect_by_XPATH(driver, '//*[@id="Scoring_DNCDemPartySupportV22020_From"]')
    input2 = expect_by_XPATH(driver, '//*[@id="Scoring_DNCDemPartySupportV22020_To"]')

    input1.clear()
    input1.send_keys(60)

    input2.clear()
    input2.send_keys(100)


    default_voter_status(driver)


def step_3_mun_2020(driver, precinct_num):
    add_step(driver, False)

    notes_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelNotes"]/div[1]/div/label')
    notes_button.click()

    notes_input = expect_by_XPATH(driver, '//*[@id="NoteText"]')
    notes_input.clear()
    notes_input.send_keys("*moved")


    default_voter_status(driver)



def step_4_mun_2020(driver, precinct_num):
    add_step(driver, False)
    canvas_status_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelCanvassStatus"]/div[1]/div/label')
    canvas_status_button.click()


    canvas_dropdown = expect_by_XPATH(driver, '//*[@id="CanvassIncludeExclude"]')
    canvas_dropdown_select= Select(canvas_dropdown)
    canvas_dropdown_select.select_by_visible_text("Include Only")

    sleep_random_time(0.25)

    moved_button = expect_by_XPATH(driver, '//*[@id="ResultID_5"]')
    toggle(moved_button, True)


    default_voter_status(driver)



def step_5_mun_2020(driver, precinct_num):
    add_step(driver, False)

    canvas_status_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelCanvassStatus"]/div[1]/div/label')
    canvas_status_button.click()


    canvas_dropdown = expect_by_XPATH(driver, '//*[@id="CanvassIncludeExclude"]')
    canvas_dropdown_select= Select(canvas_dropdown)
    canvas_dropdown_select.select_by_visible_text("Include Only")

    sleep_random_time(0.25)

    moved_button = expect_by_XPATH(driver, '//*[@id="ResultID_2"]')
    toggle(moved_button, True)


    default_voter_status(driver)



def step_6_mun_2020(driver, precinct_num):
    add_step(driver, False)



    survey_question_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelSurveyQuestions"]/div[1]/div/label')
    survey_question_button.click()

    survey_dropdown = expect_by_XPATH(driver, '//*[@id="PanelSectionSurveyQuestions"]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/select')
    survey_dropdown_select = Select(survey_dropdown)

    survey_dropdown_select.select_by_visible_text('2022 Affiliation: Party Affiliation (Public)')

    lean_rep_button = expect_by_XPATH(driver, '//*[@id="SurveyResponseIDs_1786083"]')
    toggle(lean_rep_button, True)





    default_voter_status(driver)


def step_7_mun_2020(driver, precinct_num):
    add_step(driver, False)



    survey_question_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelSurveyQuestions"]/div[1]/div/label')
    survey_question_button.click()

    survey_dropdown = expect_by_XPATH(driver, '//*[@id="PanelSectionSurveyQuestions"]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/select')
    survey_dropdown_select = Select(survey_dropdown)

    survey_dropdown_select.select_by_visible_text('2022 Affiliation: Party Affiliation (Public)')

    strong_rep_button = expect_by_XPATH(driver, '//*[@id="SurveyResponseIDs_1786084"]')
    toggle(strong_rep_button, True)


    default_voter_status(driver)




def new_step_1_mun_2020(driver, precinct_num):
    home_districts_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelDistrictsNarrow"]/div[1]/div/label')
    home_districts_button.click()

    county_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_CountyID"]')
    count_dropdown_select = Select(county_dropdown)
    count_dropdown_select.select_by_visible_text('Pinellas')

    sleep_random_time(0.5)

    precinct_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_PrecinctID"]')
    precinct_dropdown_select = Select(precinct_dropdown)
    precinct_dropdown_select.select_by_visible_text(str(precinct_num))


    party_dropdown_button = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionParty"]')
    party_dropdown_button.click()
    democrat_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[1]')
    toggle(democrat_button, True)


def new_step_2_mun_2020(driver, precinct_num):
    add_step(driver, True)

    activist_codes = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionActivistCodes"]')
    activist_codes.click()

    code_input = expect_by_XPATH(driver, '//*[@id="PanelSectionActivistCodes"]/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/select')
    
    code_input_select = Select(code_input)
    code_input_select.select_by_visible_text('Activist: VoteBadAddress2021 (Public)')

    home_districts_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelDistrictsNarrow"]/div[1]/div/label')
    home_districts_button.click()

    county_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_CountyID"]')
    count_dropdown_select = Select(county_dropdown)
    count_dropdown_select.select_by_visible_text('Pinellas')

    sleep_random_time(0.5)

    precinct_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_PrecinctID"]')
    precinct_dropdown_select = Select(precinct_dropdown)
    precinct_dropdown_select.select_by_visible_text(str(precinct_num))


    party_dropdown_button = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionParty"]')
    party_dropdown_button.click()
    democrat_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[1]')
    toggle(democrat_button, True)


    suppressions = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionSuppressions"]')
    suppressions.click()


    sleep_random_time(0.5)
    
    remove_all_sups_button = expect_by_XPATH(driver, '//*[@id="RemoveAllSuppressions"]')
    remove_all_sups_button.click()

    sleep_random_time(0.5)


def new_step_3_mun_2020(driver, precinct_num):
    add_step(driver, True)


    party_dropdown_button = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionParty"]')
    party_dropdown_button.click()

    ind_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[3]')
    toggle(ind_button, True)

    no_party_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[4]')
    toggle(no_party_button, True)

    other_party_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[5]')
    toggle(other_party_button, True)


    home_districts_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelDistrictsNarrow"]/div[1]/div/label')
    home_districts_button.click()

    county_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_CountyID"]')
    count_dropdown_select = Select(county_dropdown)
    count_dropdown_select.select_by_visible_text('Pinellas')

    sleep_random_time(0.5)

    precinct_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_PrecinctID"]')
    precinct_dropdown_select = Select(precinct_dropdown)
    precinct_dropdown_select.select_by_visible_text(str(precinct_num))


    scores_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelScoring"]/div[1]/div/label')
    scores_button.click()

    input1 = expect_by_XPATH(driver, '//*[@id="Scoring_DNCDemPartySupportV22020_From"]')
    input2 = expect_by_XPATH(driver, '//*[@id="Scoring_DNCDemPartySupportV22020_To"]')

    input1.clear()
    input1.send_keys(60)

    input2.clear()
    input2.send_keys(100)


    sleep_random_time(0)

    # default_voter_status(driver)


def new_step_4_mun_2020(driver, precinct_num):
    add_step(driver, True)


    activist_codes = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionActivistCodes"]')
    activist_codes.click()

    code_input = expect_by_XPATH(driver, '//*[@id="PanelSectionActivistCodes"]/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/select')
    
    code_input_select = Select(code_input)
    code_input_select.select_by_visible_text('Activist: VoteBadAddress2021 (Public)')

    party_dropdown_button = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionParty"]')
    party_dropdown_button.click()

    ind_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[3]')
    toggle(ind_button, True)

    no_party_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[4]')
    toggle(no_party_button, True)

    other_party_button = expect_by_XPATH(driver, '//*[@id="PanelSectionParty"]/table/tbody/tr/td[2]/input[5]')
    toggle(other_party_button, True)


    home_districts_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelDistrictsNarrow"]/div[1]/div/label')
    home_districts_button.click()

    county_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_CountyID"]')
    count_dropdown_select = Select(county_dropdown)
    count_dropdown_select.select_by_visible_text('Pinellas')

    sleep_random_time(0.5)

    precinct_dropdown = expect_by_XPATH(driver, '//*[@id="Dist_PrecinctID"]')
    precinct_dropdown_select = Select(precinct_dropdown)
    precinct_dropdown_select.select_by_visible_text(str(precinct_num))


    scores_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelScoring"]/div[1]/div/label')
    scores_button.click()

    input1 = expect_by_XPATH(driver, '//*[@id="Scoring_DNCDemPartySupportV22020_From"]')
    input2 = expect_by_XPATH(driver, '//*[@id="Scoring_DNCDemPartySupportV22020_To"]')

    input1.clear()
    input1.send_keys(60)

    input2.clear()
    input2.send_keys(100)


    sleep_random_time(0)
    
    suppressions = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionSuppressions"]')
    suppressions.click()


    sleep_random_time(1)
    
    remove_all_sups_button = expect_by_XPATH(driver, '//*[@id="RemoveAllSuppressions"]')
    remove_all_sups_button.click()

    sleep_random_time(0.5)   


def new_step_5_mun_2020(driver, precinct_num):
    add_step(driver, False)

    notes_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelNotes"]/div[1]/div/label')
    notes_button.click()

    notes_input = expect_by_XPATH(driver, '//*[@id="NoteText"]')
    notes_input.clear()
    notes_input.send_keys("*moved")


    # default_voter_status(driver)


def new_step_6_mun_2020(driver, precinct_num):
    add_step(driver, False)
    canvas_status_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelCanvassStatus"]/div[1]/div/label')
    canvas_status_button.click()

    canvas_dropdown = expect_by_XPATH(driver, '//*[@id="CanvassIncludeExclude"]')
    canvas_dropdown_select= Select(canvas_dropdown)
    canvas_dropdown_select.select_by_visible_text("Include Only")

    sleep_random_time(0.25)

    moved_button = expect_by_XPATH(driver, '//*[@id="ResultID_5"]')
    toggle(moved_button, True)

    refused_button = expect_by_XPATH(driver, '//*[@id="ResultID_2"]')
    toggle(refused_button, True)


def new_step_7_mun_2020(driver, precinct_num):
    add_step(driver, False)



    survey_question_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelSurveyQuestions"]/div[1]/div/label')
    survey_question_button.click()

    survey_dropdown = expect_by_XPATH(driver, '//*[@id="PanelSectionSurveyQuestions"]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/select')
    survey_dropdown_select = Select(survey_dropdown)

    survey_dropdown_select.select_by_visible_text('2022 Affiliation: Party Affiliation (Public)')

    strong_rep_button = expect_by_XPATH(driver, '//*[@id="SurveyResponseIDs_1786084"]')
    toggle(strong_rep_button, True)

    lean_rep_button = expect_by_XPATH(driver, '//*[@id="SurveyResponseIDs_1786083"]')
    toggle(lean_rep_button, True)


    # default_voter_status(driver)

def new_step_8_mun_2020(driver, precinct_num):
    narrow(driver)

    early_voting_button = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionEarlyVoting"]')
    early_voting_button.click()

    R_toggle = expect_by_XPATH(driver, '//*[@id="BallotReturnStatusName_117"]')
    toggle(R_toggle, True)




def create_precincts(driver, precincts, edit_steps, target_folder_name):
    output_file = open("logfile.txt", "w")
    get_page(driver)
    driver.implicitly_wait(60)
    index = 0
    while index < len(precincts):
        precinct = precincts[index]
        try:
            home_button = expect_by_XPATH(driver, '//*[@id="wrapper"]/div[1]/div[1]/div/div[2]/a')
            home_button.click()

            create_new_list_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_HyperLinkMenuCreateANewList"]')
            create_new_list_button.click()

            for step in edit_steps:
                sleep_random_time(0.25)
                step(driver, precinct)
                sleep_random_time(0.25)
                preview_results(driver)
            
            run_search(driver)
            save_list(driver, "P{} {}".format(precinct, int(time.time())), target_folder_name)
        except Exception as e:
            print("Failed on P{} with error {}".format(precinct, e))
            output_file.write("Failed on P{} with error {}".format(precinct, e))
            sleep_random_time(8)
            pass
        else:
            output_file.write("Successfully created P{}".format(precinct))
            index+=1

    output_file.close()  




def get_VBM_turf_counts(driver, selected_folder, output_file_name, edit_steps=[new_step_8_mun_2020]):
    get_page(driver)
    driver.implicitly_wait(60)
    list_folders(driver)

    folder_name = '//*[text()="{}"]'.format(selected_folder)
    expect_by_XPATH(driver, folder_name).click()

    map_regions = []

    # find num rows
    try:
        num_rows = int(expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_gvList"]/tfoot/tr/td/table/tbody/tr/td[1]/b[1]', 5).text.split()[0])
    except TimeoutException:
        num_rows = int(expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_gvList"]/tfoot/tr/td/b[1]', 5).text.split()[0])

    # set the number of rows per page to 999
    settings_button = expect_by_XPATH(driver, '//*[@id="HyperLinkSettings"]')
    settings_button.click()
    rows_p_page_input = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_VANDetailsItemDefaultRows_VANInputItemDetailsItemDefaultRows_DefaultRows"]')

    rows_p_page_input.clear()
    rows_p_page_input.send_keys("999")
   

    save_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_ButtonSave"]')
    save_button.click()

    time.sleep(1.5)
    refresh_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_RefreshFilterButton"]')
    refresh_button.click()
    time.sleep(1)


    for index in range(1,num_rows+1):


        time.sleep(random.randint(1,3))
        row_xpath = '//*[@id="ctl00_ContentPlaceHolderVANPage_gvList"]/tbody/tr[{index}]'.format(index=index)

        # check for map turf
        row_type = expect_by_XPATH(driver, row_xpath + '/td[3]/span').text
        print(row_type)
        print("ROW_TYPE:{}".format(row_type))
        if (row_type == 'Map Turf'):               
                map_region = MapRegion()
                map_region.ID = expect_by_XPATH(driver, row_xpath + '/td[2]/span').text
                map_region.NAME = expect_by_XPATH(driver, row_xpath + '/td[4]/a/span').text
                
                # get stuff on edit page
                button = expect_by_XPATH(driver, row_xpath + '/td[4]')
                edit_button = expect_by_tag(button, "a")
                edit_button.click()

                time.sleep(0.25)
                driver.switch_to.alert.accept()



                edit_list(driver)
                
                narrow_people_button = expect_by_XPATH(driver, '//*[@id="stepTypeItem4"]')
                narrow_people_button.click()

                early_voting_button = expect_by_XPATH(driver, '//*[@id="ImageButtonSectionEarlyVoting"]')
                early_voting_button.click()

                R_toggle = expect_by_XPATH(driver, '//*[@id="BallotReturnStatusName_117"]')
                toggle(R_toggle, True)

                preview_button = expect_by_XPATH(driver, '//*[@id="ResultsPreviewButton"]')
                preview_button.click()
                time.sleep(2)

                plus_button = expect_by_XPATH(driver, '//*[@id="SearchSummaryButtons"]/div[1]/div/div/span[3]')
                plus_button.click()


                map_region.PEOPLE = expect_by_XPATH(driver, '//*[@id="SearchSummaryButtons"]/div[1]/div/div/label/span').text
                map_region.DOORS = expect_by_XPATH(driver, '//*[@id="SearchSummaryButtons"]/div[1]/div/div/div/ul/li[2]/label/span').text

                home_button = expect_by_XPATH(driver, '//*[@id="wrapper"]/div[1]/div[1]/div/div[2]/a')
                home_button.click()

                driver.switch_to.alert.accept()


                map_region.display()
                map_regions.append(map_region)

                time.sleep(2)
                list_folders(driver)
                folder_name = '//*[text()="{}"]'.format(selected_folder)
                expect_by_XPATH(driver, folder_name).click()

    if os.path.isfile("{}.pkl".format(output_file_name)):
        os.remove("{}.pkl".format(output_file_name))

    with open("{}.pkl".format(output_file_name), "wb") as file:
        pickle.dump(map_regions, file)



if __name__ == '__main__':
    driver = start_driver(os.path.join(path, "chrome-data"))
    driver.maximize_window()

    # #precincts = [101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122,123,125,126,127,128,129,130,131,132,133,134,135,136,137,138,139,140,141,142,143,144,145,146,147,150,151,152,153,154,155,156,157,161,162,165,200,201,202,203,204,205,211,213,215,216,217,219,220,221,222,223,224,225,226,227,228,229,230,231,232,233,234,235,236,237,239,240,241,275,401]
    # precincts = [225,226,227,228,229,230,231,232,233,234,235,236,237,239,240,241,275,401]
    # #edit_steps = [new_step_1_mun_2020, new_step_2_mun_2020, new_step_3_mun_2020, new_step_4_mun_2020, new_step_5_mun_2020, new_step_6_mun_2020, new_step_7_mun_2020, new_step_8_mun_2020]
    # edit_steps = [new_step_1_mun_2020, new_step_2_mun_2020, new_step_3_mun_2020, new_step_4_mun_2020, new_step_5_mun_2020, new_step_6_mun_2020, new_step_7_mun_2020]
    # create_precincts(driver, precincts, edit_steps, 'New Sheet Precincts')

    get_VBM_turf_counts(driver, '**2021 Municipal St. Petersburg', 'vbm_turfs')
    print("done")

    