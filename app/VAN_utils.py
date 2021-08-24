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


def preview_results(driver):
    preview_results_button = expect_by_XPATH(driver, '//*[@id="ResultsPreviewButton"]/span[2]')
    preview_results_button.click()

def add_step(driver, add):
    sleep_random_time(0.75)
    add_step_button = expect_by_XPATH(driver, '//*[@id="AddNewStepButton"]/span[1]')
    add_step_button.click()

    sleep_random_time(0.25)
    
    if add:
        add_people_button = expect_by_XPATH(driver, '//*[@id="SearchSummaryButtons"]/div[2]/div/ul/li[1]/span')
        add_people_button.click()
    else:
        remove_people_button = expect_by_XPATH(driver, '//*[@id="SearchSummaryButtons"]/div[2]/div/ul/li[2]/span')
        remove_people_button.click()

    sleep_random_time(1)

def narrow(driver):
    sleep_random_time(0.25)
    add_step_button = expect_by_XPATH(driver, '//*[@id="AddNewStepButton"]/span[1]')
    add_step_button.click()

    sleep_random_time(0.25)
    
    narrow_people_button = expect_by_XPATH(driver, '//*[@id="SearchSummaryButtons"]/div[2]/div/ul/li[3]/span')
    narrow_people_button.click()

    sleep_random_time(0.25)

def toggle(element, on):
    if element.is_selected() == on:
        return
    else:
        element.click()
    

def run_search(driver):
    sleep_random_time(1)
    run_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_SearchRunButton"]')
    run_button.click()
    sleep_random_time(1.5)

def sleep_random_time(offset):
    sleep_time = random.randint(4, 5) + offset
    time.sleep(sleep_time)

def default_voter_status(driver):
    voter_status_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_UpdatePanelVoterStatus"]/div[1]/div/label')
    voter_status_button.click()

    reg_active_button = expect_by_XPATH(driver, '//*[@id="PanelSectionVoterStatus"]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input[1]')
    toggle(reg_active_button, True)

    reg_inactive_button = expect_by_XPATH(driver, '//*[@id="PanelSectionVoterStatus"]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input[2]')
    toggle(reg_inactive_button, True)

    likely_button = expect_by_XPATH(driver, '//*[@id="PanelSectionVoterStatus"]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input[3]')
    toggle(likely_button, False)
    
    other_button = expect_by_XPATH(driver, '//*[@id="PanelSectionVoterStatus"]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input[4]')
    toggle(other_button, True)


def save_list(driver, list_name, target_folder_name):
    save_list_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_saveAsButton"]')
    save_list_button.click()

    sleep_random_time(2)

    saved_search_button = expect_by_XPATH(driver, '//*[@id="SaveSearchRadioBtn"]')
    saved_search_button.click()
    
    folder_dropdown = expect_by_XPATH(driver, '//*[@id="Folder"]')
    folder_dropdown_select = Select(folder_dropdown)
    folder_dropdown_select.select_by_visible_text(target_folder_name)

    name_input = expect_by_XPATH(driver, '//*[@id="Name"]')
    name_input.clear()
    name_input.send_keys(list_name)

    save_button = expect_by_XPATH(driver, '//*[@id="ctl00_ContentPlaceHolderVANPage_SubmitButton"]')
    save_button.click()


def edit_list(driver):
    edit_button = expect_by_XPATH(driver, '//*[@id="addStep"]')
    edit_button.click()
