
import os
import time
import configparser
import ast
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#GENERAL CONFIGURATION
#READ CONFIG FILE
cfg = configparser.ConfigParser()
ini_config_path = os.path.join(os.getcwd(),'powerbi_workspace_auto_refresh.ini')
cfg.read(ini_config_path)
#READ INFO FROM CONFIG FILE
driver_path_at_onedrive = cfg['general_config']['driver_path_at_onedrive']
type_object_name = cfg['general_config']['type_object_name']
button_refresh_name = cfg['general_config']['button_refresh_name']
dict_list = cfg['general_config']['dict_list']
dict_list = ast.literal_eval(dict_list)
one_drive_path = os.environ['ONEDRIVE']#GET ONE DRIVE FOLDER
driver_path = os.path.join(one_drive_path, driver_path_at_onedrive)

#FUNCTIONS
#FUNCTION TO GET ELEMENT LIST FROM RECEIVED URL
def get_element_list_from_url(url_received):
    driver.get(url_received)
    driver.implicitly_wait(10)
    time.sleep(2)
    element_list_to_iterate = WebDriverWait(driver, 120).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="artifactContentList"]/div[1]/div')))
    return element_list_to_iterate

#FUNCTION TO CLICK ON BUTTON
def click_button(element_list, data_set_name):
    action = ActionChains(driver)
    for element in element_list:
        element_name = element.find_elements_by_xpath('./div')[1].text
        element_type = element.find_elements_by_xpath('./div')[2].text
        if element_name == data_set_name:
            if element_type == type_object_name:
                action.move_to_element(element).perform()
                time.sleep(1)
                buttons = element.find_elements_by_xpath('./div[2]/span/button')
                for button in buttons:
                    button_title = None
                    try:
                        button_title = button.get_attribute('title')
                    except:
                        button_title = None
                    if button_title == button_refresh_name:
                        try:
                            button.click()
                        except:
                            pass
                        time.sleep(1)

#FUNCTION TO REFRESH DATA_SET LIST FROM URL
def refresh_data_set_list_from_url(url_received,data_set_name_list):
    element_list_to_iterate = get_element_list_from_url(url_received)
    for data_set in data_set_name_list:
        try:
            click_button(element_list_to_iterate, data_set)
        except:
            pass
        time.sleep(1)

#MAIN FUNCTION START HERE
try:
    driver = webdriver.Edge(driver_path)
    action = ActionChains(driver)

    for l_dict in dict_list:
        url_received = l_dict['url']
        data_set_name_list = l_dict['data_set_list']
        refresh_data_set_list_from_url(url_received,data_set_name_list)
except:
    pass

try:
    driver.quit()
except:
    pass