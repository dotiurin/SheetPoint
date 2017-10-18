import json
import gspread
from oauth2client.client import SignedJwtAssertionCredentials
from selenium import webdriver
import time
import selenium
from getpass import getpass


def start_driver():
    """Common part of any module, so put it as function. Starts webDriver and waits it to load."""
    driver = webdriver.Chrome()
    driver.set_page_load_timeout(20)
    return driver


def share_login(link_register_path, driver):
    """Gets data for login from input, then uses them with selenium to login with webDriver"""
    login = input("Insert sharepoint login:")
    password = getpass("Insert sharepoint password:")
    driver.get(link_register_path)
    driver.find_element_by_xpath('//*[@id="cred_userid_inputtext"]').send_keys(login)
    driver.find_element_by_xpath('//*[@id="cred_password_inputtext"]').send_keys(password)
    time.sleep(3)
    driver.find_element_by_xpath('//*[@id="cred_sign_in_button"]').click()


def gmail_login():
    """Logins into gmail. "creds.json" should be in work directory, gmail account and google sheet configured.
    In method "get_worksheet()" you can insert the list u are going to work with in your google sheet as an argument

    """
    gmail_sheet_link = input("Insert link to gmail page:")
    json_key = json.load(open('creds.json'))
    scope = ['https://spreadsheets.google.com/feeds']
    credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'].encode(), scope)
    file = gspread.authorize(credentials)
    sheet = file.open_by_url(gmail_sheet_link).get_worksheet(0)  # List number as argument (default list is 0)
    return sheet


def error_check(row, sheet, driver, link_register_path, checkdonecol, checkxpath):
    """Function tries to find element by xpath after form confirmation. If element appears in webDriver function
    inserts NICE in row we passed in gsheet, if element doesnt appear function waits for two seconds and tries to find
    it again, these try are counted in 'exception_count' and if element is not found it skips that row in gspread to
    register and leave mark ERROR in failed row.

    """
    exception_count = 0
    while True:
        if exception_count <= 5:
            try:
                driver.find_element_by_xpath(checkxpath)
                sheet.update_cell(row, checkdonecol, "NICE")
                print("account done, row: " + str(row))
                driver.get(link_register_path)  # Comment this if you commit form on single page without refreshing.
                break
            except selenium.common.exceptions.ElementNotSelectableException:
                print("count " + str(exception_count))
                time.sleep(2)
                exception_count += 1
        else:
            sheet.update_cell(row, checkdonecol, "ERROR")
            print("ERROR, row: " + str(row))
            driver.find_element_by_xpath(link_register_path).click()
            break
