from main import *


def startdo():
    """This function is main executor. First it logins in gsheets, launches webDriver, logins in Sharepoint and
    fill form. Many data is needed to be manualy inserted - can see it by CAPS value. After collecting data, forms fill
    and code checks if it was registred and leaves status on google sheet. this template is base for any form we
    want to fill, so one form is one module based on this module.
    """
    sheet = gmail_login()
    driver = start_driver()

    #################### BLOCK WHERE WE INSERT INFORMATION MANUALY FOR EVERY FORM
    link_register_path = 'INSERT LINK FOR REGISTRATION'
    submit_buttonx = 'INSERT XPATH TO SUBMIT BUTTON'
    rownum = list(range(1, 683))  # Range of rows in google sheet we need to work with
    checkdonecol = "NUMBER OF COLUMN FOR STATUS"
    checkxpath = 'XPATH TO ELEMENT THAT APPEARS AFTER FORM APPLIED'
    ####################

    share_login(link_register_path, driver)
    for row in rownum:
        #################### BLOCK WHERE TO INSERT INFORMATION MANUALY (column of name, subject, etc. and xpath)
        variable = sheet.cell(row, 2).value  # Instead of '2' set column number of variable value.
        variable_xpath = 'INSERT XPATH'
        ####################
        try:
            driver.find_element_by_xpath(variable_xpath + "/option[text()='%s']" % variable).click()
        except selenium.common.exceptions.NoSuchElementException:
            sheet.update_cell(row, checkdonecol, "variable-error")
            driver.get(link_register_path)
            continue

        driver.find_element_by_xpath(submit_buttonx).click()
        driver.implicitly_wait(10)
        error_check(row, sheet, driver, link_register_path, checkdonecol, checkxpath)


if __name__ == '__main__':
    startdo()
