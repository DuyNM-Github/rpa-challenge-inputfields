# +
"""Imports, Variables, Declaration"""
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from RPA.Excel.Files import Files
from RPA.Browser.Selenium import Selenium

lib = Files()
driver = Selenium()
counter = 0

pathToExcel = "materials/rpa_challenge_mat.xlsx"
driver.open_available_browser("http://www.rpachallenge.com")  # navigate to site


# -

def readExcelFile():
    lib.open_workbook(pathToExcel)
    global data
    data = []
    try:
        print("Current active worksheet: " + lib.get_active_worksheet())
        sheetData = lib.read_worksheet("Sheet1", header=True)
        for row in sheetData:
            if row['First Name'] is not None:
                data.append(row)
    finally:
        lib.close_workbook()
        startBtn = driver.find_element("css:button")
        if startBtn.text == "START":
            startBtn.click()


def inputToBrowser():
    global counter
    inputFieldsContainer = driver.find_element("class:inputFields")
    inputFields = inputFieldsContainer.find_elements(By.CLASS_NAME, "col")
    for inputField in inputFields:
        inputLabel = inputField.find_element(By.TAG_NAME, "label").text
        if inputLabel == "First Name":
            inputField.find_element(By.TAG_NAME, "input").send_keys(data[counter]['First Name'])
        if inputLabel == "Last Name":
            inputField.find_element(By.TAG_NAME, "input").send_keys(data[counter]['Last Name '])
        if inputLabel == "Company Name":
            inputField.find_element(By.TAG_NAME, "input").send_keys(data[counter]['Company Name'])
        if inputLabel == "Role in Company":
            inputField.find_element(By.TAG_NAME, "input").send_keys(data[counter]['Role in Company'])
        if inputLabel == "Address":
            inputField.find_element(By.TAG_NAME, "input").send_keys(data[counter]['Address'])
        if inputLabel == "Email":
            inputField.find_element(By.TAG_NAME, "input").send_keys(data[counter]['Email'])
        if inputLabel == "Phone Number":
            inputField.find_element(By.TAG_NAME, "input").send_keys(data[counter]['Phone Number'])
    inputFieldsContainer.find_element(By.CLASS_NAME, "btn").click()
    counter += 1
    if(counter <= 9):
        inputToBrowser()


if __name__ == "__main__":
        readExcelFile()
        inputToBrowser()
