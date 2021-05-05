from selenium import webdriver
import pandas as pd
import openpyxl
import time

#Extracting the excel data
file = "worksheet/challenge.xlsx"

df = pd.read_excel(file, sheet_name="Sheet1")

"""
    It was necessary to convert the phones column from excel sheet to a list,
    to avoid this error: TypeError: object of type 'numpy.int64' has no len()
"""
phones = df["Phone Number"].to_list()

#Scraping the RPA Challange Website

rpa_challenge_website = "http://www.rpachallenge.com/"

"""
    The code below remove the message: 
    Browser controlled by test tool
"""
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("useAutomationExtension", False)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])

driver = webdriver.Chrome(executable_path="chrome_driver/chromedriver.exe", options=chrome_options)
driver.get(url="http://www.rpachallenge.com/")

driver.maximize_window()

start_button = driver.find_element_by_xpath("/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button")
start_button.click()

def scrap_rpa_challenge():
    """Function to scrap the RPA Challange"""
    for i in range(len(df)):
        first_name = driver.find_element_by_xpath("//input[@ng-reflect-name='labelFirstName']")
        first_name.send_keys(df["First Name"][i])
        last_name = driver.find_element_by_xpath("//input[@ng-reflect-name='labelLastName']")
        last_name.send_keys(df["Last Name "][i])
        company_name = driver.find_element_by_xpath("//input[@ng-reflect-name='labelCompanyName']")
        company_name.send_keys(df["Company Name"][i])
        role = driver.find_element_by_xpath("//input[@ng-reflect-name='labelRole']")
        role.send_keys(df["Role in Company"][i])
        address = driver.find_element_by_xpath("//input[@ng-reflect-name='labelAddress']")
        address.send_keys(df["Address"][i])
        email = driver.find_element_by_xpath("//input[@ng-reflect-name='labelEmail']")
        email.send_keys(df["Email"][i])
        phone = driver.find_element_by_xpath("//input[@ng-reflect-name='labelPhone']")
        phone.send_keys(phones[i])
        submit_button = driver.find_element_by_xpath("/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input")
        submit_button.click()


scrap_rpa_challenge()

result = driver.find_element_by_xpath("/html/body/app-root/div[2]/app-rpa1/div/div[2]/div[2]")
print(result.text)
time.sleep(2)
start = result.text.find("in")
end = result.text.find("milliseconds")
result_milliseconds = result.text[start:end].split(" ")[1]
print(result_milliseconds)

driver.quit()

with open("results.txt", mode="a") as file:
    file.write(f"Result: {result_milliseconds}\n")

"""
XPaths from challenge RPA web site
<input _ngcontent-c2="" ng-reflect-name="labelFirstName" id="WcPhQ" name="WcPhQ" class="ng-pristine ng-invalid ng-touched">
<input _ngcontent-c2="" ng-reflect-name="labelLastName" id="u3kIC" name="u3kIC" class="ng-untouched ng-pristine ng-invalid">
<input _ngcontent-c2="" ng-reflect-name="labelCompanyName" id="5X9Zx" name="5X9Zx" class="ng-untouched ng-pristine ng-invalid">
<input _ngcontent-c2="" ng-reflect-name="labelRole" id="rXFF5" name="rXFF5" class="ng-untouched ng-pristine ng-invalid">
<input _ngcontent-c2="" ng-reflect-name="labelAddress" id="gd7SB" name="gd7SB" class="ng-untouched ng-pristine ng-invalid">
<input _ngcontent-c2="" ng-reflect-name="labelEmail" id="IrXJk" name="IrXJk" class="ng-untouched ng-pristine ng-invalid">
<input _ngcontent-c2="" ng-reflect-name="labelPhone" id="p4Z55" name="p4Z55" class="ng-untouched ng-pristine ng-invalid">

Relative XPaths used to find input locations on page changes
//input[@ng-reflect-name='labelFirstName']
//input[@ng-reflect-name='labelLastName']
//input[@ng-reflect-name='labelCompanyName']
//input[@ng-reflect-name='labelRole']
//input[@ng-reflect-name='labelAddress']
//input[@ng-reflect-name='labelEmail']
//input[@ng-reflect-name='labelPhone']
"""




