from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import time
import pandas as pd
import xlsxwriter 
import array as arr
df_data = pd.read_excel('address.xlsx', index_col=1)
name = pd.read_excel('address.xlsx', index_col=0)
# print(df_data.index)`  `
workbook = xlsxwriter.Workbook('postcode.xlsx')
worksheet = workbook.add_worksheet()
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(chrome_options)
driver.get("https://www.naver.com/")
time.sleep(5)
#first search postcode
driver.find_element("xpath",'/html/body/div[2]/div[1]/div/div[3]/div[2]/div/form/fieldset/div/input').send_keys(str(df_data.index[0]))
driver.find_element("xpath",'/html/body/div[2]/div[1]/div/div[3]/div[2]/div/form/fieldset/button').click()
# Find element by id
try:
    get = (driver.find_element(By.CLASS_NAME, 'knjJ8').get_attribute('outerHTML'))
    result = get[51:56]
    worksheet.write('A1', name.index[0])
    worksheet.write('B1', result)
except NoSuchElementException:  #spelling error making this code not work as expected
    worksheet.write('A1', name.index[0])
    worksheet.write('B1', "Could not search!")
    pass
for i in range(1, len(df_data.index)):
    print(name.index[i], end=" ")
    driver.find_element("xpath",'/html/body/div[3]/div[1]/div/div[1]/div[1]/div/form/fieldset/div[1]/input').send_keys(Keys.CONTROL, 'a')
    driver.find_element("xpath",'/html/body/div[3]/div[1]/div/div[1]/div[1]/div/form/fieldset/div[1]/input').send_keys(Keys.BACKSPACE)
    driver.find_element("xpath",'/html/body/div[3]/div[1]/div/div[1]/div[1]/div/form/fieldset/div[1]/input').send_keys(str(df_data.index[i]))
    driver.find_element("xpath",'/html/body/div[3]/div[1]/div/div[1]/div[1]/div/form/fieldset/button').click()
    # time.sleep(0.5)
    try:
        get = driver.find_element(By.CLASS_NAME, 'knjJ8').get_attribute('outerHTML')
        result = get[51:56]
        worksheet.write('A'+str(i+1), name.index[i])
        worksheet.write('B'+str(i+1), result)
        print(result)
    except NoSuchElementException:  #spelling error making this code not work as expected
        worksheet.write('A'+str(i+1), name.index[i])
        worksheet.write('B'+str(i+1), "Could not search!")
        pass
workbook.close()
driver.close()