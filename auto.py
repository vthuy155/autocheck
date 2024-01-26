from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
import xlsxwriter 
import array as arr
df_data = pd.read_excel('test.xlsx', index_col=1)
name = pd.read_excel('test.xlsx', index_col=0)
# print(df_data.index)`  `
workbook = xlsxwriter.Workbook('result.xlsx')
worksheet = workbook.add_worksheet()
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(chrome_options)
driver.get("https://unipass.customs.go.kr/csp/index.do")
driver.find_element("xpath",'/html/body/div[2]/section[2]/section[1]/article[2]/div[2]/div[1]/fieldset/ul/li[2]/input').click()
driver.find_element("xpath",'/html/body/div[2]/section[2]/section[1]/article[2]/div[2]/div[1]/fieldset/form/span[2]/input[2]').send_keys(str(df_data.index[0]))
# driver.find_element("xpath",'/html/body/div[2]/section[2]/section[1]/article[2]/div[2]/div[1]/fieldset/form/span[2]/select').click()
# driver.find_element("xpath",'/html/body/div[2]/section[2]/section[1]/article[2]/div[2]/div[1]/fieldset/form/span[2]/select/option[3]').click()
driver.find_element("xpath",'/html/body/div[2]/section[2]/section[1]/article[2]/div[2]/div[1]/fieldset/form/input').click()
time.sleep(10)
# Find element by id
get = driver.find_element(By.ID, 'MYC0405102Q_impImg8').get_attribute('outerHTML')
result = "on" in get
print(name.index[0], end=" ")
print(str(df_data.index[0]), end=" ")
print("on" in get)
# a = ["a", "b" "c"]
worksheet.write('A1', name.index[0])
worksheet.write('B1', str(df_data.index[0]))
worksheet.write('C1', result)
for i in range(1, len(df_data.index)):
    print(name.index[i], end=" ")
    print(str(df_data.index[i]), end=" ")
    # time.sleep(5)
    driver.find_element("xpath",'/html/body/div[2]/section[1]/div[3]/div[1]/form[2]/div/div/table/tbody/tr[1]/td[2]/input[2]').send_keys(str(df_data.index[i]))
    driver.find_element("xpath",'/html/body/div[2]/section[1]/div[3]/div[1]/form[2]/div/div/table/tbody/tr[1]/td[1]/label/input').click()
    # driver.find_element("xpath",'/html/body/div[2]/section[1]/div[3]/div[1]/form[2]/div/div/table/tbody/tr[1]/td[2]/select').click()
    # driver.find_element("xpath",'/html/body/div[2]/section[1]/div[3]/div[1]/form[2]/div/div/table/tbody/tr[1]/td[2]/select/option[3]').click()
    driver.find_element("xpath",'/html/body/div[2]/section[1]/div[3]/div[1]/form[2]/div/footer/button').click()
    time.sleep(7)
    get = driver.find_element(By.ID, 'MYC0405102Q_impImg8').get_attribute('outerHTML')
    result = "on" in get
    worksheet.write('A'+str(i+1), name.index[i])
    worksheet.write('B'+str(i+1), str(df_data.index[i]))
    worksheet.write('C'+str(i+1), result)
    print(result)
workbook.close()
driver.close()