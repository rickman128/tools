import time
from selenium import webdriver

driver=webdriver.Ie("C:\\webdriver\\IEDriverServer.exe")

driver.get("https://www.google.com/")
time.sleep(3)
search_box = driver.find_element_by_name("q")
search_box.send_keys('今日の天気は？')
search_box.submit()
time.sleep(10)
driver.quit()