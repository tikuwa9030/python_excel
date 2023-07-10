import time

from selenium.webdriver import Chrome

try:
    driver = Chrome()
    driver.get('https://www.google.co.jp/')

    element = driver.find_element_by_name('q')
    element.send_keys('Wikipedia')
    element.submit()

    time.sleep(10)

    print(driver.find_element_by_id('result-stats').text)

finally:
    driver.quit()
