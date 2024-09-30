from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

a = 10*2+2-5/2-10+0.5/2-5-3.75
b = 123456

username = int(a)
password = b

chr_options = Options()
chr_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=chr_options)
driver.get("http://localhost/test/login.php")

user = driver.find_element(By.XPATH, '//*[@id="username"]')
pas = driver.find_element(By.XPATH, '//*[@id="password"]')
user.send_keys(username)
pas.send_keys(password)

login = driver.find_element(By.XPATH, '/html/body/div[1]/form/div[2]')
login.click()


