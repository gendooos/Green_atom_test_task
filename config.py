from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
chromedriver = r'chromedriver.exe'
service = Service(executable_path=chromedriver)
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, timeout=10000)