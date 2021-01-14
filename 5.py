from selenium import webdriver
from PIL import Image
import selenium.webdriver
# take screenshot

driver = selenium.webdriver.Chrome()
driver.get('https://sso.gem.gov.in/ARXSSO/oauth/login')
element = driver.find_element_by_id("captcha1")
location = element.location
size = element.size
driver.save_screenshot("pageImage.png")

# crop image
x = location['x']
y = location['y']
width = location['x']+size['width']
height = location['y']+size['height']
im = Image.open('pageImage.png')
im = im.crop((int(x), int(y), int(width), int(height)))
im.save('element.png')

driver.quit()