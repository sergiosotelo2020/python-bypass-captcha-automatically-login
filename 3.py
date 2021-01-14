from selenium import webdriver 
  
# Here Chrome  will be used 
driver = webdriver.Chrome() 
  
# URL of website 
url = "https://www.geeksforgeeks.org/"
  
# Opening the website 
driver.get(url) 
  
# Closes the current window 
driver.close()