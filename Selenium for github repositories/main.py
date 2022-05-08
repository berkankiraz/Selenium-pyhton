
#ismail berkan kiraz.
#You write your username and password for github in githubUserInfo.

from githubUserInfo import username, password
from selenium.webdriver.common.by import By
from selenium import webdriver
import time
import pandas as pd

class Github:
    def __init__(self , username , password):
        self.browser = webdriver.Chrome()
        self.username = username
        self.password = password
        self.repository=[]

    #singın method is used for login process.
    def signIn(self):
        url = "https://github.com/login"
        self.browser.get(url)
        time.sleep(2)
        username = self.browser.find_element(By.NAME, "login") #github css
        password = self.browser.find_element(By.NAME, "password")#github css
        username.send_keys(self.username)
        password.send_keys(self.password)
        time.sleep(1)
        self.browser.find_element(By.NAME, "commit").click()
        time.sleep(1)
    #resporitygetin is used for taking your repositry's name.
    def resporitygetin(self):
        self.browser.find_element(By.XPATH, "/html/body/div[1]/header/div[7]/details/summary").click()#github css
        time.sleep(1)
        self.browser.find_element(By.XPATH,"/html/body/div[1]/header/div[7]/details/details-menu/a[2]").click()#github css
        items = self.browser.find_elements(By.CSS_SELECTOR,".col-10.col-lg-9.d-inline-block")#github css
        time.sleep(1)
        for i in items:
            self.repository.append(i.find_element(By.CSS_SELECTOR,".d-inline-block.mb-1 h3 a").text)#github css
        print(self.repository)
        time.sleep(1)
        self.browser.close()
    #savereponames is used for writing it into the excel.
    def savereponames(self):
        df = pd.DataFrame(self.repository)
        writer = pd.ExcelWriter('myreposname.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='welcome', index=False)
        writer.save()
#object has been created here.
github = Github(username,password)
github.signIn()
github.resporitygetin()
github.savereponames()




