"""last message - without group condition"""

from tkinter import *
from tkinter import filedialog
import tkinter as tk
from selenium import webdriver
from selenium.common.exceptions import (NoSuchElementException,
                                        StaleElementReferenceException,ElementClickInterceptedException,TimeoutException)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from time import sleep
import random
import pandas
from tkinter import simpledialog
#pip install webdriver-manager


options = Options()
options.add_argument("profile-directory=Profile 1")
options.add_argument("--user-data-dir=C:/Users/THAAGAM BOT PC-1/Desktop/whatsapp Bulk/WhatsappBulk/chrome-data/test1")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

driver = webdriver.Chrome(executable_path="chromedriver.exe", options=options)
driver.get("http://web.whatsapp.com")
sleep(10)  # wait time to scan the code in second


element = WebDriverWait(driver, 10).until(
    lambda x: driver.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[1]/div[3]/div/div[1]/div/button/div[2]/span'))



def element_presence(by, xpath, time):
  element_present = EC.presence_of_element_located((By.XPATH, xpath))
  WebDriverWait(driver, time).until(element_present)


def send_whatsapp_msg(phone_no, text):
  driver.get("https://web.whatsapp.com/send?phone={}".format(phone_no))
  try:
    driver.switch_to.alert().accept()
  except Exception as e:
    pass


  try:
    element_presence(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]', 40)
    txt_box = driver.find_element(By.XPATH,"/html/body/div[1]/div[1]/div[1]/div[4]/div[1]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]")
    txt_box.click()
    txt_box.send_keys(text)
    txt_box.send_keys("\n")

  except Exception as e:
    print(e)

  sleep(random.randint(10,40))

excel = filedialog.askopenfilename(initialdir="/", title="Select A File", filetypes=(("Excel File", "*.xlsx"),
                                                                                       ("all files", "*.*")))
application_window = tk.Tk()
#excel = r"C:\Users\Raja Monsingh\PycharmProjects\WhatsappBulk\Book1.xlsx"
text = answer = simpledialog.askstring("Input", "Input Whatsapp Message",
                                parent=application_window)

df = pandas.read_excel(excel)

for i in range(len(df)):
  name = (df.loc[i, "NAME"])
  no = (df.loc[i, "NUM"])


  textcus = text.replace("{DONOR}",name)

  no = str(no)
  if len(no) == 10:
      no = ("91" + no)



  send_whatsapp_msg(no,textcus)

sleep(60)
exit()
