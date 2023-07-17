from selenium import webdriver
import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import csv
import time
from tkinter import Tk
import pyperclip as pc
from selenium.webdriver.common.by import By
import os
from PyQt5.Qt import QClipboard
import sys
from PyQt5.QtGui import QGuiApplication
from PyQt5.Qt import QApplication, QClipboard
from PyQt5 import QtWidgets

app = QtWidgets.QApplication(sys.argv)
clipboard = QGuiApplication.clipboard()
mimeData = clipboard.mimeData()

for format in mimeData.formats():
    print(f'{format}: {mimeData.data(format)}')

#messageFile = input(str("File of message = "))

messageFile = "m.txt"
messageFileObject = open(messageFile,"r")
messagesList = messageFileObject.readlines()
messagesss = "%0A".join(messagesList)
messageFileObject.close()
#imgName = input(str("Image file name? = "))

#imgPath = os.path.abspath('./photos/' + imgName)
#pc.copy (imgPath)
#filesss = input(str("File of contacts = "))
filesss = "HM.csv"

rows = []
with open(r"{}".format(filesss), encoding='UTF-8') as f:
    rows.extend(csv.reader(f,delimiter=",",lineterminator="\n"))
    '''next(rows, None)
    for row in rows:
        user = row[1]
        users.append(user)
        
        
    print(users)'''
class WhatsappBot:
    def __init__(self):

        self.mensagem = messagesss
        
        self.contatos = rows
        
        options = webdriver.ChromeOptions() 
        print("arguments needed")
        print(options.arguments)
        
        options.add_argument("--user-data-dir=C:/Users/pakaj/AppData/Local/Google/Chrome/User Data")
        options.add_argument("--profile-directory=Default")
        options.add_argument("--disable-extensions")
        # options.add_argument("--remote-debugging-port=8085")
        options.add_argument('--start-maximized')


        cap=DesiredCapabilities.CHROME.copy()

        print(options.arguments)

        service = webdriver.chrome.service.Service(executable_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe', capabilities = cap)
        #self.driver = webdriver.chrome.webdriver.WebDriver(service = service)
        #self.driver = uc.Chrome(browser_executable_path='C:\Program Files\Google\Chrome\Application\chrome.exe', options=options)
        self.driver = webdriver.Chrome(options=options)
        #self.driver = webdriver.chrome.webdriver.WebDriver()

    def SendMessages(self):
        print("here")
        linkWhatsAppCheck = 'https://web.whatsapp.com/'
        self.driver.get(linkWhatsAppCheck)
        Stop = input(str("Press Enter when WhatsApp loads completely"))

            
        for contato in self.contatos:
            try:
                link = 'https://web.whatsapp.com/send?phone=91'+ contato[1] + '&text=' + 'Dear ' + contato[0]+' ji ' + '😀'+ '%0A' + self.mensagem
                self.driver.get(link)
                time.sleep(8)
                
                #btn = self.driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button')
                if "jpeg" in pc.paste() or "jpg" in pc.paste() or "png" in pc.paste() : 
                    print("With photo" + pc.paste())
                    chatBox = self.driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]')
                    ActionChains(self.driver) \
                	.click(chatBox) \
					.key_down(Keys.CONTROL) \
					.send_keys('v') \
					.key_up(Keys.CONTROL) \
				    .perform()
                    btn = self.driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div')
                else :
                    print("No photo")
                    btn = self.driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button')
                btn.click()
                print("Success "+ contato[0] + '; ' + contato[1])
                time.sleep(5)
            except Exception as e: 
                print(e)
                print("Failed " + contato[0] + '; ' + contato[1])
                pass
            

bot = WhatsappBot()
bot.SendMessages()
bot.driver.quit()
