# -*- coding: utf-8 -*-
"""
Created on Thu Jul 18 19:52:57 2019

@author: Tyler Wu
"""

#cd .spyder-py3\
#pyinstaller.exe --onefile --icon=classCheck.ico ClassCheck(1.4).py

#time
import time
from time import strftime, localtime
# emails
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import UnexpectedAlertPresentException, NoSuchWindowException, WebDriverException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# tkinter
import tkinter as tk
from PIL import Image, ImageTk
# to navigate to driver download page
import webbrowser
import os


# navigate to chrome driver DL page
def downloadDriver():
    dl_site = "https://chromedriver.storage.googleapis.com/index.html?path=76.0.3809.68/"
    webbrowser.open_new(dl_site)
    
# get info from the tkinter entry fields
def getInfoPack():
    infopack = [UNField.get(),PWField.get(),emailField.get(),pageTimeField.get(),driverTimeField.get()]
    return infopack

# get DR times if provided
def getDRTimes(DRTime):
    # handle only-hour inputs
    if(not(":" in DRTime)):
        if(int(DRTime) >= 10):
            DRTime1 = DRTime + ":00:00"
            DRTime2 = str(int(DRTime) + 12) + ":00:00"
        else:
            DRTime1 = "0" + DRTime + ":00:00"
            DRTime2 = str(int(DRTime) + 12) + ":00:00"
    # hour and minute input
    else:
        h,m = DRTime.split(':')
        if(int(h) >= 10):
            DRTime1 = DRTime + ":00"
            DRTime2 = str(12 + int(h)) + ":" + m + ":00"
        else:
            DRTime1 = "0" + DRTime + ":00"
            DRTime2 = str(12 + int(h)) + ":" + m + ":00"
    DRTimes = [DRTime1,DRTime2]
    return DRTimes

# main Process
def mainProcess(Event=None):
    # get values from tkinter
    infopack = getInfoPack()
    # close tkinter window
    root.destroy()
    while True:
        # open chrome
        option = webdriver.ChromeOptions()
        option.add_argument("--incognito")
        driver = webdriver.Chrome('C:\\chromedriver_win32\chromedriver.exe')
        loginPage = getToLogin(driver) # returns driver 
        status = performLogin(loginPage,infopack)
        if(status == 'GTL'):
            print("GTL True, relogging in")
            continue # restart from GTL
        else:
            classPlanner = status
        while True:
            checkClasses(classPlanner,infopack)
            print("returned to mp")
            break
    
# get to Login page
def getToLogin(driver):    
    print("gtl entered")
    # open driver and go to site
    driver.get('http://my.ucla.edu/')
    button = driver.find_element_by_id('ctl00_signInLink')
    button.click()
    return driver
    
# perform login
def performLogin(driver,infopack):
    print("pl entered")
    # explicit wait for login element to appear
    WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "logon"))
    )
    # fill in username and password
    loginField = driver.find_element_by_id('logon')
    passField = driver.find_element_by_id('pass')
    loginField.send_keys(infopack[0])
    passField.send_keys(infopack[1])
    # press enter
    passField.send_keys(Keys.ENTER)

    # explicit wait for user to DUO authenticate
    try:
        classButton = WebDriverWait(driver, 120).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="header-navbar"]/megamenu-megamenu/megamenu-supergroup[2]/li'))
        )
    except TimeoutException:
        print('Timed out waiting for Duo\nDriver will reopen at next Driver Refresh Time')
        DRTime = str(infopack[4])
        if(str(infopack[4]) == ''):
            print("No driver refresh times provided. Program exited.")
            driver.quit()
            os._exit(1)
        else:
            driver.quit()
            DRTimes = getDRTimes(DRTime)
            DRTime1 = DRTimes[0]
            DRTime2 = DRTimes[1]
            print("Driver refresh times: " + DRTime1 + ' & ' + DRTime2)
            while True:
                currentDRTime = strftime("%H:%M:%S",localtime())
                if(currentDRTime == DRTime1 or currentDRTime == DRTime2):
                    print('Restarting driver at: ' + currentDRTime)
                    return 'GTL'
            
    classButton.click()
    # press class planner button
    classPlanButton = driver.find_element_by_link_text('Class Planner')
    classPlanButton.click()
    return driver
    
# check current amount of classes and loop until decrease in 
def checkClasses(driver,infopack):
    print("cc entered")
    try:
        # set driver and page reset times
        DRTime = str(infopack[4])
        if(str(DRTime) == ''):
            print("No Driver refresh Times provided.")
            DRTime1 = DRTime2 = ''
        else:
            DRTimes = getDRTimes(DRTime)
            DRTime1 = DRTimes[0]
            DRTime2 = DRTimes[1]
            print("Driver refresh times: " + DRTime1 + ' & ' + DRTime2)
        PR = infopack[3]
        # set email
        email = infopack[2]
        # bool that indicates when to exit driver
        itsTime = False
        # count initial number of lock icons
        startCount = len(driver.find_elements_by_class_name('icon-lock'))
        # check for changes in lock icons while timeOut alert not present
        while True:
            if(itsTime == True):
                break
            currentCount = len(driver.find_elements_by_class_name('icon-lock'))
            # used to double check whether a class actually opened or just no longer on class planner page
            totalCount = len(driver.find_elements_by_class_name('icon-ok')) + len(driver.find_elements_by_class_name('icon-warning-sign'))
            #print(totalCount)
            # if change detected (class opened) and still on class planner page, send email 
            if(currentCount < startCount and totalCount != 0):
                print("start lock-icon count: " + str(startCount) + " current lock-icon count: " + str(currentCount))
                sendEmail(driver,email)
            # handle lost connection
            elif(totalCount == 0):
                print("Error attempting to find Class Planner Page at: " + strftime("%H:%M:%S",localtime()) + " attempting reconnect.")
                driver.quit()
                return
            # else no change, refresh
            else:
                startPRTime = time.time()
                while True:
                    currentDRTime = strftime("%H:%M:%S",localtime())
                    currentPRTime = time.time()
                    # restart driver when DR time reached
                    if(currentDRTime == DRTime1 or currentDRTime == DRTime2):
                        print("driver restarted: " + currentDRTime)
                        itsTime = True
                        break
                    # refresh page at page refresh intervals
                    elif(currentPRTime - startPRTime >= int(PR)):
                        driver.refresh()
                        refreshButton = driver.find_element_by_class_name('icon-refresh')
                        refreshButton.click()
                        break                      
        # itsTime=true -> return to mainProcess
        driver.quit()
        return
    # exception handling
    # if driver manually closed
    except NoSuchWindowException:
        print('driver manually closed. Ending script.')
        os._exit(1)
    except WebDriverException:
        print('driver manually closed. Ending script.')
        os._exit(1)
    # handle unexpected alerts
    except UnexpectedAlertPresentException:
        print('Unexpected alert present')
        alert_obj = driver.switch_to.alert
        print('Alert text: ' + alert_obj.text)
        alert_obj.accept()
        driver.quit()
        print("driver restarted")
        return
    # handle timeOut scenarios
    except TimeoutException:
        print('Timed out waiting for page to load')
        driver.quit()
        print("driver restarted")
        return
    
# send email to user
def sendEmail(driver,email):
    print('email fn called')
    sender_email = 'tylerwuds@gmail.com'
    receiver_email  = str(email)
    
    message = MIMEMultipart('alternative')
    message["Subject"] = "ClassCheckApp: A spot has opened!"
    message["From"] = sender_email
    message["To"] = receiver_email

    text = "A spot has opened, check Class planner!"
    html="""\
    <html>
        <head></head>
        <body>
            <p> A spot has opened, check <a href="https://be.my.ucla.edu/ClassPlanner/ClassPlan.aspx">Class planner</a>!</p>
        </body>
    </html>
    """
    
    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, '22onionking22')
        server.sendmail(sender_email, receiver_email, message.as_string())
    # once email sent successfully, terminate script
    driver.close()
    os._exit(1)

###### Tkinter application to accept user input
HEIGHT = 650
WIDTH = 900

root = tk.Tk()
root.title('UCLA Class Checker')
root.iconbitmap('classCheck.ico')
root.resizable(False, False)

canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
canvas.pack()

# background img
image = Image.open("ucla-.jpeg")
bg_img = ImageTk.PhotoImage(image)
bg_label = tk.Label(image=bg_img)
bg_label.place(relwidth=1,relheight=1)

# frames
frame0 = tk.Frame(root,bg='#ffffff',bd=2)
frame0.place(relx=.5,rely=.1, relwidth=.75,relheight=.06,anchor='n')

frame1 = tk.Frame(root,bg='#2774AE',bd=2)
frame1.place(relx=.5, rely =.175,relwidth=.75,relheight=.4,anchor='n')

frame2 = tk.Frame(root,bg='#ffffff',bd=2)
frame2.place(relx=.5, rely =.6,relwidth=.75,relheight=.3,anchor='n')

frame3 = tk.Frame(root,bg='#2774AE')
frame3.place(relx=.09,rely=.96, relwidth=.22,relheight=.04,anchor='n')

#labels
labelTitle = tk.Label(frame0,font=('Calibri',17),bg='#ffffff',anchor='nw',justify='left', text = 'UCLA Class Checker')
labelTitle.place(relwidth = 1,relheight =1)

labelUN = tk.Label(frame1,font=('Calibri',17),bg='#2774AE',fg='#ffffff',text = 'Username: ')
labelUN.place(relx=.05, rely=.05,relwidth=.16,relheight=.1)

labelPW = tk.Label(frame1,font=('Calibri',17),bg='#2774AE',fg='#ffffff',text = 'Password: ')
labelPW.place(relx=.05,rely=.2,relwidth=.15,relheight=.1)

labelEmail = tk.Label(frame1,font=('Calibri',17),bg='#2774AE',fg='#ffffff',justify='left',text = 'Email: ')
labelEmail.place(relx=.05,rely=.35,relwidth=.09,relheight=.1)

labelPageTime = tk.Label(frame1,font=('Calibri',17),bg='#2774AE',fg='#ffffff',justify='left',text = 'Page Refresh Time (sec): ')
labelPageTime.place(relx=.05,rely=.48,relwidth=.34,relheight=.15)

labelDriverTime = tk.Label(frame1,font=('Calibri',17),bg='#2774AE',fg='#ffffff',justify='left',text = 'Driver Refresh Time (hr): ')
labelDriverTime.place(relx=.05,rely=.65,relwidth=.34,relheight=.15)

labelCredit = tk.Label(frame3,font=('TkDefualtFont',8),justify='left',text='Program written by Tyler Wu')
labelCredit.place(relwidth=1,relheight=1)

# instructions
labelInfo = tk.Label(frame2,font=('Courier',10),bg='#ffffff', justify='left',text = '\
If the driver is not yet installed, please install and place the\n\
extracted installation folder in your C: Drive\n\n\
For the Page Refresh Time, please enter how often (in seconds)\n\
you would like the program to refresh the page\n\n\
For the Driver Refresh Time, please enter a time when you are able\n\
to Duo authenticate (e.g. entering 8 means 8:00 AM and PM, entering\n\
9:30 means 9:30 AM and PM) If none is provided, the driver will restart\n\
whenever myUCLA times out.')
labelInfo.place(relx=.05,rely=.1)

# entry fields
UNField = tk.Entry(frame1,font=('Courier',12))
UNField.place(relx=.25, rely=.05,relwidth=.5,relheight=.1)

PWField = tk.Entry(frame1,show="*",font=('Courier',12))
PWField.place(relx=.25, rely=.2,relwidth=.5,relheight=.1)

emailField = tk.Entry(frame1,font=('Courier',12))
emailField.place(relx=.25, rely=.35,relwidth=.5,relheight=.1)

pageTimeField = tk.Entry(frame1,font=('Courier',12))
pageTimeField.place(relx=.5, rely=.5,relwidth=.15,relheight=.1)

driverTimeField = tk.Entry(frame1,font=('Courier',12))
driverTimeField.place(relx=.5, rely=.65,relwidth=.15,relheight=.1)

# buttons
# main process button
enterButton = tk.Button(frame1, text='Enter',command=lambda: mainProcess())
enterButton.place(relx=.8,rely=.7,relwidth=.1,relheight=.1)
root.bind('<Return>',mainProcess)

# driver download button
driverButton = tk.Button(frame1, text='Download Driver',command=lambda: downloadDriver())
driverButton.place(relx=.5,rely=.95,relwidth=.3,relheight=.1,anchor='s')

root.mainloop()