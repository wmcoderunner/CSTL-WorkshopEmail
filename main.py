import pymssql
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import time as tm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
'''
Ask Nick or Mian for the information of line 72, 79, and 131-133,177-178, 226-228, 273-274. Replace * with your information
1. If you are the first time to run it:
    i. Pycharm
        Download Pycharm
        https://www.jetbrains.com/help/pycharm/configuring-python-interpreter.html#add_new_project_interpreter
        The link is the way to set the interpreter up.
        Make sure your Pycharm can run helloworld program.
        Copy and paste this code to the py file, or copy and paste this file to folder
        If there are any errors in the import line, please follow this:
        Click "File" in the up left, then click "Settings", click "Project:pythonProject1" in the left of setting page, 
        click "PythonInterpreter", click "+" sign in the left bottom in the setting page, 
        search "smtplib", and install this package
        search "selenium", and install this package
        search "email", and install this package
        search "pymssql", and install this package
    ii. Chromedriver
        Make sure download correct chromedriver and put chromedriver in your desktop
        chrome://settings/help Put this link to your chrome can help you find the Chrome version
        https://chromedriver.chromium.org/downloads This link is for you to download the chrome driver
    iii. Path
        Change line 49 to your path. 
2. General Information:
    Run the program and sleep 1 minute, then go to your email to check the result, then cross your finger!
3. Before you run it:
    MAKE SURE THAT THE ATTENDEDENCE OF THE WORKSHOP HAVE MADED
    You may need to test. If you want to test, see the following:
    i. before test:
        Change the testemailaddress (line 50) to your test email, comment line 150, 188, 247, 278!!!
    ii. after test:
        IMPORTANT:After you test successfully, uncomment line 150, 188, 247, 278!!!
        Do not add the line or delete the line in this code, otherwise, the line of instruction will change!
Panacea: 
    Use the computer in the front desk and log in to Mian's account, open the pycharm, 
    open the workshop(It should be default), do not modify anything, run it and sleep one minute. 
Annotation:
    Sleep one minute means do not touch mouse and keyboard until it finished.
If you have any questions, please feel free to email me at mwu2s@semo.edu
'''
path = "C:\\Users\\sw03cstl\\"  # Change it to your path
testemailaddress = ['mwu2s@semo.edu','sw03ctl@semo.edu'] #Add your email as the test email
fname = []
lname = []
email = []
title = []
series = []
workshopid = []
presenter = []
location = []
time = []
dict = {}
date = []
description=[]


def createForm(title, content):
    browser = webdriver.Chrome(path + "Desktop\\chromedriver.exe")
    browser.maximize_window()
    wait = WebDriverWait(browser, 10)
    browser.get('https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&rver=7.3.6963.0&wp=MBI_SSL&wreply=https%3a%2f%2fwww.microsoft.com%2fmicrosoft-365%2fonline-surveys-polls-quizzes&lc=1033&id=74335&aadredir=1')
    input = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, '#i0116')))
    input.send_keys('*')
    submit = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, '#idSIButton9')))
    submit.click()
    tm.sleep(3)
    input = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, '#i0118')))
    input.send_keys('*')
    submit = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, '#idSIButton9')))
    submit.click()
    tm.sleep(3)

    submit = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, '#idBtn_Back')))
    submit.click()
    tm.sleep(3)

    browser.get(
        'https://forms.office.com/Pages/ShareFormPage.aspx?id=mHXVGbCXLES3TshVsNh8r7UxtjJegA5MkAA0BZwnkg9UMlVCVVE2Q04wVFg5MDdOUUZIRjNOUTgzNC4u&sharetoken=SSRzJvw8RAc3mXS57kfx')
    browser.maximize_window()
    tm.sleep(3)

    submit = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, '#form-container > div > div > div.office-form-duplicate-form-container > div > div > button')))
    submit.click()
    tm.sleep(10)

    browser.find_element_by_xpath('// *[ @ id = "form-designer"] / div[1] / div / div').click()

    input = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, '#form-designer > div:nth-child(1) > div > div.__title-designer__.form-designer-form-title.office-form-theme-focus-border.title-designer-container > div.title-designer-title-box.right > div > div > textarea')))
    input.send_keys(title)

    submit = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, '#form-designer > div:nth-child(1) > div > div.__title-designer__.form-designer-form-title.office-form-theme-focus-border.title-designer-container > div.title-designer-subtitle-editor > div > textarea')))
    submit.click()
    tm.sleep(1)

    input = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR,
         '#form-designer > div:nth-child(1) > div > div.__title-designer__.form-designer-form-title.office-form-theme-focus-border.title-designer-container > div.title-designer-subtitle-editor > div > textarea')))
    input.send_keys(content)

    submit = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR,
         '#content-root > div > div:nth-child(1) > div > div > div:nth-child(4) > button')))
    submit.click()
    tm.sleep(3)

    formlink = browser.find_element_by_xpath('//*[@id="flex-pane-textbox-link"]').get_attribute('value')
    print(formlink)
    return formlink



#Send the certificate for the workshop which hold after 24 hours before.
def certificate1():

    conn = pymssql.connect(host='*', user='*',
                               password='*', database='*',
                               charset="*")

    cursor = conn.cursor()
    sql = "SELECT FirstName,LastName,Email,Title,SeriesTitle,dbo.Workshops.WorkshopID,Presenter,Location,WorkshopStart,has_certificate,Attended"\
          " FROM dbo.Workshops INNER JOIN dbo.Enrollment ON dbo.Workshops.WorkshopID = dbo.Enrollment.WorkshopID" \
              " INNER JOIN dbo.Participants ON dbo.Participants.ParticipantID = dbo.Enrollment.ParticipantID"\
              " INNER JOIN dbo.WorkshopSeries ON dbo.WorkshopSeries.SeriesID = dbo.Workshops.WorkshopSeriesID"\ 
              " WHERE  WorkshopStart BETWEEN GETDATE()-7 AND GETDATE() AND Attended = 1"

    cursor.execute(sql)



    rs = cursor.fetchall()
    sql1 = "UPDATE dbo.Workshops SET has_certificate = 1 " \
          "WHERE has_certificate = 0 AND WorkshopStart BETWEEN GETDATE()-7 AND GETDATE()"

    cursor.execute(sql1)
    for item in rs:
        email.append(item[2])
        fname.append(item[0])
        lname.append(item[1])
        title.append(item[3])
        series.append(item[4])
        workshopid.append(item[5])
        presenter.append(item[6])
        location.append(item[7])
        d=str(item[8]).split(' ')[0]
        t=str(item[8]).split(' ')[1].split('.')[0]
        date.append(d)
        time.append(t)

    for i in range(len(workshopid)):
        if workshopid[i] not in dict:
            a = [title[i], series[i], presenter[i],location[i],date[i], time[i],[email[i]]]
            dict[workshopid[i]] = a
        elif workshopid[i] in dict:
            a = dict[workshopid[i]]
            a[6].append(email[i])
            dict[workshopid[i]] = a



    mail_host = 'smtp.office365.com'
    mail_user = '*'
    mail_pass = '*'
    for id in dict:
        print(id,dict[id])

        link = createForm(dict[id][1]+": "+dict[id][0],
                          "Title: "+dict[id][0]+"\nDate: "+dict[id][4]+" "+dict[id][5]+"\nLocation: Zoom\n"
                         "Presenter(s): "+dict[id][2]+ "\nPlease rate the session you attended by checking the appropriate option in the form below.")

        bcc = testemailaddress
        # Check!!!
        bcc = dict[id][6]
        receivers = bcc

        string = "Dear faculty, \n\nThank you for attending "+dict[id][1]+"\nTopic: "+dict[id][0]+"\nDATE: "+dict[id][4]+" "+dict[id][5]+\
                 "\nPresenter: "+dict[id][2]+\
                 "\nLocation: Zoom"+\
                "\n\nYour certificate of attendance is now ready for download. Please visit https://www.semo.edu/ctl/faculty-development/events.html " \
                "click “Attended Events” to download a certificate for any workshops you have attended."\
                "\n\nWe would appreciate you also filling out a brief survey regarding your experience. Your feedback helps us to continuously improve and provide the high standards you expect from CTL. " \
                "Please visit "+link+\
                "\n\nIf you have any questions regarding the survey, please email ctlsupport@semo.edu. Thank you in advance for your feedback. We look forward to seeing you again at another workshop soon."\
                "\n\nSincerely,\nCTL Staff"

        message = MIMEText(string)

        subject = 'Post ' + dict[id][1]+" Notifications"
        message['Subject'] = Header(subject, 'utf-8')

        try:
            smtpObj = smtplib.SMTP(mail_host, 587)
            smtpObj.connect(mail_host, 587)
            smtpObj.ehlo()
            smtpObj.starttls()
            smtpObj.ehlo()
            smtpObj.login(mail_user, mail_pass)
            smtpObj.sendmail(mail_user, receivers, message.as_string())
            print("Email Sent Successfully")
        except smtplib.SMTPException as e:
            print("Error: ", e)







#Send the reminder for the workshop
def reminder():
    conn = pymssql.connect(host='*', user='*',
                               password='*', database='*',
                               charset="*")

    cursor = conn.cursor()



    sql = "SELECT FirstName,LastName,Email,Title,SeriesTitle,dbo.Workshops.WorkshopID,Presenter,Location,WorkshopStart,dbo.Workshops.Description" \
          " FROM dbo.Workshops INNER JOIN dbo.Enrollment ON dbo.Workshops.WorkshopID = dbo.Enrollment.WorkshopID " \
              "INNER JOIN dbo.Participants ON dbo.Participants.ParticipantID = dbo.Enrollment.ParticipantID " \
              "INNER JOIN dbo.WorkshopSeries ON dbo.WorkshopSeries.SeriesID = dbo.Workshops.WorkshopSeriesID " \
              "WHERE has_reminder = 0 AND WorkshopStart BETWEEN GETDATE() AND GETDATE()+1.5"

    cursor.execute(sql)



    rs = cursor.fetchall()
    sql1 = "UPDATE dbo.Workshops SET has_reminder = 1 " \
          "WHERE has_reminder = 0 AND WorkshopStart BETWEEN GETDATE() AND GETDATE()+1.5"
    cursor.execute(sql1)
    for item in rs:
        email.append(item[2])
        fname.append(item[0])
        lname.append(item[1])
        title.append(item[3])
        series.append(item[4])
        workshopid.append(item[5])
        presenter.append(item[6])
        location.append(item[7])
        description.append(item[9])
        d=str(item[8]).split(' ')[0]
        t=str(item[8]).split(' ')[1].split('.')[0]
        date.append(d)
        time.append(t)

    for i in range(len(email)):
        if workshopid[i] not in dict:
            a = [title[i], series[i], presenter[i],location[i],date[i], time[i],[email[i]],description[i]]
            dict[workshopid[i]] = a
        elif workshopid[i] in dict:
            a = dict[workshopid[i]]
            a[6].append(email[i])
            dict[workshopid[i]] = a

    mail_host = 'smtp.office365.com'
    mail_user = '*'
    mail_pass = '*'
    for id in dict:
        bcc = testemailaddress
        # Check!!!
        bcc = dict[id][6]
        receivers = bcc
        meetlink=''

        if len(meetlink) == 0:
            descriplist = dict[id][7].split('\n')
            for i in descriplist:
                j = i.split(' ')
                for z in j:
                    if 'semo.zoom.us/' in z:
                        meetlink = 'http://'+z[z.find('semo'):-1]
        string = "Dear faculty, \n\nYou will get an email reminder for each workshop you are enrolled in.\n\nYou have been enrolled in the following workshop: "\
                                                                              "\nTITLE: "+dict[id][0]+"\nDATE: "+dict[id][4]+" "+dict[id][5]+\
                 "\n"+'Presenter: '+dict[id][2]+\
                 "\nLocation: Zoom - Use the following Zoom link to access the session: "+meetlink+\
                "\n\nPlease remember to attend! We look forward to seeing you there."\
                "\n\nSincerely,\nCTL"

        message = MIMEText(string)

        subject ='Workshop Reminder '+dict[id][4]
        message['Subject'] = Header(subject, 'utf-8')

        try:
            smtpObj = smtplib.SMTP(mail_host, 587)
            smtpObj.connect(mail_host, 587)
            smtpObj.ehlo()
            smtpObj.starttls()
            smtpObj.ehlo()
            smtpObj.login(mail_user, mail_pass)
            smtpObj.sendmail(mail_user, receivers, message.as_string())
            print("Email Sent Successfully")
        except smtplib.SMTPException as e:
            print("Error: ", e)

reminder()
certificate1()