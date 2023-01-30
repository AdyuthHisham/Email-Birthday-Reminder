####Remind HR of events



import pygsheets
import time
from datetime import date
import configparser
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import configparser

print("Modules imported")


config = configparser.ConfigParser()
config.read('configfile.ini')
###Excel
##Static position of cell to be searched for
scurPos = int(config['ExcelExtract']['staticPosCell'])
##Iterative position of cell to be searched for
icurPos = int(config['ExcelExtract']['iterativePosCell'])
##Static position of cell to get output from
finPos = int(config['ExcelExtract']['staticFinCell'])
sheetName = config['ExcelExtract']['sheetName']

###Mail
EmailID = config['MailExtract']['emailID']
EmailTo = config['MailExtract']['emailRec']
EmailPWD = config['MailExtract']['emailPwd']
subject = config['MailExtract']['subject']
text = config['MailExtract']['text']

print("Config data initialized")


###Globul Variables

searchVal = date.today()
searchVal = searchVal.strftime("%d/%m")

print("Global variables initialized")

flag = 0
###Excel
#authorization
gc = pygsheets.authorize(service_file='email-gsheet-reminder-bb1bf4ee109e.json')
#open the google spreadsheet (where 'PY to Gsheet Test' is the name of my sheet)
sh = gc.open(sheetName)
#select the first sheet 
wks = sh[0]
#Set initial cell position
curCell = wks.cell((icurPos,scurPos))
neighCell = wks.cell((finPos,scurPos))
###Mail
smtp = smtplib.SMTP('smtp.gmail.com', 587)
smtp.ehlo()
smtp.starttls()
smtp.login(EmailID, EmailPWD)

msg = MIMEMultipart()
msg['Subject'] = subject
msg.attach(MIMEText(text))

print("Program requirements initialized")


while curCell.value != '':
    if searchVal.__eq__(curCell.value[0:5]):
        flag =1
        neighCell = wks.cell((icurPos,finPos))
        msg.attach(MIMEText(' <> '+neighCell.value))
    icurPos = icurPos + 1
    curCell = wks.cell((icurPos,scurPos))
    time.sleep(0.2)

if flag == 1:
    to = [EmailTo]
    smtp.sendmail( from_addr=EmailID , to_addrs=to , msg=msg.as_string())
    print("Mail sent")
    smtp.quit()

print("EOF")
