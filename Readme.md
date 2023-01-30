# Automated Email for Birthdays


## Please read
**A**   
1. Make sure the files below files are present in folder
   1. Config.ini
   2. email-gsheet-reminder-bb1bf4ee109e.json 
2. Go through the below link to get a password for *emailPwd*.
https://support.google.com/mail/answer/185833?hl=en-GB
3. Birthday cells in Excel have to be in DD/MM/YYYY format with leading zeros
   
---

## So how to make it work?
**B**
1. Prepare your excel sheet
2. Open config file(configfile.ini)
   1. Set staticPosCell(Col) and iterativePosCell (row) for *Birthday* cell
        Enter the position of beginning cell with values under the *Birthday* column. E.g. Cell B2 (Row = 2, Column = 2) might be where the first values of the column *Birthday* begins.
   2. Set staticFinCell (Column) for *Name* cell
        Enter the position of the beginning cell with values under the *Name* column. E.g. Cell A2 (Row = 2, Column = 1) might be where the first values of the column *Name* begins. Only the Column value has to be noted down.
   3. Set sheetname as the name of the excel sheet
   4. Set emailID and emailPWD(Do not use personal password; refer to **A2**) for sender
   5. Set emailSender for recepient
3. Share *editor* access of the sheet to the below email ID.
test-sheet-for-email-reminder@email-gsheet-reminder.iam.gserviceaccount.com
4. Run it
