import xlrd
import smtplib as sm


book = xlrd.open_workbook('emails.xlsx')
sheet = book.sheet_by_index(0)
rows = sheet.nrows
cols = sheet.ncols

email_list = [sheet.cell_value(r,0).encode('ascii') for r in range(rows)]
sub = sheet.cell_value(0,1).encode('ascii')
body = sheet.cell_value(0,2).encode('ascii')
smObj = sm.SMTP('smtp.gmail.com', 587)
smObj.ehlo()
smObj.starttls()
smObj.login(input('enter email: '), input('enter password: '))

# Sending Emails from a excel sheet containing subject and body along with email addresses 
for i in range(rows):
    smObj.sendmail('<your email address>', str(email_list[i]), 'Subject: ' + str(sub)+ '\n' + str(body))

smObj.quit()

#remember to give a 2 or more second delay every 20 emails and a one day delay after 500 emails (google's policy)
