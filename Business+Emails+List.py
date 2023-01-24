# importing needed Python  Modules
from smtplib import SMTP 
import pytz
import datetime
import dateutil
import numpy as np
import pandas as pd


Sender = 'python.practice.bot@gmail.com'
Password = 'yetcezmcvxueacvl'
now = datetime.datetime.now()

actual_date = (now.day, now.month)

#Read from Excel sheet using Pandas Module
Emails_list = pd.read_excel('emails_list.xlsx')

#Extract the customers names and the customers emails from the excel sheet as a list
names = Emails_list['Name']
emails = Emails_list['Email']

# Connect to the server of the sender email account
with SMTP(host = 'smtp.gmail.com', port = 587) as connect_server:

    # Encrypt our emails using tls protocol
	connect_server.starttls()
    # Log in the Sender Account
	connect_server.login(Sender, Password)
    # iterate through records to get each name and email
	for item in range(len(emails)):

		name = names[item]
		email = emails[item]
    # Adjust the time of sending emails to customers using control flow if statement
		if actual_date == (20, 11):
			message = "Subject:Black Friday Discount Offer\n\nHi, {}, Don't wait and buy now.".format(name)
			connect_server.sendmail(Sender, [email], message)
		elif actual_date == (20, 12):
			message = "Subject:White Friday Discount Offer\n\nHi, {}, Don't wait and buy now.".format(name)
			connect_server.sendmail(Sender, [email], message)
		# Testing our code by sending emails and print statements
		elif actual_date == (5, 12):
			message = "Subject:Discount Offer\n\nHi, {}, Code is okay.".format(name)
			connect_server.sendmail(Sender, [email], message)
			print('Code testing result is okay.')

	
print('All Done.')