# pip install pandas
# pip install openpyxl
# pip install datetime
# pip install regex
# pip install smtplib

import pandas as pd
from datetime import datetime
import re 
import smtplib
from keys import *

# return true if birthday
def checkBirthday():
    data = pd.read_excel(r'birthdays.xlsx')
    df = pd.DataFrame(data)

    Name = df['Name'].tolist()
    Email = df['Email'].tolist()
    Birthdate = df['Birthdate'].tolist()
    Age = df['Age'].tolist()

    # populating the data to a list for easy referencing later
    all_data = []
    for name,email,birthdate,age in zip(Name,Email,Birthdate,Age):
        all_data.append([name,email,birthdate,age])
        # print(all_data)
    
    # date and time right now
    now = datetime.today().strftime("%m-%d")
    #print(formatted_date)

    # checking if the birthdate matches with today's date
    for i in all_data:
        individualBirthdate = str(i[2])
        # formatting the date using the regex library
        pattern = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
        results = re.search(pattern, individualBirthdate)
        # grouping the date and the month of his birthdate
        date = f"{results.group(2)}-{results.group(3)}"

        # return names, email's of people who have their birthday today
        wishThem = []
        if date == now:
            wishThem.append([i[0],i[1]])

        return wishThem

def wish():
    birthdays = checkBirthday()
    
    # send a birthday email to every person in the list
    for i in range(len(birthdays)):
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()

            smtp.login(EMAIL_ADDRESS,PASSWORD)
                        
            msg = f'''
            Happy Birthday {birthdays[i][0]}!

            - Sent via Anish's programmed assistant. Thank you!
            '''.encode('utf-8')
                        
            smtp.sendmail(EMAIL_ADDRESS, birthdays[i][1], msg)

wish()