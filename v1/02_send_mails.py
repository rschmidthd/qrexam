# -*- coding: utf-8 -*-
"""
Created on Wed Jul 22 14:13:53 2020

@author: ufriess
"""

import pandas as pd
import numpy as np
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

subject = 'Nachklausur Experimentalphysik 2 - Wichtige Informationen'

mail_text = open ('./stud_mail.txt', 'r', encoding='utf-8').read()
    
##### SMTP Mail server info

smtp_user = ''
smtp_passwd = ''
smtp_server = 'smtp.iup.uni-heidelberg.de'
smtp_port = 25

smtp_from = 'udo.friess@iup.uni-heidelberg.de'

# Lese Teilnehmer
teilnehmer = pd.read_excel('klausurteilnehmer.xls', header = 0,)
#teilnehmer = pd.read_excel('mail_test.xls', header = 0,)

if teilnehmer.iloc[0].Nachname is np.NaN: # first row is empty
    teilnehmer.drop(0, inplace = True)
    
teilnehmer.UID = np.nan_to_num(teilnehmer.UID).astype(int)
teilnehmer.Matrikel = np.nan_to_num(teilnehmer.Matrikel).astype(int)
teilnehmer.Gruppe = teilnehmer.Gruppe.astype(int)

teilnehmer.sort_values(by = ['Raum', 'Sitzplatz'], inplace = True)

# teilnehmer = teilnehmer.iloc[100:103]
# teilnehmer['E-mail'] = 'udo.friess@iup.uni-heidelberg.de'

# Verschicke mails
server = smtplib.SMTP(smtp_server, smtp_port)
server.ehlo()
#    server.starttls()
server.ehlo()
if smtp_user != '':
    server.login(smtp_user, smtp_passwd)

n=0 
for i, stud in teilnehmer.iterrows():
    message = mail_text.format(Vorname = stud.Vorname, 
                           Nachname = stud.Nachname,
                           Matrikel = stud.Matrikel,
                           Raum = stud.Raum,
                           Sitzplatz = stud.Sitzplatz
                           )
    msg = MIMEMultipart()
    msg['From'] = smtp_from
    msg['To'] = stud['E-mail']

    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))
    
    server.sendmail(smtp_from, stud['E-mail'], msg.as_string())
   
    n+=1
    print(n, stud['E-mail'])

server.quit()
    
    

    