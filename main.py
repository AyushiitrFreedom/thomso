import comtypes.client
import sys
import csv
import collections
import collections.abc
import os
import smtplib
from email.message import EmailMessage
import glob
import os
import tqdm
import time
import os
import sys
import comtypes.client


company_name = ""
email1 = ""
email2 = ""
email3 = ""
email4 = ""

with open('data.csv', 'r', encoding='utf8') as csv_file:
    csv_reader = csv.reader(csv_file)
    next(csv_reader)
    for index, line in enumerate(csv_reader):
        print("Index:", index)
        company_name = line[7].strip()
        email1 = line[16].strip()
        email2 = line[19].strip()
        email3 = line[22].strip()
        email4 = line[25].strip()
        if (company_name == ""):
            continue

        EMAIL_ADDRESS = "ayushgoyal.thomso@gmail.com"
        EMAIL_PASSWORD = "itxrffqjdaxaggww"

        contacts = [email1, email2, email3, email4]

        msg = EmailMessage()
        msg['Subject'] = f'Proposal for association between {company_name} and Thomso, IIT Roorkee '
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = contacts
        msg['Cc'] = "chinmay.thomso@gmail.com"

        msg.set_content(f''' Dear Sir/Madam,
            Greetings from Thomso, IIT Roorkee!
            We are honoured to invite {company_name} as a Sponsor for Thomso '23, the annual cultural fest of IIT Roorkee. It is our marquee event, which attracts over 1,00,000 people every year. This time around it’s going to be bigger and better than ever before! We have been proud hosts to great personalities like Mrs. Smriti Irani (as our Chief Guest), Sonu Nigam, Darshan Rawal, Farhan Akhtar, Salim-Sulaiman, and Jubin Nautiyal (check it out: https://youtu.be/h7gyJRWrjbg )in our previous editions (Thomso ’22: https://youtu.be/rm1bWDAHbSQ ). Kindly check the attached brochure for more information about Thomso.
            With several zonal events lined up at premium locations like Delhi, Lucknow, Jaipur, Chandigarh, Bangalore, and Ahmedabad, Thomso's reach is more than 2,00,000 people through various Newspapers, Magazines, Social Media Platforms, Radio and T.V. Channels, Online Blogs, etc.
            This is a unique opportunity for {company_name} to engage, to introduce itself to prospective customers as well as promote growth strategies with existing clientele. The attached proposal  contains details regarding sponsorship prospects, deliverables, etc.
            All the Deliverables and related things are negotiable from both parties and can be discussed further on mail or call.
            We have come a long way but would still greatly benefit from your support. With immense participation and massive outreach, this affiliation will prove to be mutually beneficial.

            Inviting you to be a part of our "THOMSO" family of IIT ROORKEE.
            ''')

        files = ['Brochure.pdf', 'Proposal.pdf']
        for file in files:
            with open(file, 'rb') as f:
                file_data = f.read()
                file_name = f.name

            msg.add_attachment(file_data, maintype='application',
                               subtype='octet-stream', filename=file_name)
            msg.add_header('X-Unsent', '1')

            print("attachment added")
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                smtp.send_message(msg)
