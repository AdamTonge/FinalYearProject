#list of imports
import imaplib
import email
import datetime
from nltk.tokenize import word_tokenize
from nltk.corpus import words
from nltk.corpus import stopwords
import pickle

userEmail = input("Whats your email?")
userPass = input("Whats your password?")

#connect to outlook securely
mail = imaplib.IMAP4_SSL('imap.outlook.com', 993)

# try to login using some user inputted data
try:
    mail.login(userEmail, userPass)
except imaplib.IMAP4.error:
    print("login failed")

# list the folders in the email and then pick 'inbox'
mail.list()
mail.select('inbox')

#search though the inbox and separate each individual email then print it to myfile
result, data = mail.uid('search', None, "ALL")

#initialize an empty dictionary
my_dict = {}

#set userTimeLower as time + 30 minutes and set userTimeUpper as the time right now
userTimeLower = datetime.datetime.now() + datetime.timedelta(minutes=30)
userTimeLower = userTimeLower.strftime("%H:%M")
userTimeUpper = datetime.datetime.now()
userTimeUpper = userTimeUpper.strftime("%H:%M")

# add words to a set from the nltk lib
english_words = set(words.words())
stop_words = set(stopwords.words())

#get the length of the uids
i = len(data[0].split())

for x in range(i):
    filtered_email = []
    #fetch the email by its uid then decode
    latest_email_uid = data[0].split()[x]
    result, email_data = mail.uid('fetch', latest_email_uid, '(RFC822)')   ###remove duplicated code with result , data
    raw_email = email_data[0][1]

    raw_email_string = raw_email.decode('utf-8', errors='ignore')
    email_message = email.message_from_string(raw_email_string)

    #get the date and parse it to just get the hr and minutes
    emailTime = email_message['Date']
    emailTime = emailTime.split(' ')
    emailTime = emailTime[4]
    emailTime = emailTime[0:5]

    #get body of email
    for part in email_message.walk():
        #check whether the time of the email is between now and 30 minutes in the future
        if(emailTime < userTimeLower and emailTime > userTimeUpper):
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True)
                #body = body.encode('utf-8')
                body = str(body, "ISO-8859-1")

                #parse the words from the emails
                words = word_tokenize(body)

                #filter out any non english words then stopwords
                for w in words:
                    if w in english_words:
                          if w not in stop_words:
                                filtered_email.append(w)

            else:
                continue

    #my_dict[time] = filtered_email

    #check if filtered_email is empty
    if filtered_email:
        print(emailTime)
        print("Keywords from email are:")
        print(filtered_email)
        break

mail.close()
mail.logout()

