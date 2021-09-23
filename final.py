import smtplib
from email.message import EmailMessage
import sys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
#import section ends

mail_error=list() #global variable that will store all the details of the mails not sent to people.
gmail=""
password=""#These two values are to be provided
name_list=list()
mail_list=list()#Global variables that will be updated from the values of the sheet at the start.
#Data values from the sheet are only updated at the start so that we don't have unnecesary multiple API requests.
mail_content=""#Global variable to store the message content in. Can be text or HTMl, will differentiate accordingly.
mail_error=list()#Global list to store the names of the people to whom mails were not sent for some reason.
mail_subject=''

def setEmailPassword(email, pw):
    global gmail,password
    gmail=email
    password=pw


def get_sheet_data(url,sheet_name,first_column,second_column):
    global name_list,mail_list
    try:
        scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
        creds=ServiceAccountCredentials.from_json_keyfile_name("HanEm-API-creds.json", scope)
        client= gspread.authorize(creds)
    except:
        print("API authentication failed")
        return 1
    try:
        wb=client.open_by_url(url)
    except:
        print("Please check URL and enable view access to everyone through URL.")
        return 1
    try:
        sheet=wb.worksheet(sheet_name)
    except:
        print("Please check sheet name. No such sheet exists in the workbook.")
        return 1

    data=sheet.get_all_values()
    df = pd.DataFrame(data)
    df.columns = df.iloc[0]
    df = df.iloc[1:]

    try:
        name_list=list(df[first_column])
        mail_list=list(df[second_column])
    except:
        print("Please check the column names that you have provided.")
        return 1

    
def get_content(file_name):
    global mail_content
    try:
        file=open(file_name)
        mail_content=file.read()
        file.close()
    except:
        print("Error in reading file.")
        return 1
    

def print_data(user_list):
    for i in user_list:
        print(i)
    print("Do you want to proceed with the next batch?")
    inp=input("Enter(y/n):")
    if inp.lower() == 'y':
        return
    else:
        exit()

def send_mail():
    global name_list,mail_list,mail_content,gmail,password,mail_error,mail_subject,gmail,password
    if ((gmail==None) or (password==None)):
        print("Email and password not provided.")
        return
    
    with smtplib.SMTP('smtp.gmail.com',587) as smtp:
        smtp.starttls()
        smtp.login(gmail,password)
        for i in range(0,len(name_list)):
            msg=EmailMessage()
            name=name_list[i]
            address=mail_list[i]
            msg['From']=gmail
            msg['To']=address
            msg['Subject']=mail_subject
        
            if '<div>' in mail_content:
                msg.add_header('Content-type','text/HTML')
                msg.set_payload(mail_content)
            else:
                msg.set_content(mail_content)

            try:
                smtp.send_message(msg)
                print(f'Mail sent to {name} at {address}')
            except:
                print(f'Mail FAILED to send to {name} at {address}.')
                mail_error.append([name,address,name_list.index(name)])

            
            