import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from secrets import *
import os
import fnmatch
import datetime
from utils import *
import shutil
import time
import json

main_path = os.getcwd()
app_path = os.path.join(main_path, "app")
output_path = os.path.join(main_path, "io", "Output")
input_path = os.path.join(main_path, "io", "Input")

sender_address = email_address
sender_pass = email_password

# Set this to False to actually send the emails
testMode = False
# Set this to true to send all the emails out without stepping. BE CAREFUL WITH THIS
dont_want_to_watch = True


def read_email_body(file_name):
    file_name = find_file(file_name, app_path, "txt")
    with open(file_name, "r") as body:
        email_body = body.read()
    return email_body

def initialize_session():
    if not testMode:
        session = smtplib.SMTP('smtp.gmail.com', 587) 
        session.starttls() 
        session.login(sender_address, sender_pass)
        return session
    else:
        print("Session not started.  Test mode is on.")

#format the date extracted from the PDF
def format_date(raw_date):
    date_1 = datetime.datetime.strptime(raw_date, "%m/%d/%y")
    dt = date_1 + datetime.timedelta(days=30)
    return '{0}/{1}/{2:02}'.format(dt.month, dt.day, dt.year % 100)

def insert_turf_email(email_body, turf, list_dict, end_date):
    list_number = " - ".join(list_dict['list_number'].split("-"))
    body = email_body.format(bc_first_name=turf['first_name'].capitalize(), turf_name=turf['turf_name_in_van'], list_number=list_number, doors=list_dict['door_count'], people=list_dict['person_count'],
    organizer_name=turf['organizer_name'], organizer_phone=turf['organizer_phone'], total_voters=turf['total_voters'], expr_date=end_date,organizer_email=turf['organizer_email_address'])
    return body

def create_email(receiver_addresses, filenames, cc_list, turf, email_body_file_name):
    list_dict = get_pdf_info(find_file(filenames[0], output_path))
    end_date = format_date(list_dict['date_generated'])
    message = MIMEMultipart()
    message['From'] = sender_address
    message['Subject'] = "Your PDF Named: " + list_dict['pdf_file_name'] + ", Expires On: " + end_date
    list_number = " - ".join(list_dict['list_number'].split("-"))
    email_body = read_email_body(email_body_file_name)
    body = insert_turf_email(email_body, turf, list_dict, end_date)
    message['To'] = ",".join(receiver_addresses)
    message['Cc'] = ",".join(cc_list)
    message.attach(MIMEText(body, 'plain'))
    message = attach_files(filenames, output_path, message)
    return message

def send_email(receiver_addresses, email, session):
    if not testMode:
        text = email.as_string()
        try:
            session.sendmail(sender_address, receiver_addresses, text)
        except:
            return False
        else:
            return True
    else:
        return False

#Extract info from a single pdf
def get_pdf_info(file_name, path=r'io\Output'):
    pdfFileObj = open(file_name, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    page = pdfReader.getPage(0).extractText()
    first_part, doors = page.split("Doors:", 1)
    date, people = page.split("People:", 1)
    date = date.split("Generated")[1]
    date = date.split(" ")[1]
    doors = int(doors.split("Affiliation")[0])
    people = int(people.split("Affiliation")[0].split()[0])
    page = pdfReader.getPage(2).extractText()
    if people != 0:
        pdf_file_name, lnum = page.split("List", 1)
        lnum = lnum.split(" ")[1]
    else:
        lnum = '0-0'
        pdf_file_name, date_part = file_name.split("_2020", 1)
    pdf_dict = {
        'list_number': lnum,
        'door_count': doors,
        'person_count': people,
        'date_generated': date,
        'pdf_file_name': pdf_file_name,
    }
    return pdf_dict


def attach_files(file_names, path, email):
    for file in file_names:
        found_file = find_file(file, path)
        if not testMode: 
            pdf = MIMEApplication(open(found_file, 'rb').read())
            pdf.add_header('Content-Disposition','attachment', filename=found_file)
            email.attach(pdf)               
    return email



#Locate a file in the output folder.  Can toggle ignoring spaces
def find_file(file_name, path, file_type="pdf", ignore_spaces=True):
    search_file_name = file_name + "*" + "." + file_type
    if ignore_spaces:
        search_file_name = search_file_name.replace(" ", "")
    for file in os.listdir(path):
        if ignore_spaces:
            foundFile = file.replace(" ", "")
        if fnmatch.fnmatch(foundFile, search_file_name):
            return os.path.join(path, file)
        # if file_name == file:
        #      return path + file
    return "NOT FOUND"            

def input_choice():
    print("Enter (Y/N):")
    choice = input()
    if(choice == "Y" or choice == "y"):
        return True
    elif(choice == "N" or choice == "n"):
        return False
    else:
        print("Please enter (Y/N):")
        return input_choice()

#Name lookup because new sheet doesn't contain organizer names
def get_organizer_name(email):
    organizer_dict = {
        "andybragg@me.com" : "Andy Bragg",
        "barb.law2020@gmail.com" : "Barb Law",
        "bryantranslations@gmail.com" : "Bryan Casanas",
        "charles.walston@gmail.com" : "Charles Walston",
        "dave@weegallery.com" : "Dave Pinto",
        "dbrownallen@gmail.com": "Debbie Allen",
        "ezrasinger@gmail.com" : "Ezra Singer",
        "freeusa68@gmail.com" : "Sharon Loughry",
        "janeathom@aol.com" : "Jane Thomas",
        "kdsteinway@gmail.com" : "Kate Steinway",
        "mariamckay46@gmail.com" : "Maria McKay",
        "Srogers1080@outlook.com" : "Sara Rogers",
        "ssinger1313@gmail.com" : "Skipper Singer",
        "stephenpeeples@mac.com" : "Steve Peeples",
        "teamvote2020@gmail.com" : "Amy Walsh",
        "gboicheff@gmail.com"   : "Grant"
    }
    return organizer_dict[email]

def clean_entry(entry, spaces_replace=True):
    if type(entry) is str:
        entry = entry.rstrip().lstrip()
        if spaces_replace:
            entry = entry.replace(" ", "")
        if entry == "" or pd.isnull(entry):
            return "N/A"
        else:
            return entry
    return entry

def clean_turf_date(turf):
    for detail in turf.keys():
        # print(detail + " " + str(type(turf[detail])))
        detail = clean_entry(turf[detail])
    return turf

def display_email_details(turf, file_name, found_file, final_cc_list):
    print("Send email to " + turf['first_name'] + " " + turf['last_name'] + " at " + turf['email_address'])
    print("CCing: " + str(final_cc_list))
    print("Expected filename: " + file_name)
    print("Found filename: " + found_file)

def write_results(turfs):
    with open("email_results" + time.strftime("%Y%m%d-%H%M%S") +".json", "w") as result:
        json.dump(turfs, result, indent=3)

def check_sent(turfs):
    for turf in turfs:
        if not bool(turf['email_sent']):
            print("EMAIL FAILED TO SEND TO " + turf['email_address'])


def send_files(dev_cc_list=["gboicheff@gmail.com"]):
    print("==================================================")
    turfs = get_volunteer_data()
    session = initialize_session()
    sent_list = []
    for turf in turfs:
        print("-------------------------------------------")
        if not pd.isnull(turf['email_to_bc']) and turf['email_to_bc'] == "y" and not pd.isnull(turf['organizer_email_address']) and not turf['organizer_email_address'] == "" and not pd.isnull(turf['email_address']) and not turf['email_address'] == "" and not pd.isnull(turf['turf_name_in_van']) and not turf['turf_name_in_van'] == "":
            final_cc_list = dev_cc_list + [turf['organizer_email_address']]
            file_name = turf['turf_name_in_van']
            found_file = find_file(file_name, output_path, "pdf")
            turf['found_file_name'] = found_file
            turf['organizer_name'] = get_organizer_name(turf['organizer_email_address'])
            turf = clean_turf_date(turf)
            display_email_details(turf, file_name, found_file, final_cc_list)
            if dont_want_to_watch or input_choice():
                if not testMode:
                    time.sleep(1.5)
                email = create_email([turf['email_address']], [file_name], final_cc_list, turf, "email_body")
                all_to_addresses = [turf['email_address']] + final_cc_list
                success = send_email(all_to_addresses, email, session)
                turf['email_sent'] = success
            else:
                turf['email_sent'] = False
            sent_list.append(turf)
    if not testMode:
        session.quit()
    print("==================================================")
    write_results(sent_list)
    check_sent(sent_list)
    

if __name__ == '__main__':
    send_files()
    # send_files(["gboicheff@gmail.com", "mjturtora@gmail.com", "fahygotv@gmail.com", "dave@weegallery.com", "stephenpeeples@mac.com"])
