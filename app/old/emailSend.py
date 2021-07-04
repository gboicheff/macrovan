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

#path = "../macrovan/io/Output/"
path = "../io/Output\PDFs/email_send/"
inputPath = "../macrovan/io/Input/"
emailSubject = "Turf PDF"
sender_address = email_address
sender_pass = email_password
print(f'START: os.getcwd = {os.getcwd()}')
# Set this to False to actually send the emails
testMode = False
# Set this to true to send all the emails out without stepping. BE CAREFUL WITH THIS
dont_want_to_watch = True

#with open("app/email_body.txt", "r") as body:
with open("email_body.txt", "r") as body:
    email_body = body.read()

def initialize_session():
    if not testMode:
        session = smtplib.SMTP('smtp.gmail.com', 587) 
        session.starttls() 
        session.login(sender_address, sender_pass)
        return session
    else:
        print("Session not started.  Test mode is on.")


def create_email(receiver_addresses, filenames, cc_list, turf):
    # organizer_attachment_dict = {
    #     "janeathom@aol.com" : "Thomas QSG Nov2020.pdf",
    #     "andybragg@me.com" : "Bragg QSG Nov2020.pdf",
    #     "ssinger1313@gmail.com" : "SingerSkipper QSG Nov2020.pdf",
    #     "kdsteinway@gmail.com" : "Steinway QSG Nov2020.pdf",
    #     "ezrasinger@gmail.com" : "SingerEzra QSG Nov2020.pdf",
    #     "dave@weegallery.com" : "Pinto QSG Nov2020.pdf",
    #     "stephenpeeples@mac.com" : "Peeples QSG Nov2020.pdf",
    #     "mariamckay46@gmail.com" : "McKay QSG Nov2020.pdf"
    #     # "mmckay23@tampabay.rr.com" : "McKay QSG Nov2020.pdf"
    # }
    list_dict = get_pdf_info(find_file(filenames[0], True))
    date = list_dict['date_generated']
    date_1 = datetime.datetime.strptime(date, "%m/%d/%y")
    dt = date_1 + datetime.timedelta(days=30)
    end_date = '{0}/{1}/{2:02}'.format(dt.month, dt.day, dt.year % 100)
    message = MIMEMultipart()
    message['From'] = sender_address
    message['Subject'] = "Your PDF Named: " + list_dict['pdf_file_name'] + ", Expires On: " + end_date
    list_number = " - ".join(list_dict['list_number'].split("-"))
    doors = list_dict['door_count']
    people = list_dict['person_count']
    if turf['organizer_phone'] == 0:
        phone = ""
    else:
        phone = turf['organizer_phone']
    if pd.isnull(turf['total_voters']):
        total_voters = turf['total_voters'] = "N/A"
    else:
        total_voters = int(turf['total_voters'])
    if not pd.isnull(turf["first_name"]):
        turf["first_name"].replace(" ", "")
    body = email_body.format(bc_first_name=turf['first_name'].capitalize(), turf_name=turf['turf_name_in_van'], list_number=list_number, doors=doors, people=people,
    organizer_name=turf['organizer_name'], organizer_phone=phone, total_voters=total_voters, expr_date=end_date,organizer_email=turf['organizer_email_address'])
    message['To'] = ",".join(receiver_addresses)
    if(len(cc_list) > 0):
        message['Cc'] = ",".join(cc_list)
    message.attach(MIMEText(body, 'plain'))
    numAttachedFiles = attachpdfs(filenames, message)
    # if turf['organizer_email_address'].rstrip() in organizer_attachment_dict:
    #     print(organizer_attachment_dict[turf['organizer_email_address'].rstrip()])
    #     attach_files([organizer_attachment_dict[turf['organizer_email_address'].rstrip()]],message)
    # else:
    #     print("Not found!: " + turf['organizer_email_address'].rstrip())
    if(numAttachedFiles == len(filenames)):
        return message
    else:
        print("Failed to create email for " + turf['first_name'] + " " + turf['last_name'] + "." + "  Email will not be sent.")
        return False

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
        print("Test mode is on.  Email not sent")

#Extract info from a single pdf
#def get_pdf_info(file_name, path=r'io\Output'):
def get_pdf_info(file_name, path=r'..\io\Output\PDFs\email_send'):
    pdfFileObj = open(path + '\\' + file_name, 'rb')
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


def attachpdfs(file_names, email):
    numFiles = 0
    for file_to_attach in file_names:
        expectedFileName = file_to_attach + ".pdf"
        fileName = file_to_attach
        fileName = fileName.replace(" ", "")
        fileName+="*" + ".pdf"
        outFileName = file_to_attach + ".pdf"
        if not testMode:  
            # Search the directory for a file that matches the fileName
            for file in os.listdir(path):
                foundFile = file.replace(" ", "")
                if fnmatch.fnmatch(foundFile, fileName):
                    pdf = MIMEApplication(open(path + file, 'rb').read())
                    pdf.add_header('Content-Disposition','attachment', filename=outFileName)
                    email.attach(pdf)
                    numFiles+=1
                    print("Attached " + file)
                    break
                    
        # Testing mode
        else:
            foundFile = "FILE NOT FOUND-------------------------------------F"
            for file in os.listdir(path):
                foundFile = file.replace(" ", "")
                foundFileName = "NOT FOUND"
                if fnmatch.fnmatch(foundFile, fileName):
                    foundFileName = file
                    test = open(path + file, 'rb')
                    test.close()
                    numFiles += 1
                    break
            # Expected on left, found on right
            print("Expected: " + expectedFileName + "    :    " + "Found: " + foundFileName)
    return numFiles

def attach_files(file_names, email):
    numFiles = 0
    for file_name in file_names:        
        if not testMode:  
            for file in os.listdir(inputPath):
                if fnmatch.fnmatch(file, file_name):
                    actual_file = MIMEApplication(open(inputPath + file, 'rb').read())
                    actual_file.add_header('Content-Disposition','attachment', filename=file_name)
                    email.attach(actual_file)
                    numFiles+=1
                    print("Attached " + file)
                    break                  
    return numFiles


#Locate a file in the output folder.  Can toggle ignoring spaces
def find_file(filename, ignore_spaces):
    filename = filename + "*" + ".pdf"
    if ignore_spaces:
        filename = filename.replace(" ", "")
    for file in os.listdir(path):
        if ignore_spaces:
            foundFile = file.replace(" ", "")
        if fnmatch.fnmatch(foundFile, filename):
            return file
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
        # "MMCKAY23@tampabay.rr.com" : "Maria McKay",
        "mariamckay46@gmail.com" : "Maria McKay",
        "Srogers1080@outlook.com" : "Sara Rogers",
        "ssinger1313@gmail.com" : "Skipper Singer",
        "stephenpeeples@mac.com" : "Steve Peeples",
        "teamvote2020@gmail.com" : "Amy Walsh"
    }
    return organizer_dict[email]

def send_files():
    #dev_cc_list = ["gboicheff@gmail.com", "mjturtora@gmail.com", "fahygotv@gmail.com", "dave@weegallery.com", "stephenpeeples@mac.com"]
    dev_cc_list = ["gboicheff@gmail.com", "mjturtora@gmail.com", "stephenpeeples@mac.com"]
    #dev_cc_list = ["mjturtora@gmail.com", "gboicheff@gmail.com"]
    print("==================================================")
    #fname=r"C:\Users\Grant\Desktop\macrovan\io\Input\Nov 2020 -Tracking All Voters.xlsx"
    fname = r"D:\Stuff\Projects\Pol\macrovan\io\Input\Nov 2020 -Tracking All Voters - blocked.xlsx"

    turfs = get_volunteer_data(fname)
    session = initialize_session()
    success = True
    sent_count = 0
    sent_list = []
    sent_file = open("emails.txt", "w")
    for turf in turfs:
        time.sleep(1.5)
        print("-------------------------------------------")
        if not pd.isnull(turf['email_to_bc']) and turf['email_to_bc'] == "y" and not pd.isnull(turf['organizer_email_address']) and not turf['organizer_email_address'] == "" and not pd.isnull(turf['email_address']) and not turf['email_address'] == "" and not pd.isnull(turf['turf_name_in_van']) and not turf['turf_name_in_van'] == "":
            print(turf['email_address'])
            first_name = turf['first_name'].rstrip().lstrip().replace(" ", "")
            last_name = turf['last_name'].rstrip().lstrip().replace(" ", "")
            turf_name = turf['turf_name_in_van']
            turf['organizer_name'] = get_organizer_name(turf['organizer_email_address'].rstrip().lstrip().replace(" ", ""))
            receiver_address = turf['email_address'].rstrip().lstrip().replace(" ", "")
            organizer_email = turf['organizer_email_address'].rstrip().replace(" ", "")

            #receiver_address = 'gboicheff@gmail.com'
            # organizer_email = 'gboicheff@gmail.com'
            #receiver_address = 'mjturtora@gmail.com'
            #organizer_email = 'm@notmail.com'  #'mjturtora@gmail.com'

            final_cc_list = dev_cc_list + [organizer_email]
            filename = turf_name
            print(turf_name)
            print("Send email to " + first_name + " " + last_name + " at " + receiver_address)
            print("CCing: " + str(final_cc_list))
            print("Expected filename: " + filename)
            foundFile = find_file(filename, True)
            print("Found filename: " + find_file(filename, True))
            if dont_want_to_watch or input_choice():
                email = create_email([receiver_address], [filename], final_cc_list, turf)
                if not testMode and foundFile != "NOT FOUND":
                    if email != False:
                        all_to_addresses = [receiver_address] + final_cc_list
                        if send_email(all_to_addresses, email, session) != False:
                            print("Email to " + first_name + " " + last_name + " sent")
                            sent_list += [first_name + " " + last_name + " " + receiver_address + " EMAIL SENT" + "filename: " + foundFile]
                            sent_count+=1
                        else:
                            print("Email to " + first_name + " " + last_name + " not sent")
                            sent_list += [first_name + " " + last_name + " " +  receiver_address + " EMAIL NOT SENT"]
                            success = False
                    else:
                        success = False
                        print("Email to " + first_name + " " + last_name + " not sent")
                        sent_list += [first_name + " " + last_name + " " + receiver_address + " EMAIL NOT SENT"]
                else:
                    sent_count+=1
                    sent_list += [first_name + " " + last_name +  "  " + receiver_address + " EMAIL NOT SENT"]
            else:
                print("Email to " + first_name + " " + last_name + " not sent")
                sent_list += [first_name + " " + last_name + " " + receiver_address + " EMAIL NOT SENT"]
                success = False
            sent_file.write(turf_name + "                :                " + foundFile + "                :                " + last_name + "\n")
    if not testMode:
        session.quit()
    if success:
        print("All " + str(sent_count) + "emails sent!")
    else:
        print("At least one email not sent!")
    print("==================================================")
    sent_file.close()
    for entry in sent_list:
        print(entry)
    

if __name__ == '__main__':
    send_files()
