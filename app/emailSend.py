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

path = "../macrovan/io/Output/"
emailBody = '''Your PDFs are attached'''
emailSubject = "Turf PDF"
sender_address = email_address
sender_pass = email_password

# Set this to False to actually send the emails
testMode = True
print(os.listdir("."))
with open("app/email_body.txt", "r") as body:
    email_body = body.read()

def initialize_session():
    if not testMode:
        session = smtplib.SMTP('smtp.gmail.com', 587) 
        session.starttls() 
        session.login(sender_address, sender_pass)
        return session
    else:
        print("Session not started.  Test mode is on.")

def create_email(receiver_addresses, filenames, first_name, last_name, cc_list, list_dict, type_message):
    message = MIMEMultipart()
    message['From'] = sender_address
    message['Subject'] = list_dict['turf_name'] + " PDF"
    list_number = " - ".join(list_dict['list_number'].split("-"))
    doors = list_dict['door_count']
    people = list_dict['person_count']
    date = list_dict['date_generated']
    date_1 = datetime.datetime.strptime(date, "%m/%d/%y")
    dt = date_1 + datetime.timedelta(days=30)
    end_date = '{0}/{1}/{2:02}'.format(dt.month, dt.day, dt.year % 100)
    body = email_body.format(first_name, list_number, end_date, doors, people, type_message)
    message['To'] = ",".join(receiver_addresses)
    if(len(cc_list) > 0):
        message['Cc'] = ",".join(cc_list)
    message.attach(MIMEText(body, 'plain'))
    numAttachedFiles = attachpdfs(filenames, message)
    if(numAttachedFiles == len(filenames)):
        return message
    else:
        print("Failed to create email for " + first_name + " " + last_name + "." + "  Email will not be sent.")
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


def attachpdfs(file_names, email):
    numFiles = 0
    for file_to_attach in file_names:
        expectedFileName = file_to_attach + ".pdf"
        fileName = file_to_attach + "*" + ".pdf"
        fileName = fileName.replace(" ", "")
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

#iterate through folder_dict and create a subfolder copying the files over for each organizer
def create_folders(folder_dict, parent_folder_name):
    parent_path = os.getcwd()
    if(os.path.isdir(parent_folder_name)):
        shutil.rmtree(parent_folder_name)
    os.mkdir(parent_folder_name)
    os.chdir(parent_folder_name)
    for subfolder in folder_dict:
        os.mkdir(subfolder)
        os.chdir(subfolder)
        for file in folder_dict[subfolder]:
            search_file = file + "*" + ".pdf"
            search_file = search_file.replace(" ", "")
            for file in os.listdir(parent_path+"\io\output"):
                found_file = file.replace(" ", "")
                print(search_file)
                print(found_file)
                print()
                if fnmatch.fnmatch(found_file, search_file):
                    shutil.copy(parent_path+"\io\output\\"+file, file)
                    break
        os.chdir("..")
    os.chdir(parent_path)

    



def send_files():
    #Add everybody to the CC list
    dev_cc_list = ["gboicheff@gmail.com", "mjturtora@gmail.com"]
    #dev_cc_list = ["gboicheff@gmail.com"]
    print("==================================================")
    turfs = get_entries()
    list_dict = extract_list_info()
    session = initialize_session()
    organizerFiles = {}
    success = True
    sent_count = 0
    sent_list = []
    for turf in turfs:
        print("-------------------------------------------")
        first_name = turf['first_name']
        last_name = turf['last_name']
        turf_name = turf['turf_name']
        building_name = turf['building_name']
        receiver_address = turf['email_address']
        bc_email = turf['bc_email_address']
        organizer_email = turf['organizer_email_address']
        type_message = turf['message']
        final_cc_list = dev_cc_list + [bc_email, organizer_email]
        turf_name_s = turf_name.split()
        print(turf_name)
        filename = turf_name + " " + first_name + " VBM"
        if organizer_email in organizerFiles:
            organizerFiles[organizer_email] += [filename]
        else:
            organizerFiles[organizer_email] = [filename]  
        print("Send email to " + first_name + " " + last_name + " at " + receiver_address)
        print("CCing: " + str(final_cc_list))
        print("Expected filename: " + filename)
        foundFile = find_file(filename, True)
        # print(list_dict)
        print("Found filename: " + find_file(filename, True))
        if input_choice():
            email = create_email([receiver_address], [filename], first_name, last_name, final_cc_list, list_dict[filename], type_message)
            if not testMode:
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
    if not testMode:
        session.quit()
    if success:
        print("All " + str(sent_count) + "emails sent!")
    else:
        print("At least one email not sent!")
    print("==================================================")
    for entry in sent_list:
        print(entry)
    create_folders(organizerFiles, "Organizers")
    

if __name__ == '__main__':
    send_files()
    #print(extract_list_info())
