from PyPDF2 import PdfFileWriter, PdfFileReader
import re
import datetime
import csv
import os
import sys
import win32com.client as win32
import win32print
import win32api
import glob
printPdf = PdfFileWriter()

class EmailMessage:
    def __init__(self, name, email_recipient,
                email_subject,
                email_message,
                attachment_location,page,encrypt):
        
        self.name = name
        self.email_recipient = email_recipient
        self.email_subject = email_subject
        self.email_message = email_message
        self.attachment_location = attachment_location
        self.page = page
        self.encrypt = encrypt

    def print(self):
        print(self.name+" - "+self.email_recipient)
        
#Sends email with attachment
def send_email(name, email_recipient,
                email_subject,
                email_message,
                attachment_location,page,encrypt):

    if os.path.exists(os.getcwd()+'\\'+attachment_location):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = email_recipient
        mail.Subject = email_subject[:-4].replace('-'," ")
        mail.Body = email_message
        mail.Attachments.Add(os.getcwd()+'\\'+attachment_location)
        mail.Send()

def sendEmails(emaillist, logfile):

    for i in emaillist:
        print("Sending : "+i.name)
        send_email(i.name,i.email_recipient,i.email_subject,i.email_message,i.attachment_location,i.page,i.encrypt)
        write_log(name,"Email Sent to "+i.email_recipient,logfile)

def write_log(name,message,filename):

    f = open(filename,"a+")
    f.write(name+";"+message+";"+str(datetime.datetime.now().timestamp())+"\n")

def print_pdf(page):
    printPdf.addPage(page)
    with open("temp/PrintFile.pdf", "wb") as outputStream:
        printPdf.write(outputStream)

def startup(logfile):
    if not os.path.isdir("temp"):
        os.mkdir("temp")
    if not os.path.isdir("log"):
        os.mkdir("log")
    f = open(logfile,"w+")
    f.close()

    if os.path.exists("temp/PrintFile.pdf"):
        os.remove(os.getcwd()+'\\'+"temp/PrintFile.pdf")    

    today = datetime.date.today()   
    first = today.replace(day=1)
    lastMonth = first - datetime.timedelta(days=1)
    date = lastMonth.strftime("-%b %Y")
    return date

def get_latest_file(dirName):
    # create a list of file and sub directories 
    # names in the given directory 
    listOfFile = os.listdir(dirName)
    allFiles = list()
    newFile = list()
    # Iterate over all the entries
    for entry in listOfFile:
        # Create full path
        fullPath = os.path.join(dirName, entry)
        # If entry is a directory then get the list of files in this directory 
        if os.path.isdir(fullPath):
            allFiles = allFiles + get_latest_file(fullPath)
        else:
            allFiles.append(fullPath)

    for i in allFiles:
        if ".pdf" not in i:
            allFiles.remove(i)
        else:               
            if len(newFile) == 0:
                newFile.append(i)
            elif os.path.getctime(i) > os.path.getctime(newFile[0]):
                newFile[0] = i
   
    return newFile

####################################################################
names = []
emails = []
passwords = []
firstpage = False
logfile = "log/Logfile-"+str(datetime.datetime.now().timestamp())+".txt"
emaillist = []

date = startup(logfile)
with open('emails.csv', 'r') as file:
    reader = csv.reader(file, delimiter = ';')
    for row in reader:
        names.append(row[0])
        emails.append(row[1])
        passwords.append(row[2])

if len(sys.argv) != 1 and os.path.exists(sys.argv[1]):
    pdff = open(sys.argv[1], "rb")
    inputpdf = PdfFileReader(pdff)
elif len(sys.argv) > 1:
    print("Too Many Arguments\nQuitting")
    exit()
else:
    print("Finding Latest PDF")
    filename = get_latest_file("S:\\")
    print(filename[0])
    yesno = input("Do you wish to use this file? (y or n): ")
    yesno = "n"
    if yesno == "y" or yesno == "Y":
        pdff = open(filename[0], "rb")
        inputpdf = PdfFileReader(pdff)
    elif yesno == "n" or yesno == "N":
        while True:
            filename = input("Please enter path to file: ")
            filename = filename.replace("\"","")
            if os.path.exists(filename):
                pdff = open(filename, "rb")
                inputpdf = PdfFileReader(pdff)
                break      
        
for i in range(int(inputpdf.numPages/2)):
    if firstpage is True:      
        page = inputpdf.getPage(i)
        output = page.extractText()
        output = output.replace("\n","")
        NameStart = output.rfind("N*Pers.-Nr.")+20

        if output.rfind("*JFI") != -1:
            NameStart = output.rfind("*JFI")+4
           
        # Debugging print(output) 
        output = output[NameStart:NameStart+30]
        name = re.findall('[A-Z][a-z]*', output)
        name = name[0] + " " + name[1]
        
        message = 'Hallo,\n\nAnbei die Lohnabrechnung von '+date[1:]+'.\nHinweis: Das PDF ist mit Deinem Geburtsdatum(DDMM) verschlüsselt\n\nLiebe Grüße\nManagement'
        recip = "y"
        
        filename = "temp/Lohnabrechnung-"+name+date+".pdf"
        filename = filename.replace(" ", "")
        if os.path.exists(filename):
            filename = "temp/Lohnabrechnung-"+name+date+"-1.pdf"
            filename = filename.replace(" ", "")
        
        if name in names:
            recip = emails[names.index(name)]
            encrypt = passwords[names.index(name)]
            if recip == "print":
                print_pdf(inputpdf.getPage(i))
                write_log(name,"Printed",logfile)
            else:
                tempPdf = PdfFileWriter()
                tempPdf.addPage(inputpdf.getPage(i))
                tempPdf.encrypt(encrypt)
                with open(filename, "wb") as outputStream:
                    tempPdf.write(outputStream)
                emaillist.append(EmailMessage(name,recip,filename[5:],message,filename,inputpdf.getPage(i),encrypt))
        else:
            while not re.match(r"[^@]+@[^@]+\.[^@]+", recip): 
                recip = input("Enter Email for(s to skip, p to print, q to quit) "+name+" : ")
                if recip == "s":
                    write_log(name,"Skipped",logfile)
                    break
                elif recip == "q":
                    print("Quitting")
                    write_log("Program ended prematurely upon request","",logfile)
                    exit()
                elif recip == "p":
                    print_pdf(inputpdf.getPage(i))
                    oneTime = input("Always Print(y) or One time(n)? ")
                    if oneTime == "y" or oneTime == "Y":
                        app = open("emails.csv","a+")
                        app.write("\n"+name+";print;0000")
                        write_log(name,"Printed",logfile)                        
                    break
                elif re.match(r"[^@]+@[^@]+\.[^@]+", recip):
                    app = open("emails.csv","a+")
                    encrypt = input("Enter DDMM password: ")
                    app.write("\n"+name+";"+recip+";"+encrypt)
                    app.close()
                    emaillist.append(EmailMessage(name,recip,filename[5:],message,filename,inputpdf.getPage(i),encrypt))
                    write_log(name,"Email Sent to "+recip,logfile)
                    break
                else:
                    print("Please enter a valid Email")
                
    else:
        firstpage = True

for i in emaillist:
    i.print()

yesno = input("Sends emails now (Y) or (N): ")
if yesno == "y" or yesno == "Y":
    sendEmails(emaillist, logfile)
            
pdff.close()
##TODO Clean up Temp folder

if os.path.exists(os.getcwd()+'\\'+"temp/PrintFile.pdf"):
    os.startfile(os.getcwd()+'\\'+"temp/PrintFile.pdf", "open")
print("Finished")
os.system('pause')


