################################################################################
#   checkEmails
#   ***ONLY WORKS WITH PYTHON 2.X***
#       --venvRequirements.txt are the necessary packages to run program
#   ***PROGRAM DISCONTINUED DUE TO NEW MOBILE APPLICATION THAT CREWS ARE USING NOW***
#
#   --> Proximity team members will have a folder called 'Proximity Emails'
#       --A microsoft outlook rule automatically moves crew emails to this folder
#   --> This program will search this Outlook folder for crew emails that have a
#       specific subject line the crews are told to write as, so the program will
#       find their emails
#   --> For each email found
#           -a text file is created containing the contents of the 
#           message and saved into its own folder on the desktop.
#           -attached photos are saved into the same folder
#   --> After creating the folders with the content of the email, it is then
#       uploaded to the FTP server
#   --> The newly created Proximity Emails folder on the desktop is then linked
#       to each corresponding parcel number within ArcGIS for easier access to
#       these newly created files
#
#################################################################################
import win32com.client
import os, time, re, sys, getpass, socket
from FileWrapper import FileWriter
from ftpConnection import ftpConnector
class CheckMailer:

    # Initialize the checker, and file writter to show the contents
    # - Normal Inbox index is 6
    def __init__(self):
        self.parcelnumPath = None
        self.outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
        #self.inbox = self.outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)
        self.inbox = self.outlook.GetDefaultFolder(6)
    # This method checks the outlook folders for 'Proximity Emails'
    # - Once it finds it, we save it
    def check(self):
        # This section of code gets all the outlook folders and
        # iterates through them to find the 'Proximity Emails'
        # folder.
        fldr_iterator = self.inbox.Folders
        desired_folder = None
        self.proxpath = None
        print('Searching Outlook for Proximity Emails''...')
        while 1:
            nextFldr = fldr_iterator.GetNext()
            if not nextFldr: break  # Quick way to iterate through the folders
            if nextFldr.Name == 'Proximity Emails':
                print('Found "Proximity Emails" in Outlook.')
                desired_folder = nextFldr
                # Make a pathway for the user on the desktop and create a folder
                # to contain the proximity emails
                self.proxpath = os.path.join('C:\\Users\\'+getpass.getuser()+'\\Desktop',desired_folder.Name)
                if not os.path.exists(self.proxpath):
                    os.mkdir(self.proxpath)
                break
        # If Proximity Emails is not found, exit the program
        if desired_folder == None:
            print(' ***''Proximity Emails'' was not found in Outlook.***\n ***Double check Outlook to make sure its spelled the same.***')
            sys.exit(0)
        # This gets the most recent message
        messages = desired_folder.Items
        message = messages.GetLast()       # <------Change this to GetLast
        # This pattern is used to find the 10 or 15 digit parcel numbers
        pattern = re.compile(r'\d{10,15}')
        totMess = 0
        # Create text file to hold the all the parcel numbers found                
        parcelNumOutput = FileWriter(self.proxpath,'Parcel Numbers.txt')
        parcelNumOutput.pl("Parcel Numbers")
        parcelNumOutput.pl("-" * 70)
        # While there is a message in the folder
        #   - Find all the parcel numbers in the subject line
        #   - Output a list of all numbers found to .txt file
        print('Finding crew emails...')
        while message:
            #Create list for all parcel numbers found in the subject line
            parcelNumber = pattern.findall(message.Subject)
            parcelNumOutput.writeList(parcelNumber)
            # For each parcel number contained in the message
            # create a directory for it
            for pn in parcelNumber:
                if not os.path.exists(self.proxpath+"\\"+pn):
                    os.mkdir(self.proxpath+"\\"+pn)
                    self.parcelnumPath = self.proxpath+"\\"+pn
                    print ("Folder created: "+self.parcelnumPath)
                    # Create file object for each parcel number
                    # containing the subject line and body of 
                    # the email
                    emailDoc = FileWriter(self.parcelnumPath,pn)

                    emailDoc.pl("-"*70+"\n"+"Subject line\n"+"-"*70)
                    emailDoc.pl(" "+message.Subject+"\n")
                    emailDoc.pl("-"*70)
                    emailDoc.pl("Body text of Message(s)\n"+"-"*70+"\n")
                    emailDoc.pl(message.Body)
                    emailDoc.pl("-"*70+"\n")
                    # Get all the attachments contained within the message
                    #   -create path string for the picture and save it
                    for att in message.Attachments:
                        # I noticed that the hydromax signature attachments
                        # are saved like this, which we will skip when processing
                        if att.FileName == 'image001.png':
                            break
                        else:
                            picPath = os.path.join(self.parcelnumPath+"\\"+ att.FileName)
                            att.SaveAsFile(picPath)
                    message = messages.GetPrevious()
                    totMess += 1
            # If the message already exists go to the next message
            # and start the iteration over again
            else:
                message = messages.GetPrevious()
                totMess += 1
        print ("-" * 75)
        print ("Total number of parcels found: %i" % parcelNumOutput.tally)
        print ("Total number of messages: %i" % totMess) 
        print ("-" * 75)

    # After parsing through all the emails. Upload them to the folder
    # on the ftp server
    def uploadToFTP(self,pathName):
        try:
            proxEmailFtpPath = "Data Upload\Proximity Team\Proximity_Emails"
            ftpSession = ftpConnector(proxEmailFtpPath)
            print("Uploading email folders to the FTP..")
            ftpSession.uploadAllThis(pathName)
            ftpSession.quit()
        except socket.error:
            print("The connection timed out. Trying again..")
            mail.uploadToFTP(mail.proxpath)

if __name__ == "__main__":
    mail = CheckMailer()
    mail.check()
    while 1:
        print("*Program may need to be re-run to get all messages")
        status = raw_input("Are we ready to upload? Enter yes to upload \n"
                    +"                        Enter no to check for emails again\n"
                    +"                        Enter quit to end the program\n:")
        if status == 'yes':
            mail.uploadToFTP(mail.proxpath)
            break
        elif status == 'no':
           print("Re-running program..")
           mail.check()
        elif status == 'quit':
            break
        else:
           print("Incorrect command. Try that again..")


    


    