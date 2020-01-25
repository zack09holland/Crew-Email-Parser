########################################################################################################################
#	findOutlookFolders
#
#	-Another method that was used to test out a different way of going through Microsoft Outlook to obtain the
#	 the needed folder to perform the task I was doing with checkEmails.py
#	-Goes through the folder inboxs for the given Outlook application user and finds the one youre looking for and
#	displays it
#
########################################################################################################################

import win32com.client
#Create an object for the outlook application 
outlook = win32com.client.Dispatch("Outlook.Application")
#Create an object to gain access to the Microsoft API
mapi = outlook.GetNamespace('MAPI')
inbox =  mapi.GetDefaultFolder(win32com.client.constants.olFolderInbox)

fldr_iterator = inbox.Folders   
desired_folder = None
# Go through all the folders found in Outlook and search for Proximity Emails
while 1:
    f = fldr_iterator.GetNext()
    print (f.Name)
    if not f: break
    if f.Name == 'Proximity Emails':
        print ('-' * 50)
        print ('found "Proximity Emails" dir')
        desired_folder = f
        break

print (desired_folder.Name)