#################################################################################
#   ftpConnection
#   
#   - Class file that connects to the ftp server and uploads all the contents
#	  of a folder
#
#################################################################################
import ftplib
import os
import getpass
from ftpTracker import FtpUploadTracker

class ftpConnector:

	def __init__(self,pathName):
		print("Connecting to FTP...")
		self.ftp = ftplib.FTP_TLS("ftp.hydromaxusa.com")
		
		#Get the username and password for the FTP connection 
		#Txt file is located on the desktop to ensure user doesnt misplace file
		filename = os.path.join('C:\\Users\\'+getpass.getuser()+'\\Desktop\\ftpLoginInfo.txt')
		fileinfo = open(filename,'r')
		loginInfo = fileinfo.readline().split(' ')
		fileinfo.close()
		username = loginInfo[0]
		password = loginInfo[1]
		self.ftp.login(username, password)
		print("Connected!")
		self.ftp.cwd(pathName)
		print("-"*75)
		print("Currently working in: "+self.ftp.pwd())

	# uploadAllthis
	#
	# Search through path given, if its a folder create a folder
	# on the ftp and check contents inside. If its a file create
	# a binary file and store it
	#
	def uploadAllThis(self,path):
		files = os.listdir(path)
		os.chdir(path)

		for f in files:
			if os.path.isfile(f):
				fileSize = os.path.getsize(f)
				uploadTracker = FtpUploadTracker(int(fileSize))
				print "    uploading file [Filesize= %s bytes]: "%fileSize, f
				binaryFile = open(f, 'rb')
				self.ftp.storbinary('STOR %s' % f, binaryFile,1024,uploadTracker.handle)
				binaryFile.close()
			elif os.path.isdir(f):
				print("Found folder: "+f)
				if f not in self.ftp.nlst():
					print("   creating folder: "+f)
					self.ftp.mkd(f)
					self.ftp.cwd(f)
					self.uploadAllThis(f)
		self.ftp.cwd('..')
		os.chdir('..')
	#
	# Send a message to quit ftp
	#
	def quitFTP(self):
		self.ftp.close()
		print("-"*75)
		print("All done!")

	

