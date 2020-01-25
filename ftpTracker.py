########################################################################
# ftpTracker(class:FtpUploadTracker)
#
#   - Class I was able to find online to help me track the ftp
#     progress when uploading files to the ftp
########################################################################
from __future__ import print_function
import sys

class FtpUploadTracker:
    sizeWritten = 0
    totalSize = 0
    lastShownPercent = 0

    def __init__(self, totalSize):
        self.totalSize = totalSize

    def handle(self, block):
        self.sizeWritten += 1024
        percentComplete = round((float(self.sizeWritten) / float(self.totalSize)) * 100)
    
        if (self.lastShownPercent != percentComplete):
            self.lastShownPercent = percentComplete
            print("    "+str(percentComplete) + "% complete",end='\r')
            sys.stdout.flush()