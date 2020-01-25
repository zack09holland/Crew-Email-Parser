#################################################################################
#   FileWrapper
#   
#   - Class file with various methods to write to file, create folder, etc.
#
#################################################################################
import os
import sys
from PyRTF import *
class FileWriter(object):
    '''
    convenient file wrapper for writing to files
    '''
    def __init__(self, path, filename):
        '''
        Constructor
        '''
        self.doc = Document()
        self.section = Section()
        self.doc.Sections.append( self.section )

        self.tally = 0  
        self.file = open(path+'\%s.rtf' % filename, "w")
    # Place string into file with 'UTF-8' encoding  
    def pl(self, a_string):
        str_uni = a_string.encode('utf-8')
        self.file.write(str_uni)
        self.file.write("\n")
    # Write to a file without using the unicode 
    #   -check if the list is greater than 1 and
    #     output the value with a '|' divider
    #   -pn = parcel number   
    def writeList(self,pn_list):
        if len(pn_list) >= 1:
            for pn in pn_list: self.file.write(pn+'|')
            self.file.write("\n")
            self.tally += 1
    # Create a folder for the parcel number 
    def createFldr(self,path,fldr_name):
        os.mkdir(path+"\\"+fldr_name)

    def flush(self):
        self.file.flush()