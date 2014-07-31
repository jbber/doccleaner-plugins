# -*- coding: utf-8 -*-
#Proof of concept. Works quite well from command line, but needs a lot of improvements.

#This script doesn't use the uno Python package, because :
# 1/ Pyuno seems to be usable only with the LibreOffice's embedded Python (can't be installed on a "system Python").
# 2/ This script needs to use the LXML binary package, which can't be easily installed with the LO's Python

#Imports:
#Packages for using the Windows COM api. 
import win32com.client
#Packages for dealing with command line parameters, path names, temporary files, etc.
import sys, getopt
import os
import tempfile
import shutil
#Package for working with ini files
import configparser
#Packages for managing the file URL system of LibreOffice
import urllib.request
import urllib.parse
#Package for the xsl processing
from doccleaner import doccleaner

def documentController(self):
    #Function for selecting the whole document and put it in a controller
    documentText = self.getText()
    cursor = documentText.createTextCursor(  )
    cursor.collapseToStart()
    cursor.goToEnd(True)
    controller = self.CurrentController
    controller.select(cursor)
    return controller

def getTrueURL(docFalseURL):
    #A function to handle the LibreOffice's file URL system
    decomposedURL = urllib.parse.urlparse( docFalseURL, 'file:///' )
    docTrueURL = decomposedURL[2]
    return docTrueURL

def main(argv):
    #Defining parameters
    #"-t" : the xsl we want to use to process the current document
    #"-b" : the button ID in the (upcoming) LibreOffice menu or toolbra
    try:
        opts, args = getopt.getopt(argv,
                                   "t:b:", ["transformationSheet=","buttonID="])
    except:
        sys.exit(2)

    xsl = None
    buttonID = None
    for opt, arg in opts:
        if opt in ("-t", "--transformationSheet"):
            xsl = arg
         if opt in ("-b", "--buttonID"):
             buttonID = arg

    config = configparser.ConfigParser()
    config.read(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'libreoffice_win_fr.ini')) #TODO : defining the file more dynamically, for better localization handling

    #Connecting to LibreOffice
    lo = win32com.client.Dispatch("com.sun.star.ServiceManager")
    objDesktop= lo.createInstance("com.sun.star.frame.Desktop")

    #Get the active document
    objDocument= objDesktop.getCurrentComponent()
    objDocumentController = documentController(objDocument)

    #TODO: check if the active document is saved (not a new unsaved doc)

    #Retrieving the active document URL, and converting it into a proper system path
    originDocURL = getTrueURL(objDocument.getURL())
    originDocPath = urllib.request.url2pathname(originDocURL)

    #Retrieving the document name
    originDocName = os.path.split(originDocPath)[1]
    transitionalDocPath = originDocPath #Will be useful later for consecutive processings

    #defining the path to the temp output document
    tmp_dir = tempfile.mkdtemp()
    newDocPath = os.path.join(tmp_dir, "~" + originDocName)
    newDocURL = getTrueURL(urllib.request.pathname2url(newDocPath)) 
    jj = 0 #the jj variable is used below, to make consecutive XSL processings, if defined in the ini file
    
    #Retrieving the XSL parameters defined in the ini file
    try:
        XSLparameters = config.get(str(buttonID), 'XSLparameter').split(";")
    except:
        XSLparameters = ""
    if XSLparameters != None:

        for XSLparameter in XSLparameters:
            #Check if there are subfiles to process consecutively instead of simulteanously (separated by a semi-colon instead of a comma)
            #NB : the script implies that in the ini file, we can have:
            # 1) one XSL parameter, and a single subfiles processing
            # 2) multiple XSL parameters, and the exact same number of consecutive subfiles processings
            # 3) multiple XSL parameters, and a single subfiles processing
            #We can never have multiple subfiles and a single XSL processing, because this use case is handled separately by the docCleaner script. If we're in this case, simply split subfiles with commas (",") instead of semi-colon (";")
            try:
                subFileArg = str(config.get(str(buttonID), 'subfile')).split(";")[jj]
            except:
                #Probably a "out of range" error, which means there is a single subfiles string to process
                subFileArg = config.get(str(buttonID), 'subfile')

            if jj > 0:
                #If there is more than one XSL parameter, we'll have to make consecutive processings
                newDocName, newDocExtension = os.path.splitext(newDocPath)
                transitionalDoc = newDoc
                newDocPath =  newDocName + str(jj)+ newDocExtension
                transitionalDocPath = newDocPath

            dc_arguments = ['--input', str(transitionalDocPath),
                            '--output', str(newDocPath),
                            '--transform', os.path.join(os.path.dirname(doccleaner.__file__),
                                    "docx", str(xsl) )
                            ]

            if subFileArg != "":
                dc_arguments.extend(('--subfile', subFileArg))

            if XSLparameter != "":
                dc_arguments.extend(('--XSLparameter', XSLparameter))

            #launch the xsl processing
            doccleaner.main(dc_arguments)
            jj +=1

    #Retrieving the contents of the new document
    #TODO: loading the new doc with "hidden" property. Need to use "com.sun.star.beans.PropertyValue", but lo.bridge_getstruct("com.sun.star.beans.PropertyValue") seems to be broken somehow
    newDocproperties = ()
    newObjDocument = objDesktop.loadComponentFromURL("file://"+newDocURL, "_blank", 0, newDocproperties )
    newDocController = documentController(newObjDocument)
    newContent = newDocController.getTransferable()

    #Inserting the new document into the original document
    objDocumentController.insertTransferable(newContent)

    #Closing the new document
    try:
        newObjDocument.close(False)
    except Exception as e:
        print(str(e))

    #Removing the temporary folder
    try:
         shutil.rmtree(tmp_dir)
         print(tmp_dir + " deleted")
    except Exception as e:
         print("Deletion of "+ str(tmp_dir) + " failed!")
         print(str(e))
         pass

if __name__ == '__main__':
    #Argument to pass: filename of the xsl to apply, e.g. "cleanDirectFormatting.xsl"
    main(sys.argv[1:])
