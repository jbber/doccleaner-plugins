# -*- coding: utf-8 -*-
#Proof of concept. Works quite well from command line, but needs a lot of improvements.

#This script doesn't use the uno Python package, because :
# 1/ Pyuno seems to be usable only with the LibreOffice's embedded Python (can't be installed on a "system Python").
# 2/ This script needs to use the LXML binary package, which can't be easily installed with the LO's Python

#TODO: Need to improve the usability: how to call this script from a button in LibreOffice?

#Imports:
#Packages for using the Windows COM api. May need a cleanup.
import win32com
win32com.__path__
from win32com import universal
from win32com.server.exception import COMException
from win32com.client import gencache, DispatchWithEvents
import winerror
import pythoncom
from win32com.client import constants, Dispatch
import win32com.client

#Packages for dealing with command line parameters, path names, temporary files, etc.
import sys, getopt
import os
import tempfile
import shutil

#Packages for managing the file URL system of LibreOffice
import urllib.request
import urllib.parse

#Package for the xsl processing
from doccleaner import doccleaner


def getTrueURL(docFalseURL):
    decomposedURL = urllib.parse.urlparse( docFalseURL, 'file:///' )
    docTrueURL = decomposedURL[2]
    return docTrueURL

def main(argv):
    #Defining parameters
    try:
        opts, args = getopt.getopt(argv, "t:", ["transformationSheet="])
    except:
        sys.exit(2)

    xsl = None
    for opt, arg in opts:
        if opt in ("-t", "--transformationSheet"):
            xsl = arg
    #Connecting to LibreOffice
    lo = win32com.client.Dispatch("com.sun.star.ServiceManager")
    objDesktop= lo.createInstance("com.sun.star.frame.Desktop")

    #Get the active document
    objDocument= objDesktop.getCurrentComponent()
    #Todo: check if the active document is saved

    #Retrieving the active document URL, and converting it into a proper system path
    originDocURL = getTrueURL(objDocument.getURL())
    decomposedURL = urllib.parse.urlparse( originDocURL, 'file:///' )
    originDocURL = decomposedURL[2]
    originDocPath = urllib.request.url2pathname(originDocURL)

    #Retrieving the document name
    originDocName = os.path.split(originDocPath)[1]
    transitionalDocPath = originDocPath #Will be useful later for consecutive processings

    #TODO: consecutive processings for each xsl parameter defined in a localized ini file
    #defining the path to the temp output document
    tmp_dir = tempfile.mkdtemp()
    newDocPath = os.path.join(tmp_dir, "new_"+originDocName)
    newDocURL = getTrueURL(urllib.request.pathname2url(newDocPath)) #REMOVE ? urllib.parse.urljoin("file:", urllib.request.pathname2url(newDocPath))

    subFileArg = "" #TODO
    XSLparameter = "" #TODO

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
    #objDesktop.loadComponentFromURL(objDocument.getURL(), "_default", 0, ()) #activating the original document

    #TODO : copy the content of the new doc to the original doc. Maybe using insertDocumentFromURL?
    #in the meantime, we settle for opening the processed document in LO Writer:
    objDocument2 = objDesktop.loadComponentFromURL("file://"+newDocURL, "_blank", 0, ())

#    TODO: removing the temporary folder
#    objDocument2.close()
#    try:
#         shutil.rmtree(folder)
#         print(folder + " deleted")
#     except:
#         pass

if __name__ == '__main__':
    #Argument to pass: filename of the xsl to apply, e.g. "cleanDirectFormatting.xsl"
    main(sys.argv[1:])