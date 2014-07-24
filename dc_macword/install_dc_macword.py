#coding: utf-8 -*-
#A script to generate applescripts calling doccleaner from the MacWord Script menu
#Not working yet!

import os
import doccleaner
#TODO: localization of datadir (or getting it dynamically, but how?):
#fr = "Données\ utilisateurs\ Microsoft"
#en = "Microsoft\ Users\ Data"
#es = ?
#pt = ?
#de = ?
#ar = ?

datadir = r"Données\ utilisateurs\ Microsoft"
SCRIPTS_PATH = os.path.join(r"~", r"Documents", datadir, r"Word\ Script\ Menu\ Items")

for path, subdirs, files in os.walk(os.path.join(os.path.dirname(doccleaner.__file__), 'docx')):
     for filename in files:
        if filename.endswith(".xsl"):
                        #Writing the applescript to the "~/Documents/Données\ utilisateurs\ Microsoft/Word\ Script\ Menu\ Items" folder
            f = open(os.path.join(SCRIPTS_PATH, title), "w")
            f.write(generateApplescript(filename))
            f.close


#TODO: adding handling of the parameters --subfile and --xslparameters
def generateApplescript(xslpath):
    #generating the template for an applescript using doccleaner
    applescript = r"""

    tell application "Microsoft Word"
        -- retrieving the path of the current document in word
        set docPath to the full name of the active document
        -- converting it to POSIX form, in order to manipulate it
        set docPath to POSIX path of docPath

        -- retrieving the doc title
        set docName to the name of the active document

    end tell

    -- defining a temporary folder
    set tempDir to path to temporary items for user domain
    -- converting the tempDir path to POSIX form, in order to manipulate it
    set tempDir to POSIX path of tempDir

    -- generating a path to a temporary file
    set tempFile to quoted form of tempDir & "~" & docName

    -- retrieving path of doccleaner
    set py to "doccleaner.py "
    workingDir = """ + os.path.dirname(doccleaner.__file__) +  """

    set callDir to quoted form of workingDir & py

    -- launching doccleaner
    do shell script callDir -i docPath -o tempFile -t """ + xslpath +  """

    -- TODO : copying content from tempFile to original doc
    tell application "Microsoft Word"

    end tell

    -- removing the tempFile
    tell application "Finder"
        delete file tempFile
        -- empty trash ?
    end tell
    """
    return applescript
