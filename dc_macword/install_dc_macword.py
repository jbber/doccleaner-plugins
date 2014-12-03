# -*- coding: utf-8 -*-
"""
Created on Mon Dec 01 11:18:57 2014

@author: Bertrand
"""

#coding: utf-8 -*-
#A script to generate applescripts calling doccleaner from the MacWord Script menu
#Not working yet!

import os
import doccleaner
from string import Template
import simplejson

#TODO: localization of datadir (or getting it dynamically, but how?):
#fr = "Données\ utilisateurs\ Microsoft"
#en = "Microsoft\ Users\ Data"
#es = ?
#pt = ?
#de = ?
#ar = ?

datadir = r"Données\ utilisateurs\ Microsoft"
SCRIPTS_PATH = os.path.join(r"~", r"Documents", datadir, r"Word\ Script\ Menu\ Items")
def load_json(filename):
    f = open(filename, "r")
    data = f.read()
    f.close()
    return simplejson.loads(data)

jsonConf = load_json(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'addin.json'))
path_dict = {}
jj = 0
for button in jsonConf["buttons"]:

    DOCCLEANER_PATH = '"' + os.path.dirname(doccleaner.__file__) + '"'

    #path_dict['DOCCLEANER_PATH'] = DOCCLEANER_PATH
    XSLNUMBER = len(button["xsl"])
    for xsl in button["xsl"]:

        
        with open("template.applescript", "r") as f:
            contents = f.read()
        
        xslid = str(xsl["xslid"])
                                                       
        path_dict[xslid] = {'XSL_PATH': str( os.path.join(os.path.dirname(__file__), "docx", xsl["XSLname"]) ), 
                            'SUBFILE': str( os.path.join(os.path.dirname(__file__), "docx", xsl["subfile"]) ),
                            'XSLPARAMETER':str( xsl["XSLparameter"] ),
                            'DOCCLEANER_PATH': DOCCLEANER_PATH,
                            'XSLNUMBER': XSLNUMBER
                            }
    
    
    path_dict[xslid]["PROCESSINGS"]= str(path_dict)

    #Telling Python that the contents of "template.applescript" are a template:
    tpl = Template(contents)
    
    #passing path_dict to the template
    contents = tpl.substitute(path_dict[xslid])

    #escaping the quotes
    contents = contents.replace('"', '\"')
                
    #Generating a list of commands for osacompile
    command_list = (contents.expandtabs()).split("\n")
    
    #adding "-e" parameter at the start of each list element. Don't add parameter if line is empty
    command_list = [' -e "{0}"'.format(command) for command in command_list if command]
    
    #appending a -o parameter to the list, containing the path to the script we want to compile
    processing_name = os.path.splitext(filename)[0] #TODO: giving a friendlier, localized name to this variable
    command_list.append(' -o '+ os.path.join(SCRIPTS_PATH,  processing_name +".scpt") )

    #launching the compilation with osacompile 
    os.system("osacompile " + ''.join(command_list)) #TODO: test...
         


