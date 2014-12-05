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
data = ""
jj = 0
for button in jsonConf["buttons"]:

    path_dict["DOCCLEANER_PATH"] = '"' + os.path.dirname(doccleaner.__file__) + '"'
    path_dict["XSLNUMBER"] = len(button["xsl"])

    data = ""    
    #data = ",".join(xsl for xsl in button["xsl"])
    print data
    data = ",".join(str(""" |{0}|:{{XSL_PATH:{1}, SUBFILE:{2}, XSLPARAMETER:{3} }}""".format(
                                                                                             xsl["xslid"],
                                                                                             '"'+str( os.path.join(os.path.dirname(__file__), "docx", xsl["XSLname"]) )+'"',
                                                                                             '"'+str( os.path.join(os.path.dirname(__file__), "docx", xsl["subfile"]) )+'"',
                                                                                             '"'+str( xsl["XSLparameter"] )+'"',                                                                                                                                                              
                                                                                            ) ) for xsl in button["xsl"])

    with open("template.applescript", "r") as f:
        contents = f.read()
        
                                                                                                                            
    path_dict["PROCESSINGS"] = data
        
    #Telling Python that the contents of "template.applescript" are a template:
    tpl = Template(contents)
    
    #passing path_dict to the template
    contents = tpl.substitute(path_dict)

    #escaping the quotes
    #contents = contents.replace('"', '\"')
                
    #Generating a list of commands for osacompile
    command_list = (contents.expandtabs()).split("\n")
    
    #adding "-e" parameter at the start of each list element. Don't add parameter if line is empty
    command_list = [' -e "{0}"'.format(command) for command in command_list if command]
    
    #appending a -o parameter to the list, containing the path to the script we want to compile
    processing_name = str(button["tag"])
    command_list.append(' -o '+ os.path.join(SCRIPTS_PATH,  processing_name +".scpt") )
    jj+=1
    #launching the compilation with osacompile 
    os.system("osacompile " + ''.join(command_list)) #TODO: test...

