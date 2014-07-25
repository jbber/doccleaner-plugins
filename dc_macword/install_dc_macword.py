#coding: utf-8 -*-
#A script to generate applescripts calling doccleaner from the MacWord Script menu
#Not working yet!

import os
import doccleaner
from string import Template
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
            with open("template.applescript", "r") as f:
                contents = f.read()

            path_dict = {'DOCCLEANER_PATH': '"' + os.path.dirname(doccleaner.__file__) + '"',
                         'XSL_PATH':        '"' + os.path.join(path,filename) + '"'}
            
            #Telling Python that the contents of "template.applescript" are a template:
            tpl = Template(contents)
            
            #passing path_dict to the template
            contents = tpl.substitute(path_dict)
            
            #escaping the quotes
            contents = contents.replace('"', '\"')
            
            #Generating a list of commands for osacompile
            command_list = (contents.expandtabs()).split("\n")
            
            #adding "-e" parameter at the start of each list element
            command_list = ['-e "{0}"'.format(command) for command in command_list]
            
            #appending a -o parameter to the list, containing the path to the script we want to compile
            processing_name = os.path.splitext(filename)[0] #TODO: giving a friendlier, localized name to this variable
            command_list.append('-o '+ os.path.join(SCRIPTS_PATH,  processing_name +".scpt") )
            
            #launching the compilation with osacompile 
            os.system("osacompile " + ''.join(command_list)) #TODO: test...
             


