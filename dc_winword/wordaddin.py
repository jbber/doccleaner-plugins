#coding: utf-8 -*-

#This is just a proof of concept with probably a lot of bugs, use at your own risks!

#It creates a new tab in the MS Word ribbon, with buttons calling docCleaner scripts from inside Word
#To uninstall it, launch it with the --unregister argument
#You can also remove it from Word (in the "Developer" tab, look for COM Addins > Remove)

#Inspired by the Excel addin provided in the win32com module demos, and the "JJ Word Addin" (I don't remember where I get it, but thanks!)


import win32com
win32com.__path__
from win32com import universal
from win32com.server.exception import COMException
from win32com.client import gencache, DispatchWithEvents
import winerror
import pythoncom
from win32com.client import constants, Dispatch
import sys
import win32com.client

import os
import win32ui
import win32con
import locale
import gettext
import simplejson

import tempfile
import shutil
import mimetypes

from doccleaner import doccleaner
#import doccleaner.localization
#win32com.client.gencache.is_readonly=False
#win32com.client.gencache.GetGeneratePath()
# Support for COM objects we use.
gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 1, bForDemand=True) # Office 9
gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 5, bForDemand=True)
gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 5) #Word

# The TLB defining the interfaces we implement
try:
    universal.RegisterInterfaces('{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}', 0, 1, 0, ["_IDTExtensibility2"])
    universal.RegisterInterfaces('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 5, ["IRibbonExtensibility", "IRibbonControl"])
except:
    pass

locale.setlocale(locale.LC_ALL, '')
user_locale = locale.getlocale()[0]

def checkIfDocx(filepath):
    if mimetypes.guess_type(filepath)[0] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return True
    else:
        return False

#TODO : localization
def init_localization():
    '''prepare l10n'''
    print(locale.setlocale(locale.LC_ALL,""))
    locale.setlocale(locale.LC_ALL, '') # use user's preferred locale

    # take first two characters of country code
    loc = locale.getlocale()

    #filename = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), "lang", "messages_%s.mo" % locale.getlocale()[0][0:2])
    filename = os.path.join("lang", "messages_{0}.mo").format(locale.getlocale()[0][0:2])
    try:
        print("Opening message file {0} for locale {1}".format(filename, loc[0]))
        #If the .mo file is badly generated, this line will return an error message: "LookupError: unknown encoding: CHARSET"
        trans = gettext.GNUTranslations(open(filename, "rb"))

    except IOError:
        print("Locale not found. Using default messages")
        trans = gettext.NullTranslations()

    trans.install()

def load_json(filename):
    f = open(filename, "r")
    data = f.read()
    f.close()
    return simplejson.loads(data)
    
    
class WordAddin:
    

    _com_interfaces_ = ['_IDTExtensibility2', 'IRibbonExtensibility']
    _public_methods_ = ['do', 'apply_style']
    _reg_clsctx_ = pythoncom.CLSCTX_INPROC_SERVER
    _reg_clsid_ = "{C5482ECA-F559-45A0-B078-B2036E6F011A}"
    _reg_progid_ = "Python.DocCleaner.WordAddin"
    _reg_policy_spec_ = "win32com.server.policy.EventHandlerPolicy"

    def __init__(self):
        self.appHostApp = None

    def apply_style(self,ctrl):     
#    #The ctrl argument is a callback for the button the user made an action on (e.g. clicking on it)
#    
#        #Creating a word object inside a wd variable
        wd = win32com.client.Dispatch("Word.Application")
        
        try:
            #Applying style
            wd.Selection.Style = wd.ActiveDocument.Styles(ctrl. Tag)
        except:
            #If style does not exist -> create it, then apply it
            wd.ActiveDocument.Styles.Add(ctrl. Tag)
            wd.Selection.Style = wd.ActiveDocument.Styles(ctrl. Tag)
        
    def do(self,ctrl):
    #This is the core of the Word addin : manipulates docs and calls docCleaner
    #The ctrl argument is a callback for the button the user made an action on (e.g. clicking on it)

        #Creating a word object inside a wd variable
        wd = win32com.client.Dispatch("Word.Application")

        try:
            #Check if the file is not a new one (unsaved)
            if os.path.isfile(wd.ActiveDocument.FullName) == True:
                #Before processing the doc, let's save the user's last modifications
                #TODO: ne fonctionne pas correctement
                wd.ActiveDocument.Save

                originDoc = wd.ActiveDocument.FullName #:Puts the path of the current file in a variable
                tmp_dir = tempfile.mkdtemp() #:Creates a temp folder, which will contain the temp docx files necessary for processing

                #TODO: If the document is in another format than docx, convert it temporarily to docx
                #At the processing's end, we'll have to convert it back to its original format, so we need to store this information

                transitionalDoc = originDoc #:Creates a temp transitional doc, which will be used if we need to make consecutive XSLT processings. #E.g..: original doc -> xslt processing -> transitional doc -> xslt processing -> final doc -> copying to original doc
                newDoc = os.path.join(tmp_dir, "~" + wd.ActiveDocument.Name) #:Creates a temporary file (newDoc), which will be the docCleaner output


                jj = 0 #:This variable will be increased by one for each XSL processing defined in the json file

                #Then, we take the current active document as input, the temp doc as output
                #+ the XSL file passed as argument ("ctrl. Tag" variable, which is a callback for the ribbon button tag)
    
                for button in self.jsonConf["buttons"]:
                    if button["tag"] == ctrl. Tag:
                        for xsl in button["xsl"]:
                            if jj > 0:
                                
                                #If there is more than one XSL sheet, we'll have to make consecutive processings
                                newDocName, newDocExtension = os.path.splitext(newDoc)
                                transitionalDoc = newDoc
                                newDoc =  newDocName + str(jj)+ newDocExtension
        
        
                            dc_arguments = ['--input', str(transitionalDoc),
                                            '--output', str(newDoc),
                                            '--transform', os.path.join(os.path.dirname(doccleaner.__file__),
                                                                        "docx", xsl["XSLname"] ) 
                                                ]
                            
                            for param in ["subfile", "XSLparameter"]:
                                if xsl[param] != 0:
                                    if param == "subfile":
                                        str_param = os.path.join(os.path.dirname(doccleaner.__file__),
                                                                        "docx", xsl[param])     
                                    else:
                                        str_param = xsl[param]
                                    
                                    dc_arguments.extend( ( '--' + param, str_param)) #",".join ( str_param )  )) 
                            
                            doccleaner.main(dc_arguments)                                    
                            jj += 1   
          

                #Opening the temp file
                wd.Documents.Open(newDoc)

                #Copying the temp file content to the original doc
                #To do this, never use the MSO Content.Copy() and Content.Paste() methods, because :
                # 1) It would overwrite important data the user may have copied to the clipboard.
                # 2) Other programs, like antiviruses, may use simulteanously the clipboard, which would result in a big mess for the user.
                #Instead, use the Content.FormattedText function, it's simple, and takes just one line of code:
                wd.Documents(originDoc).Content.FormattedText = wd.Documents(newDoc).Content.FormattedText

                #Closing and removing the temp document
                wd.Documents(newDoc).Close()
                os.remove(newDoc)

                #Saving the changes
                wd.Documents(originDoc).Save
                
                #Removing the whole temp folder
                try:
                    shutil.rmtree(tmp_dir)
                except:
                    #TODO: What kind of error would be possible when removing the temp folder? How to handle it?
                    pass
                wd.ActiveDocument.Save
            else:
                win32ui.MessageBox("You need to save the file before launching this script!"
                ,"Error",win32con.MB_OK)

        except Exception as e:

            tb = sys.exc_info()[2]
            #TODO: writing the error to a log file
            win32ui.MessageBox(str(e) + "\n" +
            str(tb.tb_lineno)+ "\n" +
            str(newDoc)
            ,"Error",win32con.MB_OKCANCEL)

    def GetImage(self,ctrl):
        #TODO : Is this function actually useful?
        #TODO : Retrieving the path from the conf file
        from gdiplus import LoadImage
        i = LoadImage( 'path/to/image.png' )
        return i

    def GetCustomUI(self,control):       
        self.jsonConf = load_json(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'winword_addin.json'))
        xml_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'winword_addin.xml')
        xml_file = open(xml_path, "r")
        
        xml_content = xml_file.read().encode('windows-1252').decode('utf-8') #.encode('windows-1252').decode('utf-8') -> or special characters will behave strangely
        xml_file.close()    

        return xml_content         

    def OnConnection(self, application, connectMode, addin, custom):
        print("OnConnection", application, connectMode, addin, custom)
        try:
            self.appHostApp = application
        except pythoncom.com_error as xxx_todo_changeme:
            (hr, msg, exc, arg) = xxx_todo_changeme.args
            print("The Word call failed with code {0}: {1}".format(str(hr), msg))
            if exc is None:
                print("There is no extended error information")
            else:
                wcode, source, text, helpFile, helpId, scode = exc
                print("The source of the error is", source)
                print("The error message is", text)
                print("More info can be found in {0} (id={1})".format(str(helpFile), helpId))
           

    def OnDisconnection(self, mode, custom):
        print("OnDisconnection")
        self.appHostApp=None


    def OnAddInsUpdate(self, custom):
        print("OnAddInsUpdate", custom)

    def OnStartupComplete(self, custom):
        print("OnStartupComplete", custom)

    def OnBeginShutdown(self, custom):
        print("OnBeginShutdown", custom)



def RegisterAddin(klass):
    if sys.version[:1] == "3":
        import winreg
    else:
        import _winreg as winreg
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Word\\Addins")
    subkey = winreg.CreateKey(key, klass._reg_progid_)
    winreg.SetValueEx(subkey, "CommandLineSafe", 0, winreg.REG_DWORD, 0)
    winreg.SetValueEx(subkey, "LoadBehavior", 0, winreg.REG_DWORD, 3)
    winreg.SetValueEx(subkey, "Description", 0, winreg.REG_SZ, "DocCleaner Word Addin")
    winreg.SetValueEx(subkey, "FriendlyName", 0, winreg.REG_SZ, "DocCleaner Word Addin")

    word = gencache.EnsureDispatch("Word.Application")
    mod = sys.modules[word.__module__]
    print("The module hosting the object is", mod)


def UnregisterAddin(klass):
    if sys.version[:1] == "3":
        import winreg
    else:
        import _winreg as winreg
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Word\\Addins\\" + klass._reg_progid_)
    except WindowsError:
        pass
def main(argv='--register'):
    init_localization()

    import win32com.server.register
    win32com.server.register.UseCommandLine( WordAddin )
    if "--unregister" in sys.argv:
        UnregisterAddin( WordAddin )
    else:

        RegisterAddin( WordAddin )
if __name__ == '__main__':
    main(sys.argv)
