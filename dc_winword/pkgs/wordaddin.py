#coding: utf-8 -*-

#This is just a proof of concept with probably a lot of bugs, use at your own risks!

#It creates a new tab in the MS Word ribbon, with buttons calling docCleaner scripts from inside Word
#To uninstall it, launch it with the --unregister argument
#You can also remove it from Word (in the "Developer" tab, look for COM Addins > Remove)

#Inspired by the Excel addin provided in the win32com module demos, and the "JJ Word Addin" (I don't remember where I get it, but thanks!)

#import pip
import win32com
win32com.__path__
from win32com import universal
from win32com.server.exception import COMException
from win32com.client import gencache, DispatchWithEvents
import winerror
import pythoncom
from win32com.client import constants, Dispatch
import win32com.client
import win32ui
import win32con
import os
import sys
if sys.version[:1] == "3":
    import winreg
else:
    import _winreg as winreg

import win32com.server.register

scriptdir, script = os.path.split(__file__)
pkgdir = os.path.join(scriptdir, 'pkgs')
sys.path.insert(0, pkgdir)
os.environ['PYTHONPATH'] = pkgdir + os.pathsep + os.environ.get('PYTHONPATH', '')
    
#from PIL import Image
import mimetypes
#from guess_language import guess_language
import locale
import gettext
import simplejson
import tempfile
import shutil
import mimetypes
import csv
import doccleaner.doccleaner

#import doccleaner.doccleaner
#from doccleaner import imageConv

#import doccleaner.localization
win32com.client.gencache.is_readonly=False
win32com.client.gencache.GetGeneratePath()
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
    

    pythoncom.CoInitialize()
    _com_interfaces_ = ['_IDTExtensibility2', 'IRibbonExtensibility']
    #add below name of funtions which will be called from winword
    _public_methods_ = ['do', 'apply_style','GetLabel', 'GetLanguage']
    _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER#CLSCTX_INPROC_SERVER
    _reg_clsid_ = "{C5482ECA-F559-45A0-B078-B2036E6F011A}"
    _reg_progid_ = "Python.DocCleaner.WordAddin"
    _reg_policy_spec_ = "win32com.server.policy.EventHandlerPolicy"

    #Convert translations csv to nested dictionary: http://stackoverflow.com/questions/11102326/python-csv-to-nested-dictionary    
    #TODO: nom dynamique pour le répertoire
    #reader = csv.reader(open(os.path.join(os.path.dirname(os.path.realpath(__file__)), "interface_translations.csv")))
    reader = csv.reader(open(os.path.join(os.path.dirname(os.path.realpath(sys.executable)), "interface_translations.csv")))
    translation_dict = {}
    headers = next(reader)[1:]
    for row in reader:
        temp_dict = {}
        
        name = row[0]
        values = []
        
        for x in row[1:]:
            values.append(str(x).encode('windows-1252').decode('utf-8'))
        for i in range(len(values)):
            if values[i]:
                temp_dict[headers[i]] = values[i]
            translation_dict[name] = temp_dict
                
    wd = win32com.client.Dispatch("Word.Application")
    #see list of MS Office language codes (MsoLanguageID Enumeration): http://msdn.microsoft.com/en-us/library/aa432635%28v=office.12%29.aspx
    #Check if MS Word is in french        
    if wd.Application.Language in (1036, 11276, 3084, 12300, 15372, 5132, 13324, 6156, 14348, 8204, 10252, 7180, 9228):
        wd_language = "fr"

    #in spanish
    elif wd.Application.Language in (2058, 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 3082, 19466, 6154, 15370, 10250, 20490, 14346, 8202):
        #wd_language = "es"
        wd_language = "en"
    #If not, we'll use buttons in english for the customized ribbon
    else:
        wd_language = "en"
    
    
    
                 
    def __init__(self):
        self.appHostApp = None
    
    def docFormat(self, filePath):
        
        #check the format of the document, stock it in a variable            
        form = {
                #docx
                ('application/vnd.openxmlformats-officedocument.wordprocessingml.document'):(".docx", 12),
                #doc
                ('application/vnd.ms-word', 'application/doc', 'application/vnd.msword', 'appl/text', 'application/winword', 'application/word', 'application/x-msw6', 'application/x-msword', 'application/msword'):(".doc", 0),
                #odt
                ('application/vnd.oasis.opendocument.text', 'application/x-vnd.oasis.opendocument.text'):(".odt", 23),
                #rtf
                ('application/rtf', 'application/x-rtf', 'text/rtf', 'text/richtext', 'application/x-soffice'):(".rtf", 6)
               }
        
        #mimetypes.guess_type returns a tuple or a string
        if type(mimetypes.guess_type(filePath, strict=True)) == tuple:
            #if it's a tuple, the mimetype is the first element in the tuple
            docMimetype = mimetypes.guess_type(filePath, strict=True)[0]
            
        elif type(mimetypes.guess_type(filePath, strict=True)) == str:
            docMimetype = mimetypes.guess_type(filePath, strict=True)
        
        
        for key in form.keys():
            if docMimetype in key:
                documentFormat = form[key]
                documentExtension = documentFormat[0]
                documentSaveFormat = documentFormat[1]
                break
            else:
                try:
                    documentExtension = mimetypes.guess_extension(docMimetype, strict=True)
                    documentSaveFormat = self.wd.ActiveDocument.SaveFormat
                except:
                    documentFormat = "other"
                
        return (docMimetype, documentExtension, documentSaveFormat)

        

                
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
#    def ConvertImages(self, ctrl):
#        
#        #Creating a word object inside a wd variable
#        wd = win32com.client.Dispatch("Word.Application")
#        wc = win32com.client.constants
#        #If document is not docx, convert it
#        initialPath = wd.ActiveDocument.FullName
#        initialExtension = self.docFormat(wd.ActiveDocument.FullName)[1]                              
#        initialSaveFormat = self.docFormat(wd.ActiveDocument.FullName)[2]  
#        if initialExtension != ".docx":
#            wd.ActiveDocument.SaveAs(FileName = wd.ActiveDocument.Name + '.docx',
#                                     FileFormat = wc.wdFormatXMLDocument )        
#        try:
#            #Check if the file is not a new one (unsaved)
#            if os.path.isfile(wd.ActiveDocument.FullName) == True:
#                #Before processing the doc, let's save the user's last modifications
#                #TODO: ne fonctionne pas correctement
#                wd.ActiveDocument.Save()
#
#                originDoc = wd.ActiveDocument.FullName #:Puts the path of the current file in a variable
#                tmp_dir = tempfile.mkdtemp() #:Creates a temp folder, which will contain the temp docx files necessary for processing
#
#                #Creates a temp  doc, 
#                newDoc = os.path.join(tmp_dir, "~" + wd.ActiveDocument.Name) #:Creates a temporary file (newDoc), which will be the docCleaner output
#
#
#                    
#
#                
#                #If there is more than one XSL sheet, we'll have to make consecutive processings
#
#        
#        
#                img_arguments = ['--input', str(originDoc),
#                                '--output', str(newDoc)                                            
#                                ]
#                                                        
#                imageConv.main(img_arguments)                                    
# 
#          
#
#                #Opening the temp file
#                wd.Documents.Open(newDoc)
#
#                #Copying the temp file content to the original doc
#                #To do this, never use the MSO Content.Copy() and Content.Paste() methods, because :
#                # 1) It would overwrite important data the user may have copied to the clipboard.
#                # 2) Other programs, like antiviruses, may use simulteanously the clipboard, which would result in a big mess for the user.
#                #Instead, use the Content.FormattedText function, it's simple, and takes just one line of code:
#                wd.Documents(originDoc).Content.FormattedText = wd.Documents(newDoc).Content.FormattedText
#                #Closing and removing the temp document
#                wd.Documents(newDoc).Close()
#                os.remove(newDoc)
#
#                #Saving the changes
##                if initialExtension != "docx":
##                    print("bla")
##                else:
#                wd.ActiveDocument.Save()
#                               
#                #Removing the whole temp folder
#                try:
#                    shutil.rmtree(tmp_dir)
#                except:
#                    #TODO: What kind of error would be possible when removing the temp folder? How to handle it?
#                    pass
#                
#            else:
#                win32ui.MessageBox("You need to save the file before launching this script!"
#                ,"Error",win32con.MB_OK)
#
#        except Exception as e:
#
#            tb = sys.exc_info()[2]
#            #TODO: writing the error to a log file
#            win32ui.MessageBox(str(e) + "\n" +
#            str(tb.tb_lineno)+ "\n" +
#            str(newDoc)
#            ,"Error",win32con.MB_OKCANCEL)
        
#        
#    def GetLanguage(self,ctrl):
#        wd = win32com.client.Dispatch("Word.Application")
#        #tests = {}
#        for paragraph in wd.ActiveDocument.Paragraphs:
#            win32ui.MessageBox(str(paragraph),"Error",win32con.MB_OK)
#            if str(paragraph.Style) == "Normal":
#                tests['paraNormal'] = True                
#            elif str(paragraph.Style) in ['Titre', 'Title']:
#                tests['title'] = True
#            elif str(paragraph.Style) in ['langue', 'Language']:
#                tests['lang'] = True
#            elif str(paragraph.Style) == "Pagination":
#                if re.match("([0-9]{1,})(?:[0-9]{2,})", str(paragraph) ) is not None:
#                    win32ui.MessageBox(str("ok"),"Error",win32con.MB_OK)
#                else:
#                    win32ui.MessageBox(str("no"),"Error",win32con.MB_OK)
        
        
#        wdStory = 6
#        wd.Selection.HomeKey(Unit=wdStory)
#        wd.Selection.Find.Text = ""
#        wd.Selection.Find.Text
#        wd.Selection.Find.Execute()
        #docContent = wd.ActiveDocument.Content()        
#        language = guess_language(docContent)
#        
#        #styles = ['','','','']
#        wdStory = 6
#        wd.Selection.HomeKey(Unit=wdStory)
#        try:        
        #wd.ActiveDocument.Selection.Find(Text="liste")
            #1: chercher paragraphe stylé en langue, le remplacer par la valeur trouvée
            #2: 
#        except:
#            print("bla")
        #win32ui.MessageBox(str(language),"Error",win32con.MB_OK) 
        #win32ui.MessageBox(docContent,"Error",win32con.MB_OK) 
    
    def do(self,ctrl):
    #This is the core of the Word addin : manipulates docs and calls docCleaner
    #The ctrl argument is a callback for the button the user made an action on (e.g. clicking on it)

        #Creating a word object inside a wd variable
        wd = win32com.client.Dispatch("Word.Application")
        wc = win32com.client.constants
        #If document is not docx, convert it


        
        try:
            #Check if the file is not a new one (unsaved)
            if os.path.isfile(wd.ActiveDocument.FullName) == True:
                #Before processing the doc, let's save the user's last modifications
                #TODO: ne fonctionne pas correctement
                initialPath = wd.ActiveDocument.FullName
                initialExtension = self.docFormat(initialPath)[1]    
                initialSaveFormat = self.docFormat(initialPath)[2]    
                if initialExtension != ".docx":
                    wd.ActiveDocument.SaveAs(FileName = os.path.join(tempfile.mkdtemp(), wd.ActiveDocument.Name + '.docx'),
                                             FileFormat = wc.wdFormatXMLDocument )
                else:
                    wd.ActiveDocument.Save()

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
                if initialExtension != ".docx":
                    wd.ActiveDocument.SaveAs(FileName = initialPath,
                                             )
                else:
                    wd.ActiveDocument.Save()
                               
                #Removing the whole temp folder
                try:
                    shutil.rmtree(tmp_dir)
                except:
                    #TODO: What kind of error would be possible when removing the temp folder? How to handle it?
                    pass
                
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
        
    def GetScreenTip(self,ctrl):
        return self.translation_dict[ctrl. Tag][self.wd_language+"_screentip"]

    def GetLabel(self,ctrl):
        try:
            label_id = ctrl. Tag
        except:
            label_id = ctrl. Id
        return self.translation_dict[str(label_id)][self.wd_language+"_label"] 
                

def RegisterAddin(klass):
#    if sys.version[:1] == "3":
#        import winreg
#        
#    else:
#        import _winreg as winreg
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

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Word\\Addins\\" + klass._reg_progid_)
    except WindowsError:
        pass
def main(argv='--register'):
    init_localization()
    win32com.server.register.UseCommandLine( WordAddin )
    if "--unregister" in sys.argv:
        UnregisterAddin( WordAddin )
    else:
        RegisterAddin( WordAddin )
    
if __name__ == '__main__':
    #Needed for pynsist - otherwise Python won't find the packages in the pkgs directory

    main(sys.argv)
