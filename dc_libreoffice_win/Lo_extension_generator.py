# -*- coding: utf-8 -*-
#This is just a beginning, this is not usable yet!
#The purpose of this script is to generate an oxt extension for LibreOffice, which will use the libreoffice_win.py script available in the directory
import os, sys
import configparser
import tempfile
import zipfile
import doccleaner

class generatingOXT():
    #This class will be used to generate an OXT extension for LibreOffice
    #General variables
    OXT_ID = "myExtension"
    OXT_VERSION = "0.0.1"
    OXT_FULLNAME = "my LO/OoO Extension"
    OXT_LICENCE = "LICENCE NAME"
    OXT_PUBLISHER = "My name"
    OXT_DESCRIPTION = "Short description of my extension"

    config = configparser.ConfigParser()
    conf_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'libreoffice_win_fr.ini')
    config.read(conf_file)

    def AddonsXCU(self):
        #function for generating a menu (addons.xcu)
        #Header
        xmlHeader = r"""
        <?xml version='1.0' encoding='UTF-8'?>
            <oor:component-data
             xmlns:oor="http://openoffice.org/2001/registry"
             xmlns:xs="http://www.w3.org/2001/XMLSchema"
             oor:name="Addons"
             oor:package="org.openoffice.Office">
                <node oor:name="AddonUI">
                    <node oor:name="OfficeMenuBar">
                        <node oor:name="{0}.OfficeMenuBar" oor:op="replace">
                            <prop oor:name="Context" oor:type="xs:string">
                                <value>com.sun.star.text.TextDocument</value>
                            </prop>
                            <prop oor:name="Title" oor:type="xs:string">
                                <value>{1}</value>
                            </prop>
                            <node oor:name="Submenu">""".format(
                                              self.OXT_ID,        #Variable 0
                                              self.OXT_FULLNAME   #Variable 1
                                              )
        #Footer
        xmlFooter = """     </node>
                        </node>
                    </node>
                </node>
            </oor:component-data>"""

        #Body: generating a node for each available XSL sheet in doccleaner
        nodeNumber = 0
        urlNameSpace = r"vnd.sun.star.script" # or r"macro:///" ?

        #TODO: generating menu from the ini file, instead from the content of the docx directory
        for path, subdirs, files in os.walk(os.path.join(os.path.dirname(doccleaner.__file__), 'docx')):#os.walk(os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), "docx")):
            for filename in files:
                if filename.endswith(".xsl"):

                    #Generating the menu label
                    label_translations = self.config.get(str(filename[:-4]), 'label')
                    translation_dictionary = dict(item.split("=") for item in label_translations.split(";"))

                    #Defining the default label for the menu item
                    try:
                        translation_tags = "\n<value>{0}</value>".format(translation_dictionary['default'])
                    except KeyError:
                        #If there is no defined default label, we take the first label defined in the INI file
                        translation_tags = "\n<value>{0}</value>".format(next (iter (translation_dictionary.keys())))

                    #Defining the label translations
                    for key in translation_dictionary.keys():
                        translation_tags += "\n<value xml:lang=\"{0}\">{1}</value>".format(
                                                                                        str(key),                                                            )
                    #TODO: creating a generic script to call with a parameter
                    xmlBody = ""
                    xmlBody += """<node oor:name="M{0}" oor:op="replace">
                            <prop oor:name="Context" oor:type="xs:string">
                                <value>com.sun.star.text.TextDocument</value>
                            </prop>
                            <prop oor:name="URL" oor:type="xs:string">
                                <value>{1}:{2}.{3}?language=Python&amp;location=application</value>
                            <prop>
                            <prop oor:name="Title" oor:type="xs:string">{4}
                            </prop>
                            <prop oor:name="Target" oor:type="xs:string">
                                <value>_self</value>
                            </prop>
                        </node>
                        """.format(str(nodeNumber),      #Variable 0 : unique ID for the generated node
                                   str(urlNameSpace),    #Variable 1 : namespace for the macro URL
                                   str(self.OXT_ID),     #Variable 2 : extension id. We'll use it as macro module name
                                   str(filename),        #Variable 3 : ID (URL) of the macro to call. We'll use the XSL filename as name for the macro, so let's use it - may change later
                                   str(translation_tags) #Variable 4 :
                                   )

                    nodeNumber += 1
        #Generating the final xml
        xml = xmlHeader + xmlBody + xmlFooter
        return xml

    def descriptionXML(self):
        #TODO
        xml = """
                <?xml version='1.0' encoding='UTF-8'?>
                <description
                 xmlns="http://openoffice.org/extensions/description/2006"
                 xmlns:dep="http://openoffice.org/extensions/description/2006"
                 xmlns:xlink="http://www.w3.org/1999/xlink">
                    <identifier value="vnd.{0}.{0}"/>
                    <version value="{1}"/>
                    <!-- <dependencies>
                        <OpenOffice.org-minimal-version value="2.1" dep:name="OpenOffice.org 2.1"/>
                    </dependencies> -->
                    <registration>
                    </registration>
                </description>""".format(
                                         OXT_ID,
                                         OXT_VERSION
                                         )
        return xml

    def manifestXML(self):
        #TODO
        #Generating the manifest.xml in the META-INF folder
        xml = """
                <?xml version="1.0" encoding="UTF-8"?>
                <manifest:manifest>
                 <manifest:file-entry manifest:full-path="{0}/" manifest:media-type="application/vnd.sun.star.basic-library"/>
                 <manifest:file-entry manifest:full-path="pkg-desc/pkg-description.txt" manifest:media-type="application/vnd.sun.star.package-bundle-description"/>
                 <manifest:file-entry manifest:full-path="Addons.xcu" manifest:media-type="application/vnd.sun.star.configuration-data"/>
                 <manifest:file-entry manifest:full-path="Office/UI/BaseWindowState.xcu" manifest:media-type="application/vnd.sun.star.configuration-data"/>
                 <manifest:file-entry manifest:full-path="Office/UI/BasicIDEWindowState.xcu" manifest:media-type="application/vnd.sun.star.configuration-data"/>
                 <manifest:file-entry manifest:full-path="Office/UI/CalcWindowState.xcu" manifest:media-type="application/vnd.sun.star.configuration-data"/>
                 <manifest:file-entry manifest:full-path="Office/UI/DrawWindowState.xcu" manifest:media-type="application/vnd.sun.star.configuration-data"/>
                 <manifest:file-entry manifest:full-path="Office/UI/ImpressWindowState.xcu" manifest:media-type="application/vnd.sun.star.configuration-data"/>
                 <manifest:file-entry manifest:full-path="Office/UI/MathWindowState.xcu" manifest:media-type="application/vnd.sun.star.configuration-data"/>
                 <manifest:file-entry manifest:full-path="Office/UI/StartModuleWindowState.xcu" manifest:media-type="application/vnd.sun.star.configuration-data"/>
                 <manifest:file-entry manifest:full-path="Office/UI/WriterWindowState.xcu" manifest:media-type="application/vnd.sun.star.configuration-data"/>
                </manifest:manifest>
                        """.format(
                                   OXT_ID
                                   )
        return xml

    def pkgDescription(self):
        #TODO
        #Generating the pkg-description.txt file in the pkg-desc folder
        txt = """{0}
                 {1}
                 {2}
                 {3}
        """.format(
                   (".").join(OXT_ID+"-",OXT_VERSION,"oxt"),
                   self.OXT_FULLNAME,
                   self.OXT_LICENCE,
                   self.OXT_DESCRIPTION
                   )

if __name__ == '__main__':
    myOXT = generatingOXT()
    print(myOXT.AddonsXCU())
