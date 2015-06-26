# -*- coding: utf-8 -*-
"""
Created on Mon Jan 12 14:16:57 2015

@author: Bertrand
"""
import shutil
import zipfile
from PIL import Image
from PIL import ImageEnhance
import os
import tempfile
import shutil
import sys
import mimetypes
import getopt
#def openDocument(fileName, subFileName, parser):
#    #opens zip file and getting the subfile (for instance, in a docx, word/document.xml)
#    mydoc = zipfile.ZipFile(fileName)
#    xmlcontent = mydoc.read(subFileName)
#    document = lxml._etree.fromstring(xmlcontent, parser)
#    return document

def createDocument(sourceFile, destFile):
    #Creating a copy of the source document
    shutil.copyfile(sourceFile, destFile)


class img():
    
    def __init__(self, sourceFile):
        self.source = Image.open(sourceFile, mode='r')

    
    def convertTo(self,targetFile, imgformat):
        return self.source.save(targetFile, format=imgformat,optimize=True)
                
        
    def is_wmf(self):
        try:           
            return self.source =='WMF'
            
        except IOError:
            return False
    
    def toRGB(self):        
        return self.source.convert('RGB')

        
    def imgRes(self):        
        return self.source.info


def main(argv):
    
    try:
        opts, args = getopt.getopt(argv, "i:o:", ["input=", "output="])

    except:# getopt.GetoptError:
        
        sys.exit(2)

    inputFile = None
    outputFile = None


    for opt, arg in opts:
        if opt in ("-i", "--input"):
            inputFile = arg
        elif opt in ("-o", "--output"):
            outputFile = arg
        
    folder = tempfile.mkdtemp()       
    createDocument(inputFile, outputFile)
    
    f = zipfile.ZipFile(outputFile, mode='r', compression=zipfile.ZIP_DEFLATED)
    for name in f.namelist():
        
        #print(name)
        f.extract(name, folder)
    f.close()
    os.chdir(folder)  
    imglist = []
    for dirpath, dirnames, filenames in os.walk(folder):
        
        for filename in filenames:
            try:
                docimg = img(os.sep.join([dirpath, filename]))   
                if filename[-4:] in (".wmf", ".jpg"):
                    imglist.append(filename)
                    docimg.convertTo(os.sep.join([dirpath, filename[:-4] +'.jpg' ]), 'JPEG')              
    
                    os.remove(os.sep.join([dirpath, filename]) ) 
            except IOError as e:
                print(str(e))
                #not an image
                pass
    
    #remplacer dans chaque fichier la chaine de caractères filename par filename[:-4] +'.png'
    os.chdir(folder)  
    for dirpath, dirnames, filenames in os.walk(folder):
                
        for filename in filenames:
            mime = mimetypes.guess_type(filename)
    
            if mime[0] in ('text/xml', None):
                print(filename + ' : ' + str(mime[0]))
                fpath = os.path.join(dirpath, filename)
                with open(fpath) as f:
                    s = f.read()     
                
                
                for imgfile in imglist:
                    s = s.replace(imgfile, imgfile[:-4] + '.jpg')
                with open(fpath, 'w') as f:
                    f.write(s)

    
    z= zipfile.ZipFile(outputFile, mode='w', compression=zipfile.ZIP_DEFLATED)
    os.chdir(folder)
    for root, dirs, files in os.walk("."):
        for f in files:
            z.write(os.path.join(root, f))
    os.chdir("..")
    z.close()
    
    try:
        shutil.rmtree(folder)
        print(folder + " deleted")
    except:
        print(folder)                         
        pass
    
    shutil.rmtree(folder)

if __name__ == '__main__':
    main(sys.argv[1:])