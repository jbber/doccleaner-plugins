# -*- coding: utf-8 -*-
"""
Created on Fri Apr  3 14:13:12 2015

@author: Bertrand
"""

from distutils.core import setup
import py2exe
#mfcfiles = [os.path.join(mfcdir, i) for i in ["mfc90.dll", "mfc90u.dll", "mfcm90.dll", "mfcm90u.dll", "Microsoft.VC90.MFC.manifest"]]
#data_files = [("Microsoft.VC90.MFC", mfcfiles),              ]

setup(console=['wordaddin.py'], 
      packages=['lxml']
      #data_files = data_files
      )