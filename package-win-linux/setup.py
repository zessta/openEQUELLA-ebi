from distutils.core import setup
import py2exe
import sys,os
import ebi

origIsSystemDLL = py2exe.build_exe.isSystemDLL
def isSystemDLL(pathname):
        if os.path.basename(pathname).lower() in ("msvcp71.dll", "gdiplus.dll", "msvcp90.dll"): 
                return 0
        return origIsSystemDLL(pathname)
py2exe.build_exe.isSystemDLL = isSystemDLL

setup(name = "EQUELLA Bulk Importer (EBI)",
      version = ebi.Version,      
      windows=[{"script":"ebi.py","icon_resources":[(0x0004,"ebismall.ico")],"copyright":"Copyright (c) 2014 Pearson plc. All rights reserved."}],
      description = "EQUELLA Bulk Importer for uploading content into the EQUELLA(R) content management system",
      )

