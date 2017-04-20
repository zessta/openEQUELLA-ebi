#!/usr/bin/env python
#Boa:App:ebi

# EQUELLA BULK IMPORTER (EBI)
# Author: Jim Kurian, Pearson plc.
#
# This program creates items in EQUELLA based on metadata specified in a csv file. The csv file
# must commence with a row of metadata xpaths (e.g. metadata/name,metadata/description,...).
# Subsequent rows should contain the metadata that populate the element identified by the
# xpath in the column header.
#
# metadata/title,metadata/description
# Our House,"This is a picture of my house, my lawn, my cat and my dog"
# Our Car,This is a picture of my car
#
# User Guide
# A detailed and comprehensive user guide is distributed with this program. Please see this for more
# detailed usage instructions and troubleshooting tips.

# system settings (do not change!)
Version = "4.71"
Copyright = "Copyright (c) 2014 Pearson plc. All rights reserved."
License = """
THE EQUELLA(R) BULK IMPORTER PROGRAM IS PROVIDED UNDER THE TERMS OF THIS LICENSE ("AGREEMENT"). ANY USE OF THE PROGRAM CONSTITUTES THE RECIPIENT'S ACCEPTANCE OF THIS AGREEMENT.

1. DEFINITIONS 
"Vendor" means Pearson plc. and its subsidiaries worldwide.

"Program" means all versions of the EQUELLA Bulk Importer software program, a software program designed for managing data in the EQUELLA(R) software program. 

"Recipient" means anyone who receives the Program under this Agreement. 

2. NO WARRANTY

THE PROGRAM IS PROVIDED ON AN "AS-IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, EITHER EXPRESS OR IMPLIED INCLUDING, WITHOUT LIMITATION, ANY WARRANTIES OR CONDITIONS OF TITLE, NON-INFRINGEMENT OR FITNESS FOR A PARTICULAR PURPOSE.
The Recipient is solely responsible for determining the appropriateness of using the Program and assumes all risks associated with its exercise of rights under this Agreement, including but not limited to the risks and costs of program errors, compliance with applicable laws, damage to or loss of data, programs or equipment, and unavailability or interruption of operations. 

3. USAGE AND DISTRIBUTION

The Recipient of the Program may use it for the management of the EQUELLA software program within the Recipient's organization only. The Recipient may only distribute the Program within the Recipient's organization.
"""
EBIDownloadPage = "http://maestro.equella.com/items/eb737eb2-ac6f-4ba3-af17-321ee6c305a1/0/"
Debug = False
SuppressVersion = False

import sys, os, traceback
import wx
import MainFrame
import ConfigParser
import platform

modules ={'MainFrame': [1, u'Main frame of EBI', 'none://MainFrame.py']}

propertiesFile = "ebi.properties"
if sys.path[0].endswith(".zip"):
    propertiesFile = os.path.join(os.path.dirname(sys.path[0]), propertiesFile)
else:
    propertiesFile = os.path.join(sys.path[0], propertiesFile)
            
display = True

class ebi(wx.App):
    global display

    def OnInit(self):
        self.main = MainFrame.create(None)
        if display:
            self.main.Show()
            self.SetTopWindow(self.main)
        return True

def alert(message):
    message = message + "\n\n" + Copyright
    app = wx.PySimpleApp()
    dialog = wx.Dialog(None, title='EQUELLA Bulk Importer ' + Version, size=wx.Size(500, 300), style=wx.RESIZE_BORDER|wx.DEFAULT_DIALOG_STYLE)
    dialog.Center()
    box = wx.BoxSizer(wx.VERTICAL)
    dialog.SetSizer(box)
    txtMessage = wx.TextCtrl(dialog, -1, message, style=wx.TE_MULTILINE|wx.TE_READONLY)    
    box.Add(txtMessage, 1, wx.ALIGN_CENTER|wx.ALL|wx.EXPAND)
    
    btnsizer = wx.StdDialogButtonSizer()
    btn = wx.Button(dialog, wx.ID_OK)
    btn.SetDefault()
    btnsizer.AddButton(btn)
    btnsizer.Realize()
    box.Add(btnsizer, 0, wx.ALIGN_CENTER|wx.CENTER|wx.ALL, 5)   
    btn.SetFocus() 
    
    dialog.ShowModal()
    dialog.Destroy()
    app.MainLoop()
      

def main():
    try:
        global display
        global Version
        global Debug
        global SuppressVersion
        global propertiesFile
        
        # create properties file
        config = ConfigParser.ConfigParser()
        config.read(propertiesFile)
        if not "Configuration" in config.sections():
            config.add_section('Configuration')
            config.set('Configuration','LoadLastSettingsFile', 'False')
            config.write(open(propertiesFile, 'w'))
        else:
            try:
                if config.has_option('Configuration','debug'):
                    Debug = config.getboolean('Configuration','debug')
            except:
                alert("Error reading properties file for debug setting: " + str(sys.exc_info()[1]))

        # usage syntax to display in command line
        usageSyntax = """USAGE:
    
ebi.py [-start] [-test] [<filename>]
ebi.exe [-start] [-test] [<filename>]

<filename>
Run the EBI visually and load the specified settings file.

-start <filename>
Run the EBI non-visually using the specified settings file.

-test <filename>
Run the EBI non-visually in test mode using the specified settings file, no items will be submitted to EQUELLA.
        """

        settingsFile = ""

        if "?" in sys.argv:
            alert(usageSyntax)
            
        else:
            usageCorrect = True
            i = 1

            #loop through arguments (skip first argument as it is the command itself)
            while i < len(sys.argv):
                if sys.argv[i] in ["-test", "-start"]:

                    # check that an argument exists after the -start or -test argument
                    if i + 1 <= (len(sys.argv) - 1):

                        # check that argument after -settings does not start with a dash
                        if sys.argv[i + 1][0] != "-":

                            # get settings filename
                            settingsFile = sys.argv[i + 1]

                            # step over filename
                            i += 1
                        else:
                            usageCorrect = False
                    else:
                        usageCorrect = False
                        
                elif sys.argv[i] not in ["-test", "-start"]:
                    settingsFile = sys.argv[i]
                    # usageCorrect = False

                i += 1

            if "-test" in sys.argv and usageCorrect :
                # run non-visually in test mode
                # print "Testing " + settingsFile
                try:
                    display = False
                    application = ebi(0)
                    if SuppressVersion:
                        Version = ""
                    application.main.createEngine(Version, Copyright, License, EBIDownloadPage, propertiesFile)                
                    application.main.setDebug(Debug)
                    if settingsFile != "":
                        if application.main.loadSettings(settingsFile):
                            application.main.startImport(testOnly = True)
                except:
                    exceptionType, exceptionValue, exceptionTraceback = sys.exc_info()
                    errorString = "ERROR: " + str(exceptionValue)
                    if Debug:
                        errorString += ': ' + ''.join(traceback.format_exception(exceptionType, exceptionValue, exceptionTraceback))
                    alert(errorString)
                
            elif "-start" in sys.argv and usageCorrect :
                try:
                    # run non-visually
                    # print "Running " + settingsFile
                    display = False
                    application = ebi(0)
                    if SuppressVersion:
                        Version = ""
                    application.main.createEngine(Version, Copyright, License, EBIDownloadPage, propertiesFile)
                    application.main.setDebug(Debug)
                    if settingsFile != "":
                        if application.main.loadSettings(settingsFile):
                            application.main.startImport(testOnly = False)
                except:
                    exceptionType, exceptionValue, exceptionTraceback = sys.exc_info()
                    errorString = "ERROR: " + str(exceptionValue)
                    if Debug:
                        errorString += ': ' + ''.join(traceback.format_exception(exceptionType, exceptionValue, exceptionTraceback))
                    alert(errorString)            
                
            elif usageCorrect:
                try:
                    # run visually starting with the main form
                    application = ebi(0)
                    if SuppressVersion:
                        Version = ""
                    application.main.createEngine(Version, Copyright, License, EBIDownloadPage, propertiesFile)                
                    application.main.setDebug(Debug)
                    if settingsFile != "":
                        application.main.loadSettings(settingsFile)
                    elif "Configuration" in config.sections() and config.getboolean('Configuration','loadlastsettingsfile'):
                        try:
                            application.main.loadSettings(config.get('State','settingsfile'))
                        except:
                            if Debug:
                                alert(str(sys.exc_info()[1]))
                        
                    application.MainLoop()
                except:
                    exceptionType, exceptionValue, exceptionTraceback = sys.exc_info()
                    errorString = "ERROR: " + str(exceptionValue)
                    if Debug:
                        errorString += ': ' + ''.join(traceback.format_exception(exceptionType, exceptionValue, exceptionTraceback))
                    alert(errorString)
            else:
                alert(usageSyntax)
    except:
        exceptionType, exceptionValue, exceptionTraceback = sys.exc_info()
        errorString = "ERROR: " + str(exceptionValue)
        if "Errno 30" in errorString:
            errorString += "\n\nEBI cannot read and write files in its current location. Try installing EBI in a different location."
            if platform.system() == "Darwin":
                errorString += "\n\nIf launching EBI from a mounted disk image (*.dmg) first copy the EBI package to Applications or another local directory."
        if Debug:
            errorString += "\n\n" + ''.join(traceback.format_exception(exceptionType, exceptionValue, exceptionTraceback))                
        alert(errorString)


if __name__ == '__main__':
    main()
