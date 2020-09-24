"""
MIT License

Copyright © 2020 Matthew Murphy

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
"""

import os
import requests
import winreg
import string
import configparser
import time
import keyboard

from bs4 import BeautifulSoup
from zipfile import ZipFile

retailVersionURL = "https://www.tukui.org/download.php?ui=elvui" #Scrape Retail ElvUI version number from here
retailDownloadURL = "https://www.tukui.org/downloads/" #Download actual Retail ElvUI zip file from here
classicVersionURL = "https://www.tukui.org/classic-addons.php?id=2"
classicDownloadURL = "https://www.tukui.org/classic-addons.php?download=2"

retailFileName = "elvui-%s.zip" #This is the naming convention for Retail ElvUI downloads. "%s" is the version number that we scrape from retailVersionURL
classicFileName = "elvui_classic.zip" #This is the default name of the latest version of ElvUI for classic, regardless of version number

configpath = '%s\\ElvUIDownloader\\' %  os.environ['APPDATA']
configname = 'elvuidownloaderconfig.ini'

config = configparser.ConfigParser()
InstallRetail = False
lastRetailVersion = ""
InstallClassic = False
lastClassicVersion = ""
WarcraftInstallDir = ""

def DownloadElvUI(downloadlink, filename, InterfaceAddonsFolder, version):
    #Download the latest version of ElvUI and write it to the working directory
    try:
        r2 = requests.get(downloadlink, allow_redirects=True)
        open(filename, 'wb').write(r2.content)
            
        #Extract to addons directory
        with ZipFile(filename, 'r') as zipObj:
            zipObj.extractall(WarcraftInstallDir+InterfaceAddonsFolder)
        os.remove(filename)
        print("Finished downloading %s!" % filename+version)
    except:
        print("ERROR: Failed to download %s" % downloadlink)

def DoRetail():
    print("Downloading latest ElvUI for Retail & extracting to Interface\Addons folder...")
    
    #First we have to scrape the current version number from the download page
    #this segment is ugly but it ain't broke so I ain't fixing it lmao
    try:
        r = requests.get(retailVersionURL)
        soup = BeautifulSoup(r.content, features="lxml")
        version = ""
        for row in soup.find('b', attrs = {'class':'Premium'}):
            version = row
        
        #Compare latest version to last downloaded version so we don't waste time
        global lastRetailVersion
        if version == lastRetailVersion:
            print("You should already have the latest version of ElvUI for Retail (%s)!" % version)
        else:
            #Use the scraped version number to set the filename for the download, assuming they don't change the naming convention
            lastRetailVersion = version
            WriteConfig()
            filename = retailFileName % version
            InterfaceAddonsFolder = "\_retail_\Interface\Addons"
            
            DownloadElvUI(retailDownloadURL+filename, filename, InterfaceAddonsFolder, "")
    except:
        print("ERROR: Failed to retrieve %s" % retailVersionURL)
   
def DoClassic():
    #The Classic download always comes from the same link so we don't really need to do much here
    print("Downloading latest ElvUI for Classic & extracting to Interface\Addons folder...")
    
    try:
        r = requests.get(classicVersionURL)
        soup = BeautifulSoup(r.content, features="lxml")
        version = ""
        for row in soup.find('b', attrs = {'class':'VIP'}):
            version = row
        
        global lastClassicVersion
        if version == lastClassicVersion:
            print("You should already have the latest version of ElvUI for Classic (%s)!" % version)
        else:
            lastClassicVersion = version
            WriteConfig()
            filename = classicFileName
            InterfaceAddonsFolder = "\_classic_\Interface\Addons"
            
            DownloadElvUI(classicDownloadURL, filename, InterfaceAddonsFolder, " (%s)" % version)
    except:
        print("ERROR: classicfail")
    
def DoCleanup():
    print("TODO")
    
def WriteConfig():
    config['SETTINGS'] = {
        'WarcraftInstallDir': WarcraftInstallDir,
        'InstallRetail': InstallRetail,
        'lastRetailVersion': lastRetailVersion,
        'InstallClassic': InstallClassic,
        'lastClassicVersion': lastClassicVersion
    }
    with open(configpath + configname, 'w') as configfile:
        config.write(configfile)
    
def ReadConfig():
    
    #Gotta globalize
    global InstallRetail
    global lastRetailVersion
    global InstallClassic
    global lastClassicVersion
    global WarcraftInstallDir
    
    #Because "False" == True
    stringToBool = {
        "True": True,
        "False": False
    }
    
    config.read(configpath + configname)
    settings = config['SETTINGS']
    WarcraftInstallDir = settings['WarcraftInstallDir']
    InstallRetail = stringToBool[settings['InstallRetail']]
    lastRetailVersion = settings['lastRetailVersion']
    InstallClassic = stringToBool[settings['InstallClassic']]
    lastClassicVersion = settings['lastClassicVersion']
    
def DoConfig():

    #Gotta globalize
    global InstallRetail
    global InstallClassic
    global WarcraftInstallDir

    print("Configuring settings...\I'm gonna ask you for some basic info.")
        
    #Try to find the World of Warcraft installation directory via the Windows registry
    try:
        regkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\World of Warcraft")
        WarcraftInstallDir = winreg.QueryValueEx(regkey, "InstallLocation")[0]
        print("\nWorld of Warcraft installation found in: "+WarcraftInstallDir+"!")
    except:
        WarcraftInstallDir = input("\nI couldn't find your WoW installation!\nPlease input the location of 'World of Warcraft Launcher.exe' (where your '_retail_' and/or '_classic_' folders are).\nINPUT: ")
    
    #Prompt for Retail/Classic ElvUI installation
    print("\nDo you want this program to download ElvUI for Retail? (Y/N)")
    while True:
        if keyboard.is_pressed("y"):
            InstallRetail = True
            print("Okay, I'll download ElvUI for Retail.")
            break
        elif keyboard.is_pressed("n"):
            InstallRetail = False
            print("Okay, I'll skip Retail.")
            break
            
    #Quick sleep here to prevent reading the same keystroke for both inputs
    time.sleep(1) 
    
    print("\nDo you want this program to download ElvUI for Classic? (Y/N)")
    while True:
        if keyboard.is_pressed("y"):
            InstallClassic = True
            print("Okay, I'll download ElvUI for Classic.")
            break
        elif keyboard.is_pressed("n"):
            InstallClassic = False
            print("Okay, I'll skip Classic.")
            break
    
    #Sleep here just so there's a brief pause after collecting the inputs to make it look nice I guess?
    time.sleep(1)
    
    #If we're configred to install either version of ElvUI
    if InstallRetail or InstallClassic:
        #Create the appdata directory to store the config file in
        if not os.path.exists(configpath):
            os.makedirs(configpath)
        WriteConfig()
    else:
        #If we're NOT configured to install ElvUI, clean up all traces of the program for clean removal
        if os.path.exists(configpath + configname):
            print("\nLooks like I'm not configured to download ElvUI at all, so I'll go ahead and delete my config file so you can remove me completely.")
            os.remove(configpath + configname)
            os.rmdir(configpath)
            time.sleep(0.5)
        else:
            print("\nIf you don't want me to download ElvUI for you, then why did you download me in the first place? You silly goose.")
            time.sleep(0.5)
        

def InitializeDownloader():
    #Check to see if we've already got a config file. If not, set one up.
    if not os.path.exists(configpath + configname):
        try:
            DoConfig()
        except OSError:
            #TODO Run basic mode if DoConfig fails for any reason.
            print ("ERROR: Can't create a config file! Running in BASIC mode for now (I'll have ask you what you want to install every time you run the program).")
    else:
        #Give the user the chance to change their settings before proceeding
        print("Initializing... Press SHIFT now if you want to change your settings.")
        waittime = time.time() + 3
        ReadConfig()
        while time.time() < waittime:
            if keyboard.is_pressed("shift"):
                DoConfig()
    if InstallRetail:
        DoRetail()
    if InstallClassic:
        DoClassic()
        
    #We're done here.
    print("Closing in 3 seconds...")
    time.sleep(3)
                
InitializeDownloader()