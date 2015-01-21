# OpenEditPadProProjects
Windows VB Script that opens one or more "*.EPP" project files in EditPadPro.

#############################################################################
open_editpad_projects_1.vbs
#############################################################################

-----------------------------------------------------------------------------
DESCRIPTION
-----------------------------------------------------------------------------
[v1.02.00] Open EditPad projects.
Opens one or more "*.EPP" project files in EditPad.
(Perhaps one day soon, Jan and the good folks at EditPad will make this script unnecessary!)

Softintheheadware: entertainment and utility in simple but effective packages.
For useful utilities and code, visut us on the Wub at 
https://github.com/softintheheadware/

-----------------------------------------------------------------------------
INSTALLATION
-----------------------------------------------------------------------------
1. Under "GLOBAL CONSTANTS" below, edit the value of c_sExePath 
   to point to the location of EditPad.EXE on your computer
2. Edit the function getProjectFileList to 
   to return a comma-separated list of each ".EPP" file 
   (include the full path for each) you want to open.
3. Create a folder "Scripts" in the root of your C:\ drive. 
4. Create a folder "open_editpad_projects" in the "Scripts" folder.
5. Save this script to "C:\Scripts\open_editpad_projects\open_editpad_projects.vbs"
6. Create a Windows shortcut for the script with: 
   Target: C:\WINDOWS\system32\cscript.exe "C:\Scripts\open_editpad_projects\open_editpad_projects.vbs"
   Start in: C:\Scripts\open_editpad_projects\
7. Copy the Windows shortcut to desired location(s), 
   ie your desktop, Start Menu, QuickLaunch toolbar, pin it to your task bar, etc.
8. To run from command prompt, use CSCRIPT:
   CSCRIPT "C:\Scripts\open_editpad_projects\open_editpad_projects.vbs"

-----------------------------------------------------------------------------
LICENSE
-----------------------------------------------------------------------------
This code may be used freely as the softintheheadware name remains
and you tell everyone how great we are. Donations are welcome, as is praise!
This code is provided free, and we assume no responsibility for any damage 
to your computer, data, or mind, that may result from its use!  
Please use it carefully and responsibly! :-)
For full license, see: 
https://github.com/softintheheadware/OpenEditPadProProjects/blob/master/LICENSE

-----------------------------------------------------------------------------
CHANGE LOG
-----------------------------------------------------------------------------
WHEN           WHO       WHAT
06/26/11 Sun   apple-o   Created script v1.0
04/21/13 Sun   apple-o   Version 1.1 gets project file list from function getProjectFileList instead of constant c_sProjectFileList
12/15/13 Sun   apple-o   Version 1.2 adds sRootPath parameter to getProjectFileList to prompt user for root path
