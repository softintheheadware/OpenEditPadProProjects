' ################################################################################################################################################################
' open_editpad_projects_1.vbs
' ################################################################################################################################################################

' -----------------------------------------------------------------------------
' DESCRIPTION
' -----------------------------------------------------------------------------
' [v1.02.00] Open EditPad projects.
' Opens one or more "*.EPP" project files in EditPad.
' (Perhaps one day soon, Jan and the good folks at EditPad will make this script unnecessary!)
' 
' Softintheheadware: entertainment and utility in simple but effective packages.
' For useful utilities and code, visut us on the Wub at 
' https://github.com/softintheheadware/

' -----------------------------------------------------------------------------
' INSTALLATION
' -----------------------------------------------------------------------------
' 1. Under "GLOBAL CONSTANTS" below, edit the value of c_sExePath 
'    to point to the location of EditPad.EXE on your computer
' 2. Edit the function getProjectFileList to 
'    to return a comma-separated list of each ".EPP" file 
'    (include the full path for each) you want to open.
' 3. Create a folder "Scripts" in the root of your C:\ drive. 
' 4. Create a folder "open_editpad_projects" in the "Scripts" folder.
' 5. Save this script to "C:\Scripts\open_editpad_projects\open_editpad_projects.vbs"
' 6. Create a Windows shortcut for the script with: 
'    Target: C:\WINDOWS\system32\cscript.exe "C:\Scripts\open_editpad_projects\open_editpad_projects.vbs"
'    Start in: C:\Scripts\open_editpad_projects\open_editpad_projects.vbs
' 7. Copy the Windows shortcut to desired location(s), 
'    ie your desktop, Start Menu, QuickLaunch toolbar, pin it to your task bar, etc.
' 8. To run from command prompt, use CSCRIPT:
'    CSCRIPT "C:\Scripts\open_editpad_projects\open_editpad_projects.vbs"

' -----------------------------------------------------------------------------
' LICENSE
' -----------------------------------------------------------------------------
' This code may be used freely as the softintheheadware name remains
' and you tell everyone how great we are. Donations are welcome, as is praise!
' This code is provided free, and we assume no responsibility for any damage 
' to your computer, data, or mind, that may result from its use!  
' Please use it carefully and responsibly! :-)
' For full license, see: 
' https://github.com/softintheheadware/OpenEditPadProProjects/blob/master/LICENSE

' -----------------------------------------------------------------------------
' CHANGE LOG
' -----------------------------------------------------------------------------
' WHEN           WHO       WHAT
' 06/26/11 Sun   apple-o   Created script v1.0
' 04/21/13 Sun   apple-o   Version 1.1 gets project file list from function getProjectFileList instead of constant c_sProjectFileList
' 12/15/13 Sun   apple-o   Version 1.2 adds sRootPath parameter to getProjectFileList to prompt user for root path

' ################################################################################################################################################################
' <THE SCRIPT>
' ################################################################################################################################################################

' ================================================================================================================================================================
' GLOBAL CONSTANTS
' ================================================================================================================================================================
const c_sExePath = "U:\Program Files\EditPadPro7\EditPadPro7.exe" ' portable installation
const c_bPromptUserForErrors = true
const c_bDebugMode = false

' ================================================================================================================================================================
' DECLARE GLOBAL VARIABLES
' ================================================================================================================================================================
Dim m_sAppDir
Dim m_sAppFile
Dim m_sAppVersion
Dim m_dStartTime
Dim m_dEndTime
Dim m_sResults
Dim m_lngSeconds
Dim m_iCountOpened

' ================================================================================================================================================================
' SET GLOBAL VARIABLES
' ================================================================================================================================================================

' Path for the current script
m_sAppDir = Left(WScript.ScriptFullName,InstrRev(WScript.ScriptFullName,"\"))

' Filename of the current script
m_sAppFile = NameOnly(WScript.ScriptFullName)

' Script version #
m_sAppVersion = "1.0.0.0"

' Date/Time script started
m_dStartTime = now

' Results
m_sResults = UnescapeText( _ 
	"**********************************************************************\n" & _ 
	cstr(m_dStartTime) & " " & m_sAppFile & " started...\n" & _ 
	"**********************************************************************\n" & _ 
	"")

m_sResults = ""

' ================================================================================================================================================================
' PROMPT USER TO SELECT A ROOT FOLDER
' ================================================================================================================================================================
dim sMenu
dim sValue
dim sRootFolder
sMenu = ""
sMenu = sMenu & "Select root path to load projects from " & vbcrlf
sMenu = sMenu & "1) T:\Users\devuser\Documents" & vbcrlf
sMenu = sMenu & "2) C:\Users\devuser\Documents" & vbcrlf
sMenu = sMenu & "3) H:\Documents" & vbcrlf
sMenu = sMenu & "Else type projects root path" & vbcrlf
sMenu = sMenu & "or blank to quit" & vbcrlf
sValue = InputBox(sMenu, "Root path?", "")  ' prompt, title, default
sValue = trim(sValue)
if (sValue = "1") then
	sRootFolder = "T:\Users\devuser\Documents"
elseif (sValue = "2") then
	sRootFolder = "C:\Users\devuser\Documents"
elseif (sValue = "3") then
	sRootFolder = "H:\Documents"
elseif (sValue = "") then
	sRootFolder = ""
end if

' ================================================================================================================================================================
' OPEN THE FILES
' ================================================================================================================================================================
if len(sRootFolder) > 0 then
	m_iCountOpened = OpenMultipleFilesWithApplication(getProjectFileList(sRootFolder), c_sExePath, c_bPromptUserForErrors)
end if

' ================================================================================================================================================================
' FINISHED
' ================================================================================================================================================================
m_dEndTime = now

m_lngSeconds = datediff("s", m_dStartTime, m_dEndTime)

m_sResults = UnescapeText( _ 
	"************************************************************\n" & _ 
	"COMPLETED SCRIPT `" & m_sAppFile & "`" & "\n" & _
	"************************************************************\n" & _ 
	"Started : " & cstr(m_dStartTime) & "\n" & _ 
	"Finished: " & cstr(m_dEndTime) & "\n" & _ 
	"Duration: " & cstr(m_lngSeconds) & " seconds" & "\n" & _ 
	"Opened  : " & cstr(m_iCountOpened) & " files" & "\n" & _
	"")

'WScript.echo m_sResults
msgbox m_sResults

WScript.Quit()

' ################################################################################################################################################################
' </THE SCRIPT>
' ################################################################################################################################################################

' ################################################################################################################################################################
' <OPEN A FILE IN SPECIFIED PROGRAM>
' ################################################################################################################################################################

' /////////////////////////////////////////////////////////////////////////////
' RECEIVES:
' sRootFolder - can be used to prepend paths to files in list
' EXAMPLES: 
'     "C:\Users\devuser\Documents"
'     "H:\Documents"

Public Function getProjectFileList(sRootFolder) ' as string
	' -----------------------------------------------------------------------------
	' INIT
	Dim sList
	sList = ""
	
	' CONSTANTS
	'Comparison Constants
	'The following comparison constants can be used anywhere in your code in place of actual values:
	'Constant	Value	Description
	const vbBinaryCompare = 0 ' Perform a binary comparison.
	const vbTextCompare = 1 ' Perform a textual comparison.
	const vbDatabaseCompare = 2 ' Perform a comparison based upon information contained in the database where the comparison is to be performed.
	
	' -----------------------------------------------------------------------------
	' REMOVE TRAILING SLASH FROM ROOT FOLDER
	sRootFolder = trim(sRootFolder)
	if right(sRootFolder,1) = "\" then
		sRootFolder = left(sRootFolder, len(sRootFolder)-1)
	end if
	
	' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	' BEGIN OPEN PROJECTS
	' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Main notes
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	AppendString sList, ",", "<sRootFolder/>\00 Main EditPadPro project.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Project X - Monkeys and Matthew Broderick, together again
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	AppendString sList, ",", "U:\Programs\JavaScript\ProjectX\00 ProjectX PROJECT.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Project Y - A digital exorcist for nefarious Wub possessions (proof of concept, not needed for Wubs since they are benevolent, but easier to test with, final version will be used for steve jobs, djinnis and daemons)
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	AppendString sList, ",", "<sRootFolder/>\2014-11-24 ProjectY\00 ProjectY.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "<sRootFolder/>\2014-11-24 ProjectY\04 Python client.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	
'TEMPORARILY DISABLE:
if (true=false) then
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Project Z - Enhance the Wub to enable it to teleport out of harm's way.
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	AppendString sList, ",", "<sRootFolder/>\ProjectZ\00 ProjectZ.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "<sRootFolder/>\ProjectZ\2014-09-30 ProjectZ #2\00 ProjectZ #2.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "<sRootFolder/>\ProjectZ\04 ProjectZ files.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "<sRootFolder/>\ProjectZ\05 ProjectZ files #2.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "<sRootFolder/>\ProjectZ\06 ProjectZ quantum theory files.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "<sRootFolder/>\ProjectZ\11 teleportation.EPP"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "<sRootFolder/>\ProjectZ\12 tardis.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "<sRootFolder/>\ProjectZ\2014-12-09 probability\00 HeartOfGold PROJECT.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "<sRootFolder/>\ProjectZ\2014-11-05 fix variable names\ProjectZ fix var names.EPP"
	wscript.sleep 2000 ' sleep 2 seconds
	
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' ProjectA fix side effect of the 2nd half of Michael Valentine Smith's book, and the time traveling catepillar
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	AppendString sList, ",", "<sRootFolder/>\ProjectA fix side effect\00 ProjectA.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "<sRootFolder/>\ProjectA fix side effect\2014-07-28 fix\uh oh\00 butterfly effect.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	
end if
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' JOHNNY'S EXTRACURRICULAR PROJECTS FOLDER #1
	' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	AppendString sList, ",", "U:\Scripts\00 vbs scripts (PROJECT).epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "U:\00 jbgoode U drive (PROJECT).epp"
	wscript.sleep 2000 ' sleep 2 seconds
	AppendString sList, ",", "T:\Users\jbgoode\Documents\00 johnny T files.epp"
	wscript.sleep 2000 ' sleep 2 seconds
	
	' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	' END OPEN PROJECTS
	' ----------------------------------------------------------------------------------------------------------------------------------------------------------------
	
	' -----------------------------------------------------------------------------
	' INSERT ROOT FOLDER PATH WHEREVER TAG FOUND
	sList = replace(sList, "<sRootFolder/>", sRootFolder)
	
	' -----------------------------------------------------------------------------
	' RETURN LIST OF PROJECTS TO OPEN
	getProjectFileList = sList
	
End Function ' end @getProjectFileList

' /////////////////////////////////////////////////////////////////////////////
' OPEN MULTIPLE FILES WITH A SPECIFIED PROGRAM

Private Function OpenMultipleFilesWithApplication(ByVal sFileList, ByVal sOpenWithApp, ByVal bPromptUserForErrors)
	dim iCount
	dim sRoutineName
	dim arrFiles
	dim lngLoop
	dim bOpened
	
	sRoutineName = "OpenMultipleFilesWithApplication"
	iCount = 0
	arrFiles = split(sFileList, ",")
	for lngLoop = 0 to ubound(arrFiles) ' -1
		'DoEvents
		
		if c_bDebugMode then
			msgbox UnescapeText("" & _ 
				sRoutineName & "\n" & _ 
				"OPENING FILE `" & arrFiles(lngLoop) & "`\n" & _ 
				"IN APP `" & c_sExePath & "`\n" & _ 
				"")
		end if
		
		bOpened = OpenFileWithApplication(arrFiles(lngLoop), c_sExePath, bPromptUserForErrors)
		if bOpened then
			iCount = iCount + 1
		end if
	next ' lngLoop
	
	OpenMultipleFilesWithApplication = iCount
End Function ' OpenMultipleFilesWithApplication

' /////////////////////////////////////////////////////////////////////////////
' OPEN A FILE WITH A SPECIFIED PROGRAM

Private Function OpenFileWithApplication(ByVal sFileToOpen, ByVal sOpenWithApp, ByVal bPromptUserForErrors)
    ' MANUAL ERROR HANDLING
    On Error Resume Next
    
    ' DECLARATIONS
    Dim bResult
    Dim wshShell
    Dim sRun
    
    bResult = false
    
    ' MAKE SURE FILE AND APPLICATION ARE NOT EMPTY
    if ( (len(trim(sOpenWithApp)) > 0) and (len(trim(sFileToOpen)) > 0) ) then
	    ' INITIALIZE
	    Set wshShell = WScript.CreateObject ("WSCript.shell")
	    'Set wshShell = CreateObject("WScript.Shell")
	    
	    ' POPULATE TEMPLATE WITH FILE AND APPLICATION
	    sRun = "`<app/>` `<file/>`"
	    sRun = Replace(sRun, "<app/>", sOpenWithApp)
	    sRun = Replace(sRun, "<file/>", sFileToOpen)
	    sRun = UnescapeText(sRun)
	    
	    ' OPEN THE FILE WITH THE APPLICATION
	    'WScript.Echo sRun
	    'wshShell.run sRun, 1, True
		wshShell.run sRun, 1, False
		'shell.Run Chr(34) & sFile & Chr(34), 6, false
		
		' SHOW ANY ERRORS
	    if err.number = 0 then
	    	bResult = true
	    else
	    	if bPromptUserForErrors then
		    	msgbox "" & _
		    		"THE SHELL COMMAND:" & "wshShell.run " & sRun & ", 1, False" & vbcrlf & _ 
		    		"FAILED WITH ERROR #" & cstr(err.number) & ": " & err.description & vbcrlf & _ 
		    		""
		    end if
		    err.clear
		    bResult = false
	    end if
	else
		bResult = false
    end if ' MAKE SURE FILE AND APPLICATION ARE NOT EMPTY
    
    ' RELEASE OBJECTS FROM MEMORY
    Set wshShell = Nothing
    
    ' RETURN RESULT
    OpenFileWithApplication = bResult
    
End Function ' OpenFileWithApplication

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Running external command (VBS & WSH)
'http://www.virtualhelp.me/scripts/57-vb-script/424-running-external-command-vbs-wsh
'
'There are various ways to execute an external command from VBS. Here are a few examples.
'
'1. Using Exec
'
'If you want to run an application in a child command-shell, providing access to the StdIn/StdOut/StdErr streams:
'
'Example 1: Capturing the exit code:
'    Dim WshShell, oExec
'    Set WshShell = CreateObject("WScript.Shell")
'    Set oExec = WshShell.Exec("notepad")
'    Do While oExec.Status = 0
'        WScript.Sleep 100
'    Loop
'    WScript.Echo oExec.Status
'
'Example 2: Capturing the output:
'    Wscript.Echo runCMD("dir C:\")
'    Function runCMD(strRunCmd)
'    	Set objShell = WScript.CreateObject("WScript.Shell")
'    	Set objExec = objShell.Exec(strRunCmd)
'    	strOut = ""
'    	Do While Not objExec.StdOut.AtEndOfStream
'	        strOut = strOut & objExec.StdOut.ReadLine()
'    	Loop
'    	Set objShell = Nothing
'    	Set objExec = Nothing
'    	runCMD = strOut
'    End Function
'
'2.Using Run
'
'If you want to run a program in a new process:
'
'    object .Run(strCommand, [intWindowStyle], [bWaitOnReturn])
'
'intWindowStyle is an integer value indicating window style. Here's a table of styles:
'
'intWindowStyle  Description  
'0 Hides the window and activates another window. 
'1 Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time. 
'2 Activates the window and displays it as a minimized window. 
'3 Activates the window and displays it as a maximized window. 
'4 Displays a window in its most recent size and position. The active window remains active. 
'5 Activates the window and displays it in its current size and position. 
'6 Minimizes the specified window and activates the next top-level window in the Z order. 
'7 Displays the window as a minimized window. The active window remains active. 
'8 Displays the window in its current state. The active window remains active. 
'9 Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window. 
'10 Sets the show-state based on the state of the program that started the application. 
'
'bWaitOnReturn an option to either wait for the process to return or continue without it (can be either true or false)
'
'Example 1: change directory to C:\ and run dir command:
'    Dim oShell
'    Set oShell = WScript.CreateObject ("WScript.Shell")
'    oShell.run "cmd /K CD C:\ & Dir"
'    Set oShell = Nothing

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' SEE ALSO: 

' How to: Execute a file/program in VBScript
' Published Feb 10 2004, 11:03 AM by Ryan Farley 
' http://customerfx.com/pages/crmdeveloper/2004/02/10/how-to-execute-a-file-program-in-vbscript.aspx

' Executing an EXE inside a VBScript file that has spaces in the path.
' Mar 14 2004 12:13 AM
' http://www.iislogs.com/steveschofield/89240


' ################################################################################################################################################################
' </OPEN A FILE IN SPECIFIED PROGRAM>
' ################################################################################################################################################################





' ################################################################################################################################################################
' <GENERAL FUNCTIONS>
' ################################################################################################################################################################

' /////////////////////////////////////////////////////////////////////////////
' MAKES THE WONDROUS IIf FUNCTION AVAILABLE IN VBSCRIPT...

public function IIf(bTest, ValueIfTrue, ValueIfFalse)
	if bTest = True then
		IIf = ValueIfTrue
	else
		IIf = ValueIfFalse
	end if
end function ' IIf

' /////////////////////////////////////////////////////////////////////////////
' begin #Div
' makes Div function available from within a cell
' NOTE: for Mod, use A = 19 Mod 6.7

' Examples of Div & Mod: 
' iDays = Div(iTotalHours, 24)
' iHours = iTotalHours Mod 24

Public Function Div(ByVal iNumber, ByVal iDivisor) ' As Long
    Div = iNumber \ iDivisor
End Function ' end @Div

' /////////////////////////////////////////////////////////////////////////////

'msgbox "arrData contains:" & vbcrlf & DumpArray(arrData, vbtab & "arrData(<index/>)=<value/>" & vbcrlf)

Public Function DumpArray(ByVal arrData, ByVal sTemplate)
	dim sDump
	dim sNext
	dim iLoop
	
	sDump = ""
    for iLoop = 0 to ubound(arrData)
    	sNext = sTemplate
    	sNext = replace(sNext, "<index/>", cstr(iLoop))
    	sNext = replace(sNext, "<value/>", arrData(iLoop))
    	sDump = sDump & sNext
    next ' iLoop
    
    DumpArray = sDump
End Function ' DumpArray

' /////////////////////////////////////////////////////////////////////////////
' begin #DumpDictionary

' -----------------------------------------------------------------------------
' CHANGE LOG
' WHEN           WHO       WHAT
' 11/27/10 Sat   apple-o   added change log, logging for sub-dictionaries

' EXAMPLE: ?DumpDictionary(sdCols, vbtab, vbtab, "=", vbcrlf)

' EXAMPLE: ?DumpDictionary(sdDict, "", "", "=", ",")

Public Function DumpDictionary(ByRef sdDict, ByVal sOuterIndent, ByVal sInnerIndent, ByVal sPairDelimiter, ByVal sItemDelimiter)
    Dim sMessage 'As String
    Dim sNextValue 'As String
    If Not (sdDict Is Nothing) Then
        Dim varKey 'As Variant
        For Each varKey In sdDict.Keys
            If IsDictionary(sdDict.Item(varKey)) Then
                sNextValue = "(Scripting.Dictionary)" & sItemDelimiter & DumpDictionary2(sdDict.Item(varKey), sInnerIndent, sPairDelimiter, sItemDelimiter, 1)
            Else
                sNextValue = CStr(sdDict.Item(varKey))
            End If
            sMessage = sMessage & IIf(sMessage = "", "", sItemDelimiter) & sOuterIndent & CStr(varKey) & sPairDelimiter & sNextValue
        Next 'varKey
    End If
    DumpDictionary = sMessage
End Function ' end @DumpDictionary
Public Function DumpDictionary2(ByRef sdDict, ByVal sIndent, ByVal sPairDelimiter, ByVal sItemDelimiter, ByVal iIndentLevel)
    Dim sMessage 'As String
    Dim sNextValue 'As String
    If Not (sdDict Is Nothing) Then
        Dim varKey 'As Variant
        For Each varKey In sdDict.Keys
            If IsDictionary(sdDict.Item(varKey)) Then
                sNextValue = "(Scripting.Dictionary)" & sItemDelimiter & DumpDictionary2(sdDict.Item(varKey), sIndent, sPairDelimiter, sItemDelimiter, iIndentLevel + 1)
            Else
                sNextValue = CStr(sdDict.Item(varKey))
            End If
            sMessage = sMessage & IIf(sMessage = "", "", sItemDelimiter) & String(iIndentLevel, sIndent) & CStr(varKey) & sPairDelimiter & sNextValue
        Next 'varKey
    End If
    DumpDictionary2 = sMessage
End Function ' end @DumpDictionary2

' /////////////////////////////////////////////////////////////////////////////
' begin #IsDictionary

' -----------------------------------------------------------------------------
' CHANGE LOG
' WHEN           WHO       WHAT
' 11/27/10 Sat   apple-o   added change log
' 04/23/10 Sat   apple-o   modified for vbs (where all types are variant) to be able to tell if an object is a scripting.dictionary
Public Function IsDictionary(ByRef MyValue) 'As Boolean
    Dim bReturnValue 'As Boolean
    bReturnValue = False
    On Error Resume Next
    if TypeName(MyValue) = "Dictionary" then 
    	bReturnValue = True
    end if
    
    ' VBA version relies on strong typing: 
    'Dim sdDict 'As Scripting.Dictionary
    'Set sdDict = MyValue
    'If Err.Number = 0 Then
    '    bReturnValue = True
    'End If
    'Set sdDict = Nothing
    
    ' VBS version checks for Scripting.Dictionary methods & properties: 
    
    
    IsDictionary = bReturnValue
End Function ' end @IsDictionary

' /////////////////////////////////////////////////////////////////////////////
' begin #PathOnly

' -----------------------------------------------------------------------------
' CHANGE LOG
' WHEN           WHO       WHAT
' 08/01/08       apple-o   created

Private Function PathOnly(ByVal sName)   'As String
    Dim iPos 'As Integer
    sName = Replace(sName, "/", "\")
    iPos = InStrRev(sName, "\")
    If iPos > 0 Then
        sName = Left(sName, iPos)
    End If
    PathOnly = sName
End Function ' end @PathOnly

' /////////////////////////////////////////////////////////////////////////////
' begin #NameOnly

' -----------------------------------------------------------------------------
' CHANGE LOG
' WHEN           WHO       WHAT
' 08/01/08       apple-o   support forward slashes

Private Function NameOnly(ByVal sName)   'As String
    Dim iPos 'As Integer
    sName = Replace(sName, "/", "\")
    iPos = InStrRev(sName, "\")
    If iPos > 0 Then
        sName = Right(sName, Len(sName) - iPos)
    End If
    NameOnly = sName
End Function ' end @NameOnly

' /////////////////////////////////////////////////////////////////////////////
' begin #FileExists

' -----------------------------------------------------------------------------
' CHANGE LOG
' WHEN           WHO       WHAT
' 11/27/10 Sat   apple-o   added change log

Private Function FileExists(ByVal sFile)
    Dim FileSystemObject1 'As FileSystemObject
    Dim bFound 'As Boolean
    'Set FileSystemObject1 = New Scripting.FileSystemObject ' CreateObject("Scripting.FileSystemObject")
    Set FileSystemObject1 = CreateObject("Scripting.FileSystemObject")
    bFound = FileSystemObject1.FileExists(sFile)
    Set FileSystemObject1 = Nothing
    FileExists = bFound
End Function ' end @FileExists

' /////////////////////////////////////////////////////////////////////////////

Private Function IsInt(ByVal MyValue)
    On Error Resume Next
    Dim MyInteger
    MyInteger = CInt(MyValue)
    If Err.Number = 0 Then
        IsInt = True
    Else
        IsInt = False
    End If
    On Error GoTo 0
End Function ' IsInt

' /////////////////////////////////////////////////////////////////////////////

Private Function IsLong(ByVal MyValue)
    On Error Resume Next
    Dim MyLong
    MyLong = CLng(MyValue)
    If Err.Number = 0 Then
        IsLong = True
    Else
        IsLong = False
    End If
    On Error GoTo 0
End Function ' IsLong

' /////////////////////////////////////////////////////////////////////////////

Private Function RemoveFileExtension(ByVal sFile)
    Dim iPos
    iPos = InStrRev(sFile, ".")
    If iPos > 0 Then
        RemoveFileExtension = Left(sFile, iPos)
    Else
        RemoveFileExtension = sFile
    End If
End Function ' RemoveFileExtension

' /////////////////////////////////////////////////////////////////////////////
' NOTES: FOR FURTHER ESCAPING OPTIONS SEE: 
'        http://www.w3.org/MarkUp/html3/latin1.html
'        OR 
'        http://www.w3.org/TR/REC-html40/sgml/entities.html

Private Function UnescapeText(ByVal sText)
    'Debug.Print CStr(Now) & " Starting UnescapeText, Application.ScreenUpdating = " & CStr(Application.ScreenUpdating)
    
    ' SUBSTITUTE TIMESTAMP FOR <now/>
    sText = Replace(sText, "<now/>", CStr(Now))
    
    ' PRESERVE &, escaped \, escaped `
    sText = Replace(sText, "&", "&amp;")
    sText = Replace(sText, "\\", "&92;")
    sText = Replace(sText, "\`", "&126;")
    
    ' UNESCAPE ` TO "
    sText = Replace(sText, "`", Chr(34))
    
    ' UNESCAPE \n TO vbCrLf
    sText = Replace(sText, "\n", vbCrLf)
    
    ' UNESCAPE \t TO vbTab
    sText = Replace(sText, "\t", vbTab)
    
    ' RESTORE &, escaped \, escaped `
    sText = Replace(sText, "&126;", "`")
    sText = Replace(sText, "&92;", "\")
    sText = Replace(sText, "&amp;", "&")
    
    ' RETURN VALUE
    UnescapeText = sText
End Function ' UnescapeText

' /////////////////////////////////////////////////////////////////////////////
' begin #GetEscapedText

' TAKE A STRING AND ESCAPE SPECIAL CHARS
' (FOR PEOPLE WHO ARE NOT GOOD AT REGULAR EXPRESSIONS)
' SUPPORTS THE FOLLOWING (IF YOU NEED OTHERS, ADD THEM)

' CHAR       ESCAPED
' vbcrlf     \n
' vbtab      \t
' \          \\

' -----------------------------------------------------------------------------
' CHANGE LOG:
' DATE           MODIFIED BY       CHANGE
' 06/01/2006     Apple-O           created function
' 03/29/2009     Apple-O           changed fn name from fn_sEscapeChar to GetEscapedText
'
'sText = GetEscapedText(sText)
Private Function GetEscapedText(ByVal sText) 'As String
    sText = Replace(sText, "\", "\\")
    sText = Replace(sText, vbTab, "\t")
    sText = Replace(sText, vbCrLf, "\n")
    GetEscapedText = sText
End Function ' end @GetEscapedText

' /////////////////////////////////////////////////////////////////////////////
' AppendToString
' To call as a sub, saves typing, etc.
Private Sub AppendToString(ByRef sString, ByVal sDelimeter, ByVal sAddition)
    sString = sString & IIf(sString = "", "", sDelimeter) & sAddition
End Sub

' /////////////////////////////////////////////////////////////////////////////
' same as AppendToString

Private Sub AppendString(ByRef sString, ByVal sDelimeter, ByVal sAddition)
    sString = sString & IIf(sString = "", "", sDelimeter) & sAddition
End Sub

' /////////////////////////////////////////////////////////////////////////////

Private Function GetLengthOfLongestString(byval arrList)
	Dim iLoop
	Dim iMaxLength
	iMaxLength = 0
	For iLoop = 0 to ubound(arrList)
		If len(arrList(iLoop)) > iMaxLength Then
			iMaxLength = len(arrList(iLoop))
		End If
	Next ' iLoop
	GetLengthOfLongestString = iMaxLength
End Function ' GetLengthOfLongestString

' /////////////////////////////////////////////////////////////////////////////
'sResult = ErrToString(err, "Error #<Err.Number/>: <Err.Description/>")

Private Function ErrToString(ByVal MyErr, ByVal sTemplate)
    'On Error Resume Next
    Dim sResult
    if sTemplate = "" then 
    	sTemplate = "Error #<Err.Number/>: <Err.Description/>"
    end if
    sResult = sTemplate
    sTemplate = Replace(sTemplate, "<Err.Description/>", MyErr.Description)
    sTemplate = Replace(sTemplate, "<Err.Number/>", MyErr.Number)
    ErrToString = sTemplate
End Function ' ErrToString

' /////////////////////////////////////////////////////////////////////////////

Private Sub ErrToDictionary(ByRef sdDictionary, ByVal MyErr)
    'On Error Resume Next
    if sdDictionary is nothing then 
    	Set sdDictionary = CreateObject("Scripting.Dictionary")
    end if
    sdDictionary.Item("Err.Description") = MyErr.Description
    sdDictionary.Item("Err.Number") = MyErr.Number
End Sub ' ErrToDictionary

' /////////////////////////////////////////////////////////////////////////////
' FROM: http://www.webmasterkb.com/Uwe/Forum.aspx/vbscript/7127/How-can-I-do-VB-Like-by-using-vbscript

Private Function IsNotLike(byval sString, byval sPattern)
	'bull
	
if instr(1, sString, "automation") > 0 then
DebugOutCr "IsNotLike"
end if
	if instr(1, sString, sPattern) > 0 then
		IsNotLike = false
	else
		IsNotLike = true
	end if
	exit function
	
	
	dim regEx
	Set regEx = CreateObject("vbscript.regexp") ' set re = new regexp
	regEx.Global = true
	regEx.IgnoreCase = true
	
	sPattern = replace(sPattern, "\", "\\")
	sPattern = replace(sPattern, ".", "\.")
	sPattern = replace(sPattern, "*", ".*")
	
	'While not incorrect, the "{1}" is unnecessary:
	'sPattern = replace(sPattern, "%", ".{1}")
	sPattern = replace(sPattern, "%", ".")
	
	'sPattern = replace(sPattern, "#", "/d")
	sPattern = replace(sPattern, "#", "\d")
	
	regEx.pattern = sPattern
	IsNotLike = not(regEx.Test(sString))
End Function ' IsNotLike

' /////////////////////////////////////////////////////////////////////////////
' FROM: http://www.webmasterkb.com/Uwe/Forum.aspx/vbscript/7127/How-can-I-do-VB-Like-by-using-vbscript

Private Function IsLike(byval sString, byval sPattern)
	dim regEx
	Set regEx = CreateObject("vbscript.regexp") ' set re = new regexp
	regEx.Global = true
	regEx.IgnoreCase = true
	
	sPattern = replace(sPattern, "\", "\\")
	sPattern = replace(sPattern, ".", "\.")
	sPattern = replace(sPattern, "*", ".*")
	
	'While not incorrect, the "{1}" is unnecessary:
	'sPattern = replace(sPattern, "%", ".{1}")
	sPattern = replace(sPattern, "%", ".")
	
	'sPattern = replace(sPattern, "#", "/d")
	sPattern = replace(sPattern, "#", "\d")
	
	regEx.pattern = sPattern
	IsLike = regEx.Test(sString)
End Function ' isLike

'FROM: http://www.vbscriptexperts.com/general/Q_26776857-Using-RegExp-to-Simulate-Like.jsp
' by: chris_bottomley Posted on 2010-07-28 at 23:45:29ID: 34723685
' LOoking at like for like:

' Will accept control codes between the strings whereas
Function IsLike1(Value, Pattern)
	Dim regEx
	Set regEx = CreateObject("vbscript.regexp")
	With regEx
		.Global = True
		.IgnoreCase = True
		.Pattern = "^" & Replace(Pattern, "*", ".{0,}") & "$"
		IsLike1 = .Test(Value)
	End With
	Set regEx = Nothing
End Function

' WIll ensure the only characters between the components are alphanumerics and common punctuation.
Function IsLike2(ByVal TextVal, ByVal Pattern)
	Dim RXP
	IsLike2 = False
	'Set RXP = New RegExp
	Set RXP = CreateObject("vbscript.regexp")
	With RXP
		.Global = True
		.IgnoreCase = True
		.Pattern = "^" & Replace(Pattern, "*", "[\x20-\x7E]*") & "$"
		IsLike2 = .Test(TextVal)
	End With
	Set RXP = Nothing
End Function

' ################################################################################################################################################################
' </GENERAL FUNCTIONS>
' ################################################################################################################################################################

' ################################################################################################################################################################
' <DRIVE MAPPING FUNCTIONS>
' ################################################################################################################################################################

' /////////////////////////////////////////////////////////////////////////////

Private Sub MapDrive(byval bShowErrorInPopup, byval sLetter, byval sDrive)
	Dim objNetwork
	on error resume next
	
	DisconnectDrive sLetter
	
	Set objNetwork = CreateObject("WScript.Network") 
	objNetwork.MapNetworkDrive sLetter, sDrive
	
	if err.number = 0 then
		WScript.echo cstr(now) & " Mapped " & chr(34) & sLetter & chr(34) & _ 
			" to " & chr(34) & sDrive & chr(34) & "."
	else 
		WScript.echo cstr(now) & " Could not map " & chr(34) & sLetter & chr(34) & _ 
			" to " & chr(34) & sDrive & chr(34) & " due to " & _ 
			"Error #" & cstr(err.number) & ": " & err.Description
		
		if bShowErrorInPopup then 
			msgbox cstr(now) & " Could not map " & chr(34) & sLetter & chr(34) & _ 
				" to " & chr(34) & sDrive & chr(34) & " due to " & _ 
				"Error #" & cstr(err.number) & ": " & err.Description
		end if
		
		err.clear
	end if
	
	set objNetwork = nothing
End Sub ' MapDrive

' SAMPLE CODE:
' Map network drives
'WScript.echo cstr(now) & " Starting: map_network_drives.vbs"
'MapDrive False, "T:", "\\MyServer\MyDept"
''MapDrive "T:", "\\12.34.567.89\MyDept"

' /////////////////////////////////////////////////////////////////////////////
' FileName : disconnectdrive.vbs
' Purpose  : To disconnect specified network drive
' Usage    : Usage  : cscript //nologo disconnectdrive.vbs <drive>
' To Debug : cscript //D //X //nologo disconnectdrive.vbs <drive>
' Example : Usage  : cscript //nologo disconnectdrive.vbs i:
' Date     : 08/01/2008
' Source: http://haripotter.wordpress.com/2008/08/02/how-to-disconnect-a-network-drive-using-vbscript/

Private Function DisconnectDrive(byval strDrive)
	dim bResult
	dim objNetwork
	dim objFSO
	
	on error resume next
	
	bResult = false
	
	set objNetwork = WScript.CreateObject("WScript.Network")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'strDrive = WScript.Arguments(0)
	
	if(objFSO.DriveExists(strDrive)) then
		objNetwork.RemoveNetworkDrive strDrive, True, True
		if err.number = 0 then
			WScript.Echo cstr(now) & " Disconnected network drive: " & strDrive
			bResult = true
		else 
			err.clear
		end if
	end if
	
	if err.number = 0 then err.clear
	
	Set objFSO = nothing
	Set objNetwork = nothing
End function ' DisconnectDrive

' ################################################################################################################################################################
' </DRIVE MAPPING FUNCTIONS>
' ################################################################################################################################################################

