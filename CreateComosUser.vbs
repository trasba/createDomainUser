Set WshShell = WScript.CreateObject("WScript.Shell")

input = InputBox("Syntax:" & vbnewline & "Firstname (e.g. John)," & vbnewline & "Lastname (e.g. Doe)," & vbnewline & "Email without @siemens.com (e.g. John.Doe)," & vbnewline & "GID (e.g. z002asdf)" & vbnewline & "e.g.: John,Doe,John.Doe,z002asdf","Create new User in AD")
var = split(input,",")

if (ubound(var)=3) then
for i=0 to 3
var(i) = trim(var(i))
next 'i

Set wshShell = CreateObject( "WScript.Shell" )
temp = wshShell.ExpandEnvironmentStrings( "%TEMP%")
tmpFile = temp & "\createuser.bat"

Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(tmpFile,True)

objFile.Write "@echo off" & vbCrLf 

objFile.Write " :: BatchGotAdmin" & vbCrLf 
objFile.Write " :-------------------------------------" & vbCrLf
'REM  --> Check for permissions
objFile.Write " >nul 2>&1 "&chr(34)&"%SYSTEMROOT%\system32\cacls.exe"&chr(34)&" "&chr(34)&"%SYSTEMROOT%\system32\config\system"&chr(34)&"" & vbCrLf

'REM --> If error flag set, we do not have admin.
objFile.Write " if '%errorlevel%' NEQ '0' (" & vbCrLf
objFile.Write "     echo Requesting administrative privileges..." & vbCrLf
objFile.Write "    goto UACPrompt" & vbCrLf
objFile.Write " ) else ( goto gotAdmin )" & vbCrLf

objFile.Write " :UACPrompt" & vbCrLf
objFile.Write "     echo Set UAC = CreateObject^("&chr(34)&"Shell.Application"&chr(34)&"^) > "&chr(34)&"%temp%\getadmin.vbs"&chr(34)&"" & vbCrLf
objFile.Write "     set params = %*:"&chr(34)&"="&chr(34)&""&chr(34)&"" & vbCrLf
objFile.Write "     echo UAC.ShellExecute "&chr(34)&"cmd.exe"&chr(34)&", "&chr(34)&"/c %~s0 %*"&chr(34)&", "&chr(34)&chr(34)&", "&chr(34)&"runas"&chr(34)&", 1 >> "&chr(34)&"%temp%\getadmin.vbs"&chr(34)&"" & vbCrLf

objFile.Write "     "&chr(34)&"%temp%\getadmin.vbs"&chr(34) & vbCrLf
objFile.Write "     del "&chr(34)&"%temp%\getadmin.vbs"&chr(34) & vbCrLf
objFile.Write "     exit /B" & vbCrLf

objFile.Write " :gotAdmin" & vbCrLf
objFile.Write "     pushd "&chr(34)&"%CD%"&chr(34) & vbCrLf
objFile.Write "     CD /D "&chr(34)&"%~dp0"&chr(34) & vbCrLf
objFile.Write " :--------------------------------------" & vbCrLf

Include("config.cfg")
objFile.Write "dsadd user " & strDomain & strOptions & vbCrLf

objFile.Write "pause" & vbCrLf
objFile.Write "del "&chr(34)& tmpFile &chr(34) & vbCrLf
objFile.Write "     exit /B" & vbCrLf

objFile.Close
	WshShell.Run chr(34) & tmpFile & chr(34)

else
	msgbox "Illegeal number of arguments." & vbnewline & "Enter 4 comma seperated arguments." & vbnewline & "Creation of user was aborted!"
end if

Sub Include(Byval filename)
  Dim codeToInclude
  Dim FileToInclude

  Const OpenAsDefault = -2
  Const FailIfNotExist = 0
  Const ForReading = 1
  Const OpenFileForReading = 1
  Dim FSO: Set FSO = CreateObject("Scripting.FileSystemObject")

  'Check for existance of include
 If Not FSO.FileExists(filename) Then
    wscript.Echo "Include file not found."
    Set FSO = Nothing
    Exit Sub
  End If

  'open file to include
 Set FileToInclude = FSO.OpenTextFile(filename, ForReading, _
  FailIfNotExist, OpenAsDefault)

  'read all contet of the file
 codeToInclude = FileToInclude.ReadAll
  
  'close file after reading
 FileToInclude.Close
 
  'now cleanup the unused objects
 Set FSO = Nothing
  Set FileToInclude = Nothing

  'now execute code from include file
 ExecuteGlobal codeToInclude
End Sub
