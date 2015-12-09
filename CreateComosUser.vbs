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

objFile.Write "dsadd user " & chr(34) & "CN=" & var(0) & " " & var(1) & ",ou=Cirtix,DC=ft52,DC=lokal" & chr(34) & " -samid " & var(3) & " -pwd COMOS$2015 -mustchpwd yes -memberOf " & chr(34) & "cn=PCS7ESUSER,ou=Cirtix,DC=ft52,DC=lokal" & chr(34) & " " & chr(34) & "cn=ComosCirtixUser,ou=Cirtix,DC=ft52,DC=lokal" & chr(34) & " -fn " & chr(34) & var(0) & chr(34) & " -ln " & chr(34) & var(1) & chr(34) & " -display " & chr(34) & var(0) & " " & var(1) & chr(34) & " -email " & var(2) & "@siemens.com -upn " & var(3) & "@ft52.lokal" & vbCrLf

objFile.Write "pause" & vbCrLf
objFile.Write "del "&chr(34)& tmpFile &chr(34) & vbCrLf
objFile.Write "     exit /B" & vbCrLf

objFile.Close
	WshShell.Run chr(34) & tmpFile & chr(34)

else
	msgbox "Illegeal number of arguments." & vbnewline & "Enter 4 comma seperated arguments." & vbnewline & "Creation of user was aborted!"
end if
