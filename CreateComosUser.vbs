Set WshShell = WScript.CreateObject("WScript.Shell")
'prompt for user input
input = InputBox("Syntax:" & vbnewline & "Firstname (e.g. John)," & vbnewline & "Lastname (e.g. Doe)," & vbnewline & "Email without @siemens.com (e.g. John.Doe)," & vbnewline & "GID (e.g. z002asdf)" & vbnewline & "e.g.: John,Doe,John.Doe,z002asdf","Create new User in AD")
input = replace(input, " ", "") 'delete all blanks in the input
var = split(input,",") 'split the input on all ","
'msgbox "C:\Users\z002mjvx\Desktop\scripts\test.bat " & var(0) & " " & var(1) & " " & var(2) & " " & var(3)'DEBUGGING

'check that the user entered 4 values
if (ubound(var)=3) then
  'call the bat script with the according parameters
	WshShell.Run "C:\Users\z002mjvx\Desktop\scripts\test.bat " & var(0) & " " & var(1) & " " & var(2) & " " & var(3)
else
	msgbox "Illegeal number of arguments." & vbnewline & "Enter 4 comma seperated arguments." & vbnewline & "Creation of user was aborted!"
end if
