@echo off

:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %*", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"
:--------------------------------------


REm set /p fn=Enter First Name:
REm set /p ln=Enter Last Name:
REm set /p email=Enter Email:
REm set /p gid=Enter logon:
REm dsadd user "CN=%fn% %ln%,ou=Cirtix,DC=ft52,DC=lokal" -samid %gid% -pwd COMOS$2015 -mustchpwd yes -memberOf "cn=PCS7ESUSER,ou=Cirtix,DC=ft52,DC=lokal" "cn=ComosCirtixUser,ou=Cirtix,DC=ft52,DC=lokal" -fn %fn% -ln %ln% -display "%fn% %ln%" -email %email%@siemens.com -upn %gid%@ft52.lokal
dsadd user "CN=%1 %2,ou=Cirtix,DC=ft52,DC=lokal" -samid %4 -pwd COMOS$2015 -mustchpwd yes -memberOf "cn=PCS7ESUSER,ou=Cirtix,DC=ft52,DC=lokal" "cn=ComosCirtixUser,ou=Cirtix,DC=ft52,DC=lokal" -fn %1 -ln %2 -display "%1 %2" -email %3@siemens.com -upn %4@ft52.lokal

REM echo on
REM echo %1
pause
