'Script to copy all files that contain the word "log" in the file name to a remote server
'I used the winscp cli to upload the files via sftp
'To run. Open a command line prompt to the directory conataining the vbs file. 
'Execute the following command 
'cscript Jama_Copy.vbs "<Source_Dir>" "<UserID>" "<Password>" "<Destination>"
'Where <Source_Dir> is the directory with the files you wish to copy from. Path needs to end with a "\"
'Where <Destination> is the directory you want to copy to. Path needs to end with a "\"

set WshShell = CreateObject("WScript.Shell")
'WshShell.Run "cmd"
'WshShell.AppActivate "cmd"

FromDir = Wscript.Arguments(0)
UserID = Wscript.Arguments(1)
Pass = Wscript.Arguments(2)
Destination = Wscript.Arguments(3)
TargetDir = "portland05"

'May need an update based on where your WinSCP is installed
winscpPath = ("CD C:\Program Files {(}x86{)}\WinSCP")


'Open Winscp
WshShell.SendKeys winscpPath
WshShell.SendKeys "~"

'Log in to WinSCP
WshShell.SendKeys "winscp.com sftp://" & UserID & ":"& Pass & TargetDir
WshShell.SendKeys "~"

'Upload files
WshShell.SendKeys "put *log* " & Destination 
WshShell.SendKeys "~"

WshShell.SendKeys "exit"
WshShell.SendKeys "~"
