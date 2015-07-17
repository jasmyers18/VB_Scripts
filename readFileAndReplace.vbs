'vbscript to search through each file in a directory and replace all occurances that contain "before" with "after"
'To run. Open a command line prompt to the directory conataining the vbs file. 
'Execute the following command 
'cscript readFileAndReplace.vbs "<Directory>"
'Where <Directory> is the directory with the files you wish to edit

option explicit

dim fso, objFile, strText, strNewText, Folder, file

Const ForReading = 1
Const ForWriting = 2

Folder = Wscript.Arguments(0)

Set fso = CreateObject("Scripting.FileSystemObject")
set Folder = fso.GetFolder(Folder)


For each file in Folder.Files
	
	Set objFile = fso.OpenTextFile(file, ForReading)
	strText = objFile.ReadAll
	objFile.Close
	strNewText = Replace((LCase(strText)), "before", "after")

	Set objFile = fso.OpenTextFile(file, ForWriting)
	objFile.WriteLine strNewText
	objFile.Close

Next	


