'vbscript to search through a directory and rename all files that have "contour" in the name with "jama"
'To run. Open a command line prompt to the directory conataining the vbs file. 
'Execute the following command 
'cscript jama_readFileNameAndReplace.vbs "<Directory>"
'Where <Directory> is the directory with the file you wish to edit


option explicit

dim fso, Folder, file

Folder = Wscript.Arguments(0)

Set fso = CreateObject("Scripting.FileSystemObject")
set Folder = fso.GetFolder(Folder)

For each file in Folder.Files
	file.name = Replace((LCase(file.name)), "contour", "jama")

Next	


