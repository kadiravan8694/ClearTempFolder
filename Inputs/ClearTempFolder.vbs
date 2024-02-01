Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
'strFolderPath = Replace(objShell.SpecialFolders("Desktop"), "Desktop", "local settings\temp")
strFolderPath = WScript.CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)
'msgbox(strFolderPath)
Set objFolder = objFSO.GetFolder(strFolderPath)
On Error Resume Next
 
For Each objFile In objFolder.Files
	objFSO.DeleteFile objFile
Next
 
For Each objFolder In objFolder.subFolders
	objFSO.DeleteFolder objFolder
Next