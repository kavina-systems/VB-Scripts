' =========================================================
' Tittle: Delete old files
' Version : v0.1
' Date: 05/07/2010
' Author: K Grey
' Comapny : KSL
' Comment: delete files in folder older than 90 days

' Version History  :
' =========================================================

Option Explicit 
On Error Resume Next 
Dim xFSO, xFolder, sDirectoryPath  
Dim xfilesCollection, xfiles, sDir  
Dim NOdays 

' Specify Path
sDirectoryPath = "C:\foldername" 

' Specify Number of Days
NOdays = 90

Set xFSO = CreateObject("Scripting.FileSystemObject") 
Set xFolder = xFSO.GetFolder(sDirectoryPath) 
Set xfilesCollection = xFolder.Files 

For each xfiles in xfilesCollection

If xfiles.DateLastModified < (Date() - NOdays) Then 
xfiles.Delete(True) 
End If      
Next 
