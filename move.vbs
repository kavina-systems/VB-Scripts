' =========================================================
' Tittle: Move files
' Version : v0.1
' Date: 01/010/2009
' Author: K Grey
' Comapny : KSL
' Comment: move files from one folder to another

' Version History  :
' =========================================================

'this will move and rename all file in a folder.
'incase there are more than file a prefix is added to the new filename

On Error Resume next

sdir = "c:\temp\" 'source directory
ddir = "c:\temp1\" 'destination directory
fname = "Sample" ' new file name

' create file list
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(sdir) 'Enter folder path here, single "." uses current directory 
Set fc = f.Files

' List All the Files in a Folder
n=0 'prefix for multi files
For Each f1 in fc
	write.debug f1.Name & " " & vbCrLf 'name and ext
'	write.debug Left(f1.Name,InStr(1,f1.Name,".",1)-1) & vbCrLf 'name only

    'move and rename
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If n=0 then
		objFSO.MoveFile sdir & f1.Name , ddir & fname & ".pdf"
	Else
		objFSO.MoveFile sdir & f1.Name , ddir & fname & "-" & n & ".pdf"	
	End If
	n=n+1
		
Next