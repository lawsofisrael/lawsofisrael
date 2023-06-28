'usage
' put files in c:\convert
' create empty folder c:\converted
' run  cscript.exe convert.vbs
' move result c:\converted to dest\converted_docx
' move input c:\convert to dest\convert


' there was a problem problem of path too long more than 256 so i put it in c:\

' Set the directory path where the .doc files are located

Set objWord = CreateObject("Word.Application")
objWord.Visible = False

' Set the directory path where the .doc files are located
strFolderPath = "c:\convert"
strFolderPath2 = "c:\converted"

' Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the folder object
Set objFolder = objFSO.GetFolder(strFolderPath)
Set objFolder2 = objFSO.GetFolder(strFolderPath2)

' Variable to keep track of the total number of .doc files
totalFiles = 0

' Loop through the files in the directory to count the .doc files
For Each objFile In objFolder.Files
    If LCase(objFSO.GetExtensionName(objFile.Path)) = "doc" Then
        totalFiles = totalFiles + 1
    End If
Next

' Variable to keep track of the current file number
currentFile = 0

' Loop through the files in the directory to process .doc files
For Each objFile In objFolder.Files
    If LCase(objFSO.GetExtensionName(objFile.Path)) = "doc" Then
        ' Increment the current file count
        currentFile = currentFile + 1
        
        ' Get the file name without extension
        fileName = objFSO.GetBaseName(objFile.Path)
        
        ' Check if the corresponding .docx file already exists
        If Not objFSO.FileExists(objFSO.BuildPath(objFolder.Path, fileName & ".docx")) _
           and Not objFSO.FileExists(objFSO.BuildPath(objFolder2.Path, fileName & ".docx")) Then
            ' Load the .doc file
            Set objDoc = objWord.Documents.Open( objFile.Path, , True, False )

	    
            ' Save the .doc file as .docx
            strNewFilePath = objFSO.BuildPath(objFolder2.Path, fileName & ".docx")
            objDoc.SaveAs2 strNewFilePath , 16 , , , False

            ' WdSaveFormat_wdFormatDocumentDefault=16 = Word document (.docx) file format
            
            ' Close the .doc file
            objDoc.Close

            ' Print current file number and file name
            WScript.Echo "Processing file " & currentFile & " of " & totalFiles & ": " & fileName
        Else
	    If objFSO.FileExists(objFSO.BuildPath(objFolder.Path, fileName & ".docx")) Then
		objFSO.MoveFile objFSO.BuildPath(objFolder.Path, fileName & ".docx"), objFSO.BuildPath(objFolder2.Path, fileName & ".docx")
	    End If
        
            ' Print skipped file message if .docx file already exists
            WScript.Echo "Skipped file " & currentFile & " of " & totalFiles & ": " & fileName & " (already converted)"
        End If
    End If
Next

' Quit Word application and release the object
objWord.Quit
Set objWord = Nothing

' Clean up objects
Set objFSO = Nothing
Set objFolder = Nothing

WScript.Echo "Conversion completed!"