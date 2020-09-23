Dim fso, startFolder, OlderThanDate

Set fso = CreateObject("Scripting.FileSystemObject")
' Folder to start deleting - subfolders will also be cleaned
startFolder = "C:\inetpub\logs\LogFiles\" 
'keep last 7 days log file - adjust if needed
OlderThanDate = DateAdd("d", -7, Date)  
DeleteOldFiles startFolder, OlderThanDate

Function DeleteOldFiles(folderName, BeforeDate)
   Dim folder, file, fileCollection, folderCollection, subFolder

   Set folder = fso.GetFolder(folderName)
   Set fileCollection = folder.Files
   For Each file In fileCollection
      If file.DateLastModified < BeforeDate Then
         fso.DeleteFile(file.Path)
      End If
   Next

    Set folderCollection = folder.SubFolders
    For Each subFolder In folderCollection
       DeleteOldFiles subFolder.Path, BeforeDate
    Next
End Function
