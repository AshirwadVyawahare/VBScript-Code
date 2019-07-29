Set FSO = CreateObject("Scripting.FileSystemObject")
Set Folder_Handler = FSO.getfolder("C:\Ashirwad\")


msgbox Folder_Handler.files.count
  Set File_Handler1 = FSO.CreateTextFile("C:\Ashirwad\FileProperties1.csv", True)

For each file1 in Folder_Handler.files
  DtCreated = file1.DateCreated
  DtModified = file1.DateLastAccessed
  FlSize = file1.size/1024
  FlType = file1.Type
  flName = file1.Name

  File_Handler1.Writeline ("File Name: ," & flName & ",Date Last accessed: ," & DtModified & ",Date Created: ,"& dtCreated & ",File Size: ," & FormatNumber(FlSize, 2) & " MB,File Type: ," & FlType)
Next
