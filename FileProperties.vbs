Set FSO = CreateObject("Scripting.FileSystemObject")
Set File_Handler = FSO.CreateTextFile("C:\Ashirwad\FileProperties.txt", True, 8)

For i = 1 to 50
  File_Handler.Writeline("This is line number " & i)
Next
File_Handler.close

Set File_Created = FSO.GetFile("C:\Ashirwad\FileProperties.txt")
DtCreated = File_Created.DateCreated
DtModified = File_Created.DateLastAccessed
FlSize = File_Created.size
FlType = File_Created.Type
FlAtt = File_Created.Attributes

msgbox FlSize

Set File_Handler1 = FSO.CreateTextFile("C:\Ashirwad\FileProperties.csv", True)
File_Handler1.Writeline ("Date Last accessed: ," & DtModified)
File_Handler1.Writeline ("Date Created: ,"& dtCreated)
File_Handler1.Writeline ("File Size: ," & FlSize)
File_Handler1.Writeline ("File Type: ," & FlType)
File_Handler1.Writeline ("File Attribute: ," & FlAtt)