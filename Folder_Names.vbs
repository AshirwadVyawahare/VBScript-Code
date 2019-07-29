Folder_Path = "D:\Ashirwad\"   'Write desired path here

Set FSO = CreateObject("Scripting.FileSystemObject")
Set Folder_Handler = FSO.getfolder(Folder_Path)

Set File_Handler1 = FSO.CreateTextFile(Folder_Path & "FileProperties1.csv", True)
Set colSubfolders = Folder_Handler.Subfolders

File_Handler1.Writeline ("Folders: ")

For Each objSubfolder in colSubfolders
    File_Handler1.Writeline ("Folder Name: ," & objSubfolder.Name)
Next

File_Handler1.Writeline ("Files:")

For each file1 in Folder_Handler.files
    File_Handler1.Writeline ("File Name: ," & file1.Name & ",File Type: ," & file1.Type)
Next