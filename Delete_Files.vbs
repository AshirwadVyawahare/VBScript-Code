Delete_Files

Sub Delete_Files
  Set FSO = CreateObject("Scripting.FileSystemObject")
  Set Folder_Handler = FSO.getfolder("D:\Ashirwad")

  Set File_Handler1 = FSO.CreateTextFile("D:\FileProperties1.csv", True)

  For each file1 in Folder_Handler.files
    DtCreated = file1.DateCreated
    
    'If modified date is older than 30 days
    If (DateDiff("d", DtCreated, now) >= 0) Then
      flName = file1.Name
      Flpath = file1.path      
      ConfirmDelete = MsgBox ("Are you sure you want to delete " & flName & " file?", _
      vbYesNoCancel OR VBDefaultButton2, "Delete all files")

      If ConfirmDelete = VbCancel then
        Wscript.Quit
      ElseIf ConfirmDelete = VbYes then
        FSO.DeleteFile(file1.path)  
        File_Handler1.Writeline ("File Name: ," & Flpath)
      End If  
    End If
  Next

End Sub