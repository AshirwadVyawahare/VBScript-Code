On Error Resume Next
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists("C:\Ashirwad\ScriptLog.txt") Then
  Set objFile = objFSO.GetFile("C:\Ashirwad\ScriptLog.txt", 4)
Else
  Set objFile = objFSO.OpenTextFile("C:\Ashirwad\ScriptLog.txt",8, true)
  If err.number<>0 Then
    wscript.echo err.description, err.number, err.source
  End If
  wscript.echo "created new file successfully"
End If
 
option1 = inputbox("what operation you want to perform on file")
option11 = ucase(option1)

Select case option11
Case "DELETE"
  ConfirmDelete = MsgBox ("Are you sure you want to delete these files?", _
  Vbyesno OR VBDefaultButton2, "Delete all files")

  If ConfirmDelete = VbNo then
    Wscript.Quit
  End If  
  Set objFile = objFSO.DeleteFile("C:\Ashirwad\ScriptLog.txt")  
  wscript.echo "deleted newly file successfully"
Case "WRITE"
  objFile.WriteLine("Im being fooled")
Case Else
  wscript.echo "Wrong operation is provided."
End Select