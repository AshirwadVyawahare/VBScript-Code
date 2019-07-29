Read one by one charactor from notepad:
Set obj=createobject("Scripting.Filesystemobject")
set ob1=obj.OpenTextFile("C:\Documents and Settings\Administrator\Desktop\test1.txt")
While ob1.AtEndOfStream <> True
  a=ob1.Read(1)
  msgbox a
Wend
====================================================

