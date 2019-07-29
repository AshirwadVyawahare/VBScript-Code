Set FSO = CreateObject("Scripting.FileSystemObject")
Set FName = FSO.OpenTextFile("C:\Documents and Settings\Ashirwad.Vyawahare\Desktop\Text.txt", 8, true)

for i = 1 to 5 
  for k = 1 to 5-i      
    Str = Str  & "+"            
  Next 

                   
  for l = 1 to 2*i-1      
    Str = str & "#"           
  Next 
  
  'new line charactor
  'Str = str & vbcr
  FName.Writeline (str)
  Str1 = str1 & vbcr & str
  str = ""
Next

for i = 1 to 5
  for k = 1 to i       
    Str = Str  & "+"           
  Next 

  for k = 1 to (5-i)*2-1
    Str = str & "#"
  Next
  'Str = str & vbcr
  FName.Writeline (str)
  Str1 = str1 & vbcr & str
  str = ""
Next
msgbox str1

