for i = 1 to 5 
  for k = 1 to 5-i      'it will print the spaces equal to the number entered by the user [1] 
    Str = Str & " "           'printf(" "); 
  Next 

                   
  for l = 1 to 2*i-1      'first line will print 1 * only and will increase by 1 each time [2] 
    Str = str & "1"           'printf("%s","*"); 
  Next 
  
  'new line charactor
  Str = str & vbcr
  
Next

for i = 1 to 5
  for k = 1 to i      'it will print the spaces equal to the number entered by the user [1] 
    Str = Str & " "           'printf(" "); 
  Next 

  for k = 1 to (5-i)*2-1
    Str = str & "1"
  Next
  Str = str & vbcr
Next
msgbox str

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FName = FSO.OpenTextFile("C:\Ashirwad\Text.txt", 8, true)

FName.Writeline (str)