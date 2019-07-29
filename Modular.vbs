n = inputbox("Enter number")

i = n
a = 10

do While i/10 > 9
  a = a * 10  
  If i/a <= 9 Then
    Exit do
  End If
Loop

digit = n

do while a >= 1
  n = int(digit/a)
  digit = digit mod a
  
  a = a/10
  msgbox n
Loop

  'msgbox int(n)
