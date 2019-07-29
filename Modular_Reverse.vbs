Dim modular()
num =  inputbox("Enter any number")

IF isnumeric(num)=true Then
  'msgbox("Numeric ")& isnumeric(num)
  msgbox isnumeric(num)
End If
i=0
do while num > 9
  ReDim Preserve modular(i)
  modular(i) = num mod 10
  num = int(num/10)

  i=i+1
Loop

  ReDim Preserve modular(i)
  modular(i) = num

For i = UBound(modular) to 0 step -1
  msgbox modular(i)
Next