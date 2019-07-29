Dim A
ReDim Preserve B(0)
A = Array(1,8,3,3,3,4,5,6,8,8,8,8,5,5,5,5)
B(0) = A(0)
For i = 0 to UBound(A)

  for  j = 0 to UBound(B)
    if (B(j) = A(i)) Then
       Exit For
    End If
  Next

  If j > UBound(B) Then
    Redim Preserve B(j)
    B(j) = A(i)
  End If 

Next

for k = 0 to UBound(B)
	myString = myString & " " & B(k)
Next
MsgBox myString