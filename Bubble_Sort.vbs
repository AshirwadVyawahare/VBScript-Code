Dim A
A = Array(1,8,3,3,3,4,5,6,8,8,8,8,5,5,5,5,2)
'A = Array(4,2,1,0,1,2,3,4,5,2,1)

For i = 0 to UBound(A)
  For k = 0 to UBound(A)-1
    If (A(k) > A(k+1)) Then
      temp = A(k) 
      A(k) = A(k+1)
      A(k+1) = temp
    End If
  Next
Next

For j = 0 to UBound(A)
  myString = myString & " " & A(j)
Next
MsgBox myString