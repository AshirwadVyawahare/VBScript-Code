a = fact(5)
msgbox a


Function fact ( num )
  If num = 0 or num = 1 Then
    fact = 1
    Exit function
  End If
  
  fact = num * fact (num-1)
End Function