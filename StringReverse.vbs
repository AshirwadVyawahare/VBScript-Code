a = RevStr(" ABCED ")
msgbox a


Function RevStr (Str)
  If Str = " " OR len(str) = 1 Then
    RevStr = str
    Exit Function
  End If

  RevStr = RevStr(mid(str,2)) & left(str,1)
End Function