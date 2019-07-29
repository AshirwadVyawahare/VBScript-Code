str = "abcd,ef,gh,ijk,lmnop,"

a =  split(str, ",")
wscript.echo (ubound(a))

  b = instr(str, ",")
  i = 1
  temp_str = mid(str,b+1)

  do 
    b = instr(temp_str,",")
    temp_str = mid(temp_str,b+1)
    If b=0 Then
      exit do
    End If
    i=i+1
    'wscript.echo i
  Loop while true 
wscript.echo i