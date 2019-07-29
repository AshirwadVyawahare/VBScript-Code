Test()
WScript.Quit(0)	



Function Test()
	Dim sString	'By doing this here the variable is now "Private"
			'(or so we are lead to believe)

	sString="hello "
	MsgBox "sString is =" & sString
	
	Call DoStuff( sString )  
	
	MsgBox "sString is =" & sString	'Should still be "hello " if the variable is Private
End Function
	
	
'Function DoStuff(sString)		'This is the same as "Function DoStuff(ByRef sString)"
'Function DoStuff(ByVal sString)	'This will make it that sString variable Private
Function DoStuff(ByRef sString)	'This is the same as "Function DoStuff( sString)"
	sString=sString&sString&sString
	DoStuff=True
End Function
