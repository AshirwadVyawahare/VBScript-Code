Dim objIEDebugWindow

Debug "This is a great way to display intermediate results in a separate window."


Sub Debug( myText )
  ' Uncomment the next line to turn off debugging
  ' Exit Sub

  If Not IsObject( objIEDebugWindow ) Then
    Set objIEDebugWindow = CreateObject( "InternetExplorer.Application" )
    objIEDebugWindow.Navigate "about:blank"
    objIEDebugWindow.Visible = True
    objIEDebugWindow.ToolBar = False
    objIEDebugWindow.Width   = 200
    objIEDebugWindow.Height  = 300
    objIEDebugWindow.Left    = 10
    objIEDebugWindow.Top     = 10
    Do While objIEDebugWindow.Busy
      WScript.Sleep 100
    Loop
    objIEDebugWindow.Document.Title = "IE Debug Window"
    objIEDebugWindow.Document.Body.InnerHTML = _
                 "<b>" & Now & "</b></br>"
  End If

  objIEDebugWindow.Document.Body.InnerHTML = _
                   objIEDebugWindow.Document.Body.InnerHTML _
                   & myText & "<br>" & vbCrLf
End Sub