Test

Sub Test
  'Option Explicit 
  Start = Now 
  
  on error resume next

  CONVERSION_FACTOR = 1048576*1024

  Set wshShell = WScript.CreateObject( "WScript.Shell" )
  strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

  If Err.Number <> 0 Then
    Wscript.Echo "Err.Number:", Err.Number    
    Wscript.Echo "Err.Description:", Err.Description    
    Wscript.Echo "Err.source:", Err.source    
    Err.Clear
    Exit Sub
  End If

  Set objWMIService = GetObject("winmgmts://" & strComputerName )
  If Err.Number <> 0 Then
    Wscript.Echo "Err.Number:", Err.Number    
    Wscript.Echo "Err.Description:", Err.Description    
    Wscript.Echo "Err.source:", Err.source    
    Err.Clear
    Exit Sub
  End If

  Set colLogicalDisk = objWMIService.InstancesOf("Win32_LogicalDisk")
  If Err.Number <> 0 Then
    Wscript.Echo "Err.Number:", Err.Number    
    Wscript.Echo "Err.Description:", Err.Description    
    Wscript.Echo "Err.source:", Err.source    
    Err.Clear
    Exit Sub
  End If

  If colLogicalDisk.Count = 0 Then
    Wscript.Echo "No logical drives are installed on this computer."
  Else
    For Each objLogicalDisk In colLogicalDisk
      FreeMegaBytes = objLogicalDisk.FreeSpace / CONVERSION_FACTOR

      'When the items are separated by a comma, a blank space is automatically inserted between the items.
      Wscript.Echo objLogicalDisk.DeviceID , FreeMegaBytes, "GB"
    Next
  End If

  Wscript.Echo colLogicalDisk.Count
  TimeElaps = DateDiff("s", Start, now)
  Wscript.Echo "TimeElaps:", TimeElaps, "Sec"
End Sub
 
