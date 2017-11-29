

On Error Resume Next
 
strComputer = "yourcomputername"
arrNewDNSServerSearchOrder = Array()
 
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colNicConfigs = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE Index = 1")
 
'WScript.Echo VbCrLf & "Computer: " & strComputer
 
For Each objNicConfig In colNicConfigs
 ' WScript.Echo VbCrLf & "  Network Adapter " & objNicConfig.Index
  'WScript.Echo "    DNS Server Search Order - Before:"
  If Not IsNull(objNicConfig.DNSServerSearchOrder) Then
    For Each strDNSServer In objNicConfig.DNSServerSearchOrder
      'WScript.Echo "        " & strDNSServer
    Next
  End If
  intSetDNSServers = _
   objNicConfig.SetDNSServerSearchOrder(arrNewDNSServerSearchOrder)
  If intSetDNSServers = 0 Then
    'WScript.Echo "    Replaced DNS server search order list."
  Else
    'WScript.Echo "    Unable to replace DNS server search order list."
  End If
Next
  
Set colNicConfigs = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE Index = 1")
 
For Each objNicConfig In colNicConfigs
  'WScript.Echo VbCrLf & "  Network Adapter " & objNicConfig.Index
  'WScript.Echo "    DNS Server Search Order - After:"
  If Not IsNull(objNicConfig.DNSServerSearchOrder) Then
    For Each strDNSServer In objNicConfig.DNSServerSearchOrder
      'WScript.Echo "        " & strDNSServer
    Next
  End If
Next
