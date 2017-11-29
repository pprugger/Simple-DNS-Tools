On Error Resume Next
 
strComputer = "yourcomputername"
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colNicConfigs = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
For Each objNicConfig In colNicConfigs
  strDNSSuffixSO = ""
  strDNSServerSO = ""
  strDNSHostName = objNicConfig.DNSHostName
  strIndex = objNicConfig.Index
  strDescription = objNicConfig.Description
  strDNSDomain = objNicConfig.DNSDomain
  strDNSSuffixSO = ""
  If Not IsNull(objNicConfig.DNSDomainSuffixSearchOrder) Then
    For Each strDNSSuffix In objNicConfig.DNSDomainSuffixSearchOrder
      strDNSSuffixSO = strDNSSuffixSO & VbCrLf & String(37, " ") & _
 strDNSSuffix
    Next
  End If
  strDNSServerSO = ""
  If Not IsNull(objNicConfig.DNSServerSearchOrder) Then
    For Each strDNSServer In objNicConfig.DNSServerSearchOrder
      strDNSServerSO = strDNSServerSO & VbCrLf & String(37, " ") & _
 strDNSServer
    Next
  End If
  strDomainDNSRegistrationEnabled = _
 objNicConfig.DomainDNSRegistrationEnabled
  strFullDNSRegistrationEnabled = objNicConfig.FullDNSRegistrationEnabled
  strDNSSettings = strDNSSettings & VbCrLf & VbCrLf & _
   "  Network Adapter " & strIndex & VbCrLf & _
   "    " & strDescription & VbCrLf & VbCrLf & _
   "    DNS Domain:                      " & strDNSDomain & VbCrLf & _
   "    DNS Domain Suffix Search Order:" & strDNSSuffixSO & VbCrLf & _
   "    DNS Server Search Order:" & strDNSServerSO & VbCrLf & _
   "    Domain DNS Registration Enabled: " & _
   strDomainDNSRegistrationEnabled & VbCrLf & _
   "    Full DNS Registration Enabled:   " & _
   strFullDNSRegistrationEnabled
Next
 
WScript.Echo VbCrLf & "DNS Settings" & VbCrLf & VbCrLf & _
 "Host Name: " & strDNSHostName & strDNSSettings


