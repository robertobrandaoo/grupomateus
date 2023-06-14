strComputer="."
On Error Resume Next
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set IPSettingsDomain = objWMIService.ExecQuery ("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration where (IPEnabled = 'True') AND (DNSDomain='mateus.dc')")
Set IPSettings = objWMIService.ExecQuery ("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration where (IPEnabled = 'True')")
if (IPSettingsDomain = "") then
For Each objIPv4 in IPSettings
For i=LBound(objIPv4.IPAddress) to UBound(objIPv4.IPAddress)
If InStr(objIPv4.IPAddress(i),":") = 0 Then Echo objIPv4.IPAddress(i)
WScript.Echo objIPv4.IPAddress(i)
NEXT
NEXT
else
For Each objIPv4 in IPSettingsDomain
For i=LBound(objIPv4.IPAddress) to UBound(objIPv4.IPAddress)
If InStr(objIPv4.IPAddress(i),":") = 0 Then Echo objIPv4.IPAddress(i)
WScript.Echo objIPv4.IPAddress(i)
NEXT
NEXT
end if
