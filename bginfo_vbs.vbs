strComputer="."
On Error Resume Next
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set IPSettings = objWMIService.ExecQuery ("SELECT * FROM Win32_NetworkAdapterConfiguration where IPEnabled = 'True' and DNSDomain = 'mateus.dc'")
Set IPSettings2 = objWMIService.ExecQuery ("SELECT * FROM Win32_NetworkAdapterConfiguration where IPEnabled = 'True'")
For Each objIPv4 in IPSettings
For i=LBound(objIPv4.IPAddress) to UBound(objIPv4.IPAddress)
If InStr(objIPv4.IPAddress(i),":") = 0 Then Echo objIPv4.IPAddress(i)
NEXT
NEXT