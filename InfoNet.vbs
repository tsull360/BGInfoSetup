strComputer = "."
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
For Each Item in colItems ' for each NIC in the collection
	Echo Item.Description
	If Item.DHCPEnabled(false) Then
		Echo "DHCP Enabled"
	Else
		Echo "DHCP Disabled"
	End If
	Echo "Address: " + Item.IPAddress(0)
	Echo "Gateway: " + Item.DefaultIPGateway(0)
	Echo "Mask: " + Item.IPSubnet(0)
	Echo "DNS: " + Item.DNSServerSearchOrder(0)
Next
Set colItems = Nothing
Set objWMIService = Nothing