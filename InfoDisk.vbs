strComputer = "."
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMI.ExecQuery("Select * from Win32_LogicalDisk WHERE DriveType=3")
For Each Item in colItems ' for each drive in the collection
	Echo Item.Name & " " & Round(Item.Freespace/1024/1024/1024) & "/" & Round(Item.Size/1024/1024/1024)
Next
Set colItems = Nothing
Set objWMIService = Nothing