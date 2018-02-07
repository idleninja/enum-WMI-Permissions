' set perms in compexp.msc, compmgmt.msc > WMI
'sc sdset scmanager D:(A;;CC;;;AU)(A;;CCLCRPRC;;;IU)(A;;CCLCRPRC;;;SU)(A;;CCLCRPWPRC;;;SY)(A;;KA;;;BA)(A;;KA;;;S-1-5-21-838102356-342305600-1392588124-236951)(A;;CC;;;AC)S:(AU;FA;KA;;;WD)(AU;OIIOFA;GA;;;WD)
'subinacl.exe /service * /grant="<domain_name>\<group_name>" =RLQSE



On Error Resume Next

dim objWMIService 
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\<server_name>\root\cimv2")
launch_location =  left(Wscript.ScriptFullName, len(Wscript.ScriptFullName)-len(Wscript.ScriptName))
isVirtual = false


function LogData(TextFileName, TextToWrite)
	Const forwriting = 2
	Const ForAppending = 8
	Const ForReading = 1
	Set fso = CreateObject("Scripting.FileSystemObject")

	  If fso.fileexists(TextFileName) = False Then
	      'Creates a replacement text file 
	      fso.CreateTextFile TextFileName, True
	  End If

	Set WriteTextFile = fso.OpenTextFile(TextFileName,ForAppending, False)

	WriteTextFile.WriteLine TextToWrite
	WriteTextFile.Close
End Function


' ------------------------- MAIN ---------------------------------
' Win32_Process, Win32_ComputerSystem, Win32_OperatingSystem, Win32_PhysicalMemory, Win32_Processor, Win32_LogicalDisk, Win32_ComputerSystemProduct, Win32_SystemEnclosure, Win32_BIOS, Win32_BaseBoard, Win32_NetworkAdapterConfiguration, Win32_NetworkAdapter, Win32_Service
	
Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_Process")
For Each objWMI in colWMIQuery
    process_info = "Description: " & objWMI.Description & "," _
    & "Creation Date: " & "Creation Date: " & objWMI.CreationDate  & "," _  
    & "CommandLine: "  & objWMI.CommandLine  & "," _ 
    & "Caption: " &  objWMI.Caption   & "," _ 
    & "Priority: " &  objWMI.Priority   & "," _ 
    & "ProcessId: " & objWMI.ProcessId   & "," _ 
    & "ParentProcessId: " & objWMI.ParentProcessId   & "," _ 
    & "Name: "  & objWMI.Name   & "," _ 
    & "ExecutablePath: " & "ExecutablePath: " & objWMI.ExecutablePath 
    LogData launch_location & "\wmi_logging.csv", process_info
    
Next
LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
For Each objWMI in colWMIQuery
    comp_info = "Domain: " & objWMI.Domain & "," _
    & "Name: "  & objWMI.Name   & "," _   
    & "UserName: "  & objWMI.UserName  & "," _ 
    & "Model: " &  objWMI.Model   & "," _ 
    & "Manufacturer: " &  objWMI.Manufacturer 
    LogData launch_location & "\wmi_logging.csv", comp_info

    if objWMI.Model = "Virtual Machine" then
        isVirtual = true
    end if    
Next
LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
For Each objWMI in colWMIQuery
    os_info = "Caption: " & objWMI.Caption & "," _
    & "CSDVersion: "  & objWMI.CSDVersion   & "," _   
    & "Version: "  & objWMI.Version  & "," _ 
    & "Description: " &  objWMI.Description 
    LogData launch_location & "\wmi_logging.csv", os_info
    
Next
LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_PhysicalMemory")
For Each objWMI in colWMIQuery
    mem_info = "DeviceLocator: " & objWMI.DeviceLocator & "," _
    & "TypeDetail: "  & objWMI.TypeDetail & "," _   
    & "FormFactor: "  & objWMI.FormFactor  & "," _ 
    & "MemoryType: " &  objWMI.MemoryType  & "," _ 
    & "DataWidth: " &  objWMI.DataWidth  & "," _ 
    & "TotalWidth: " &  objWMI.TotalWidth  & "," _ 
    & "BankLabel: " &  objWMI.BankLabel  & "," _ 
    & "Status: " &  objWMI.Status  & "," _ 
    & "Speed: " &  objWMI.Speed  & "," _ 
    & "SerialNumber: " &  objWMI.SerialNumber  & "," _ 
    & "PartNumber: " &  objWMI.PartNumber  & "," _ 
    & "Capacity: " &  objWMI.Capacity  & "," _ 
    & "Manufacturer: " &  objWMI.Manufacturer  & "," _ 
    & "Tag: " &  objWMI.Tag  
    LogData launch_location & "\wmi_logging.csv", mem_info
    
Next
LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_PhysicalMemory")
For Each objWMI in colWMIQuery
    mem_info = "DeviceLocator: " & objWMI.DeviceLocator & "," _
    & "TypeDetail: "  & objWMI.TypeDetail & "," _   
    & "FormFactor: "  & objWMI.FormFactor  & "," _ 
    & "MemoryType: " &  objWMI.MemoryType  & "," _ 
    & "DataWidth: " &  objWMI.DataWidth  & "," _ 
    & "TotalWidth: " &  objWMI.TotalWidth  & "," _ 
    & "BankLabel: " &  objWMI.BankLabel  & "," _ 
    & "Status: " &  objWMI.Status  & "," _ 
    & "Speed: " &  objWMI.Speed  & "," _ 
    & "SerialNumber: " &  objWMI.SerialNumber  & "," _ 
    & "PartNumber: " &  objWMI.PartNumber  & "," _ 
    & "Capacity: " &  objWMI.Capacity  & "," _ 
    & "Manufacturer: " &  objWMI.Manufacturer  & "," _ 
    & "Tag: " &  objWMI.Tag  
    LogData launch_location & "\wmi_logging.csv", mem_info
    
Next
LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_Processor")
For Each objWMI in colWMIQuery
    proc_info = "Manufacturer: " &  objWMI.Manufacturer  & "," _
    & "NumberOfCores: "  & objWMI.NumberOfCores & "," _   
    & "MaxClockSpeed: "  & objWMI.MaxClockSpeed  & "," _ 
    & "NumberOfLogicalProcessors: " &  objWMI.NumberOfLogicalProcessors  & "," _ 
    & "Name: " &  objWMI.Name  & "," _ 
    & "AddressWidth: " &  objWMI.AddressWidth
    LogData launch_location & "\wmi_logging.csv", proc_info
    
Next
LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk")
For Each objWMI in colWMIQuery
    disk_info = "Size: " &  objWMI.Size  & "," _
    & "FreeSpace: "  & objWMI.FreeSpace & "," _   
    & "DeviceID: "  & objWMI.DeviceID  & "," _ 
    & "FileSystem: " &  objWMI.FileSystem  & "," _ 
    & "DriveType: " &  objWMI.DriveType  & "," _ 
    & "Description: " &  objWMI.Description & "," _ 
    & "VolumeSerialNumber: " &  objWMI.VolumeSerialNumber & "," _ 
    & "VolumeName: " &  objWMI.VolumeName
    LogData launch_location & "\wmi_logging.csv", disk_info
    
Next
LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_ComputerSystemProduct")
For Each objWMI in colWMIQuery
    sys_info = "UUID: " &  objWMI.uuid  & "," _
    & "IdentifyingNumber: "  & objWMI.IdentifyingNumber
    LogData launch_location & "\wmi_logging.csv", sys_info    
Next
LogData launch_location & "\wmi_logging.csv", vbcrlf


if isVirtual = false then

	Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_SystemEnclosure")
	For Each objWMI in colWMIQuery
	    syst_info = "ChassisTypes: " &  objWMI.ChassisTypes  & "," _
	    & "SerialNumber: "  & objWMI.SerialNumber
	    LogData launch_location & "\wmi_logging.csv", system_info	    
	Next
	LogData launch_location & "\wmi_logging.csv", vbcrlf

end if

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_BIOS")
For Each objWMI in colWMIQuery
    bios_info = "SerialNumber: "  & objWMI.SerialNumber
    LogData launch_location & "\wmi_logging.csv", bios_info    
Next
LogData launch_location & "\wmi_logging.csv", vbcrlf


Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_BaseBoard")
For Each objWMI in colWMIQuery
    baseboard_info = "SerialNumber: "  & objWMI.SerialNumber
    LogData launch_location & "\wmi_logging.csv", baseboard_info    
Next
LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapterConfiguration")
For Each objWMI in colWMIQuery
    netadapt_info = "Index: " &  objWMI.Index  & "," _
    & "DHCPEnabled: "  & objWMI.DHCPEnabled & "," _   
    & "MACAddress: "  & objWMI.MACAddress  & "," _ 
    & "IPSubnet: " &  objWMI.IPSubnet  & "," _ 
    & "IPAddess: " &  objWMI.IPAddress  & "," _ 
    & "Caption: " &  objWMI.Caption & "," _ 
    & "DefaultIPGateway: " &  objWMI.DefaultIPGateway & "," _ 
    & "IPEnabled: " &  objWMI.IPEnabled
    LogData launch_location & "\wmi_logging.csv", netadapt_info
 Next
 LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_NetworkAdapter")
For Each objWMI in colWMIQuery
    network_info = "Index: " &  objWMI.Index  & "," _
    & "Manufacturer: "  & objWMI.Manufacturer & "," _   
    & "NetConnectionID: "  & objWMI.NetConnectionID 
    LogData launch_location & "\wmi_logging.csv", network_info
 Next
 LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_Service")
For Each objWMI in colWMIQuery
    svc_info = "ProcessId: " &  objWMI.ProcessId  & "," _
    & "ServiceType: "  & objWMI.ServiceType & "," _   
    & "StartName: "  & objWMI.StartName  & "," _   
    & "DisplayName: "  & objWMI.DisplayName  & "," _   
    & "State: "  & objWMI.State  & "," _   
    & "StartMode: "  & objWMI.StartMode  & "," _   
    & "PathName: "  & objWMI.PathName  & "," _   
    & "DesktopInteract: "  & objWMI.DesktopInteract  & "," _   
    & "Name: "  & objWMI.Name  & "," _   
    & "AcceptStop: "  & objWMI.AcceptStop  & "," _   
    & "AcceptPause: "  & objWMI.AcceptPause
    LogData launch_location & "\wmi_logging.csv", svc_info
 Next
 LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_TCPIPPrinterPort")
For Each objWMI in colWMIQuery
    svc_info = "HostAddress: " &  objWMI.HostAddress  & "," _
    & "Name: "  & objWMI.Name
    LogData launch_location & "\wmi_logging.csv", svc_info
 Next
 LogData launch_location & "\wmi_logging.csv", vbcrlf

Set colWMIQuery = objWMIService.ExecQuery ("Select * from Win32_Printer")
For Each objWMI in colWMIQuery
    svc_info = "HostAddress: " &  objWMI.PortName  & "," _
    & "PortName: "  & objWMI.PortName & "," _
    & "Name: "  & objWMI.Name
    LogData launch_location & "\wmi_logging.csv", svc_info
 Next
 LogData launch_location & "\wmi_logging.csv", vbcrlf

