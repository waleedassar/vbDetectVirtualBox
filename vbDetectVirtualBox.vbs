' https://twitter.com/waleedassar
' http://waliedassar.com/
' Simple WMI WQL queries for detecting VirtualBox VM's
VBoxFound = False

set objX = GetObject("winmgmts:\\.\root\cimv2")

' Win32_NetworkAdapterConfiguration aka NICCONFIG
Set NicQ = objX.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration")
For Each Nic in NicQ
	if Not IsNull(Nic.MACAddress) And Not IsNull(Nic.Description) Then
		MacAddress = LCase(CStr(Nic.MACAddress))
		Description = LCase(CStr(Nic.Description))
		'We want to detect the VirtualBox guest, not the host
		If InStr(1,MacAddress,"08:00:27:") = 1 And InStr(1,Description,"virtualbox") = 0 Then
		   WScript.Echo "Win32_NetworkAdapterConfiguration ==> Nic.MACAddress: " & Nic.MACAddress
		   VBoxFound = True
		End If
	End If
Next

'Win32_SystemDriver aka sysdriver
Set SySDrvQ = objX.ExecQuery("SELECT * FROM Win32_SystemDriver")
For Each SysDrv in SysDrvQ
    DescSysDrv = SysDrv.Description
	DispSysDrv = SysDrv.DisplayName
    NameSysDrv = SysDrv.Name
	PathSysDrv = SysDrv.PathName
	If Not IsNull(DescSysDrv) Then
	   If DescSysDrv = "VirtualBox Guest Driver" Or DescSysDrv = "VirtualBox Guest Mouse Service" Or DescSysDrv = "VirtualBox Shared Folders" Or DescSysDrv = "VBoxVideo" Then
	      WScript.Echo "Win32_SystemDriver ==> SysDrv.Description ==> " & DescSysDrv
		  VBoxFound = True
	   End If
	End If
	
	If Not IsNull(DispSysDrv) Then
	   If DispSysDrv = "VirtualBox Guest Driver" Or DispSysDrv = "VirtualBox Guest Mouse Service" Or DispSysDrv = "VirtualBox Shared Folders" Or DispSysDrv = "VBoxVideo" Then
	      WScript.Echo "Win32_SystemDriver ==> SysDrv.DisplayName ==> " & DispSysDrv
		  VBoxFound = True
	   End If
	End If
	
	If Not IsNull(NameSysDrv) Then
	   If NameSysDrv = "VBoxGuest" Or NameSysDrv = "VBoxMouse" Or NameSysDrv = "VBoxSF" Or NameSysDrv = "VBoxVideo" Then
	      WScript.Echo "Win32_SystemDriver ==> SysDrv.Name ==> " & NameSysDrv
		  VBoxFound = True
	   End If
	End If
	
    If Not IsNull(PathSysDrv) Then
	   PathSysDrv_l = LCase(PathSysDrv)
	   If InStr(1,PathSysDrv_l,"vboxguest.sys") > 0 Or InStr(1,PathSysDrv_l,"vboxmouse.sys") > 0 Or InStr(1,PathSysDrv_l,"vboxsf.sys") > 0 Or InStr(1,PathSysDrv_l,"vboxvideo.sys") > 0 Then
	      WScript.Echo "Win32_SystemDriver ==> SysDrv.PathName ==> " & PathSysDrv
		  VBoxFound = True
	   End If
	End If
Next

' Win32_NTEventLog aka NTEventLog
Set EvtLogQ = objX.ExecQuery("SELECT * FROM Win32_NTEventlogFile")
For Each EvtLogX In EvtLogQ
    If Not IsNull(EvtLogX) Then
	   FileNameEvtX = CStr(EvtLogX.FileName)
	   FileNameEvtX_l = LCase(FileNameEvtX)
	   If FileNameEvtX_l = "sysevent" Or FileNameEvtX_l = "system" Then
	      SourcesEvtX = EvtLogX.Sources
		  For Each SourceEvtX in SourcesEvtX
		      SourceEvtX_l = LCase(CStr(SourceEvtX))
			  If SourceEvtX_l = "vboxvideo" Then
			     WScript.Echo "Win32_NTEventlogFile ==> EvtLogX.Sources ==> " & SourceEvtX
				 VBoxFound = True
			  End If
		  Next
	   End If
	End If
Next

' Win32_BIOS aka bios
Set BiosQ = objX.ExecQuery("SELECT * FROM Win32_BIOS")
For Each Bios in BiosQ
    If Not IsNull(Bios) Then
	   If Not IsNull(Bios.Manufacturer) Then
	      ManufacturerBios = LCase(CStr(Bios.Manufacturer))
		  If InStr(1,ManufacturerBios,"innotek gmbh") > 0 Then
		     WScript.Echo "Win32_BIOS ==> Bios.Manufacturer ==> " & Bios.Manufacturer
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(Bios.SMBIOSBIOSVersion) Then
	      SMBIOSBIOSVersionBios = LCase(CStr(Bios.SMBIOSBIOSVersion))
		  If InStr(1,SMBIOSBIOSVersionBios,"virtualbox") > 0 Then
		     WScript.Echo "Win32_BIOS ==> Bios.SMBIOSBIOSVersion ==> " & Bios.SMBIOSBIOSVersion
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(Bios.Version) Then
	      VersionBios = LCase(CStr(Bios.Version))
		  If InStr(1,VersionBios,"vbox   - 1") > 0 Then
		     WScript.Echo "Win32_BIOS ==> Bios.Version ==> " & Bios.Version
			 VBoxFound = True
		  End If
	   End If
	End If
Next

' Win32_DiskDrive aka diskdrive
Set DiskDriveQ = objX.ExecQuery("SELECT * FROM Win32_DiskDrive")
For Each DiskDrive in DiskDriveQ
    If Not IsNull(DiskDrive) Then
	   If Not IsNull(DiskDrive.Model) Then
	      ModelDskDrv = LCase(DiskDrive.Model)
		  If ModelDskDrv = "vbox harddisk" Then
		     WScript.Echo "Win32_DiskDrive ==> DiskDrive.Model ==> " & DiskDrive.Model
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(DiskDrive.PNPDeviceID) Then
	      PNPDeviceIDDskDrv = LCase(DiskDrive.PNPDeviceID)
		  If InStr(1,PNPDeviceIDDskDrv,"diskvbox") > 0 Then
		     WScript.Echo "Win32_DiskDrive ==> DiskDrive.PNPDeviceID ==> " & DiskDrive.PNPDeviceID
			 VBoxFound = True
		  End If
	   End If
	End If
Next

' Win32_StartupCommand aka Startup
Set StartupQ = objX.ExecQuery("SELECT * FROM Win32_StartupCommand")
For Each Startup in StartupQ
    If Not IsNull(Startup) Then
	   If Not IsNull(Startup.Caption) Then
	      CaptionStartup = LCase(CStr(Startup.Caption))
		  If CaptionStartup = "vboxtray" Then
		     WScript.Echo "Win32_StartupCommand ==> Startup.Caption ==> " & Startup.Caption
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(Startup.Command) Then
	      CommandStartup = LCase(CStr(Startup.Command))
		  If InStr(1,CommandStartup,"vboxtray.exe") > 0 Then
		     WScript.Echo "Win32_StartupCommand ==> Startup.Command ==> " & Startup.Command
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(Startup.Description) Then
	      DescStartup = LCase(CStr(Startup.Description))
		  If DescStartup = "vboxtray" Then
		     WScript.Echo "Win32_StartupCommand ==> Startup.Description ==> " & Startup.Description
			 VBoxFound = True
		  End If
	   End If
	End If
Next

'Win32_ComputerSystem aka ComputerSystem
Set ComputerSystemQ = objX.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each ComputerSystem in ComputerSystemQ
    If Not IsNull(ComputerSystem) Then
	   If Not IsNull(ComputerSystem.Manufacturer) Then
	      ManufacturerComputerSystem = LCase(CStr(ComputerSystem.Manufacturer))
		  If ManufacturerComputerSystem  = "innotek gmbh" Then
		     WScript.Echo "Win32_ComputerSystem ==> ComputerSystem.Manufacturer ==> " & ComputerSystem.Manufacturer
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(ComputerSystem.Model) Then
	      ModelComputerSystem = LCase(CStr(ComputerSystem.Model))
		  If ModelComputerSystem  = "virtualbox" Then
		     WScript.Echo "Win32_ComputerSystem ==> ComputerSystem.Model ==> " & ComputerSystem.Model
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(ComputerSystem.OEMStringArray) Then
	      OEMStringArrayComputerSystem = ComputerSystem.OEMStringArray
	      For Each OEM In OEMStringArrayComputerSystem
		      OEM_l = LCase(OEM)
			  If InStr(1,OEM_l,"vboxver_") > 0 Or InStr(1,OEM_l,"vboxrev_") > 0 Then
			     WScript.Echo "Win32_ComputerSystem ==> ComputerSystem.OEMStringArray ==> " & OEM
				 VBoxFound = True
			  End If
		  Next
	   End If
	End If
Next

'Win32_Service aka service
Set ServiceQ = objX.ExecQuery("SELECT * FROM Win32_Service")
For Each Service in ServiceQ
    If Not IsNull(Service) Then
	   If Not IsNull(Service.Caption) Then
	      CaptionService = LCase(CStr(Service.Caption))
		  If CaptionService = "virtualbox guest additions service" Then
		     WScript.Echo "Win32_Service ==> Service.Caption ==> " & Service.Caption
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(Service.DisplayName) Then
	      DisplayNameService = LCase(CStr(Service.DisplayName))
		  If DisplayNameService = "virtualbox guest additions service" Then
		     WScript.Echo "Win32_Service ==> Service.DisplayName ==> " & Service.DisplayName
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(Service.Name) Then
	      NameService = LCase(CStr(Service.Name))
		  If NameService = "vboxservice" Then
		     WScript.Echo "Win32_Service ==> Service.Name ==> " & Service.Name
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(Service.PathName) Then
	      PathNameService = LCase(CStr(Service.PathName))
		  If InStr(1,PathNameService,"vboxservice.exe") > 0 Then
		     WScript.Echo "Win32_Service ==> Service.PathName ==> " & Service.PathName
			 VBoxFound = True
		  End If
	   End If
	End If
Next


'Win32_LogicalDisk aka LogicalDisk
Set LogicalDiskQ = objX.ExecQuery("SELECT * FROM Win32_LogicalDisk")
For Each LogicalDisk in LogicalDiskQ
    If Not IsNull(LogicalDisk) Then
	   If Not IsNull(LogicalDisk.DriveType) Then
	      If LogicalDisk.DriveType = 3 Then
		     If Not IsNull(LogicalDisk.VolumeSerialNumber) Then
			    VolumeSerialNumberLogicalDisk = LCase(LogicalDisk.VolumeSerialNumber)
				If VolumeSerialNumberLogicalDisk = "fceae0a3" Then
			       WScript.Echo "Win32_LogicalDisk ==> LogicalDisk.VolumeSerialNumber ==> " & LogicalDisk.VolumeSerialNumber
				   VBoxFound = True
				End If
			 End If
		  ElseIf LogicalDisk.DriveType = 5 Then
		     If Not IsNull(LogicalDisk.VolumeName) Then
			    VolumeNameLogicalDisk = LCase(LogicalDisk.VolumeName)
				'Volume name should be "VBOXADDITIONS_4."
				If InStr(1,VolumeNameLogicalDisk,"vboxadditions") > 0 Then
			       WScript.Echo "Win32_LogicalDisk ==> LogicalDisk.VolumeName ==> " & LogicalDisk.VolumeName
				   VBoxFound = True
				End If
			 End If		  
		  End If
	   End If
	End If
Next

'Win32_LocalProgramGroup
Set LogicalProgramGroupQ = objX.ExecQuery("SELECT * FROM Win32_LogicalProgramGroup")
For Each LocalProgramGroup in LogicalProgramGroupQ
    If Not IsNull(LocalProgramGroup) Then
	   NameLocalProgramGroup = LCase(LocalProgramGroup.Name)
	   If InStr(1,NameLocalProgramGroup,"oracle vm virtualbox guest additions") > 0 Then
	      WScript.Echo "Win32_LogicalProgramGroup ==> LocalProgramGroup.Name ==> " & LocalProgramGroup.Name
		  VBoxFound = True
	   End If
	End If
Next



'Win32_NetworkAdapter aka NIC
Set NicQQ = objX.ExecQuery("SELECT * FROM Win32_NetworkAdapter")
For Each NIC_x in NicQQ
	if Not IsNull(NIC_x.MACAddress) And Not IsNull(NIC_x.Description) Then
		MacAddress_x = LCase(CStr(NIC_x.MACAddress))
		Description_x  = LCase(CStr(NIC_x.Description))
		'We want to detect the VirtualBox guest, not the host
		If InStr(1,MacAddress_x,"08:00:27:") = 1 And InStr(1,Description_x,"virtualbox") = 0 Then
		   WScript.Echo "Wow: Win32_NetworkAdapter ==> NIC.MacAddress: " & NIC_x.MACAddress
		   VBoxFound = True
		End If
	End If
Next


'Win32_Process aka process
Set ProcessQ = objX.ExecQuery("SELECT * FROM Win32_Process")
For Each Process in ProcessQ
    If Not IsNull(Process) Then
	   If Not IsNull(Process.Description) Then
	      DescProcess = LCase(Process.Description)
		  If DescProcess = "vboxservice.exe" Or DescProcess = "vboxtray.exe" Then
		     WScript.Echo "Win32_Process ==> Process.Description ==> " & Process.Description
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(Process.Name) Then
	      NameProcess = LCase(Process.Name)
		  If NameProcess = "vboxservice.exe" Or NameProcess = "vboxtray.exe" Then
		     WScript.Echo "Win32_Process ==> Process.Name ==> " & Process.Name
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(Process.CommandLine) Then
	      CmdProcess = LCase(Process.CommandLine)
		  If InStr(1,CmdProcess,"vboxservice.exe") > 0 OR InStr(1,CmdProcess,"vboxtray.exe") > 0 Then
		     WScript.Echo "Win32_Service ==> Process.CommandLine ==> " & Process.CommandLine
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(Process.ExecutablePath) Then
	      ExePathProcess = LCase(Process.ExecutablePath)
		  If InStr(1,ExePathProcess,"vboxservice.exe") > 0 OR InStr(1,ExePathProcess,"vboxtray.exe") > 0 Then
		     WScript.Echo "Win32_Service ==> Process.ExecutablePath ==> " & Process.ExecutablePath
			 VBoxFound = True
		  End If
	   End If
	End If
Next

'Win32_BaseBoard aka BaseBoard
Set BaseBoardQ = objX.ExecQuery("SELECT * FROM Win32_BaseBoard")
For Each BaseBoard in BaseBoardQ
    If Not IsNull(BaseBoard) Then
	   If Not IsNull(BaseBoard.Manufacturer) Then
	      ManufacturerBaseBoard = LCase(BaseBoard.Manufacturer)
		  If ManufacturerBaseBoard = "oracle corporation" Then
		     WScript.Echo "Win32_BaseBoard ==> BaseBoard.Manufacturer ==> " & BaseBoard.Manufacturer
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(BaseBoard.Product) Then
	      ProductBaseBoard = LCase(BaseBoard.Product)
		  If ProductBaseBoard = "virtualbox" Then
		     WScript.Echo "Win32_BaseBoard ==> BaseBoard.Product ==> " & BaseBoard.Product
			 VBoxFound = True
		  End If
	   End If
	End If
Next

'Win32_SystemEnclosure aka SystemEnclosure
Set SystemEnclosureQ = objX.ExecQuery("SELECT * FROM Win32_SystemEnclosure")
For Each SystemEnclosure in SystemEnclosureQ
    If Not IsNull(SystemEnclosure) Then
	   If Not IsNull(SystemEnclosure.Manufacturer) Then
	      ManufacturerSystemEnclosure = LCase(SystemEnclosure.Manufacturer)
		  If ManufacturerSystemEnclosure = "oracle corporation" Then
		     WScript.Echo "Win32_SystemEnclosure ==> SystemEnclosure.Manufacturer ==> " & SystemEnclosure.Manufacturer
			 VBoxFound = True
		  End If
	   End If
	End If
Next

'Win32_CDROMDrive aka cdrom
Set CDRomQ = objX.ExecQuery("SELECT * FROM Win32_CDROMDrive")
For Each CDRom in CDRomQ
    If Not IsNull(CDRom) Then
	   If Not IsNull(CDRom.Name) Then
	      NameCDRom = LCase(CDRom.Name)
		  If NameCDRom = "vbox cd-rom" Then
		     WScript.Echo "Win32_CDROMDrive ==> CDRom.Name ==> " & CDRom.Name
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(CDRom.VolumeName) Then
	      VolumeNameCDRom = LCase(CDRom.VolumeName)
		  'Volume name should be "VBOXADDITIONS_4."
		  If InStr(1,VolumeNameCDRom,"vboxadditions") > 0 Then
		     WScript.Echo "Win32_CDROMDrive ==> CDRom.VolumeName ==> " & CDRom.VolumeName
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(CDRom.DeviceID) Then
	      DeviceIDCDRom = LCase(CDRom.DeviceID)
		  If InStr(1,DeviceIDCDRom,"cdromvbox") > 0 Then
		     WScript.Echo "Win32_CDROMDrive ==> CDRom.DeviceID ==> " & CDRom.DeviceID
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(CDRom.PNPDeviceID) Then
	      PNPDeviceIDCDRom = LCase(CDRom.PNPDeviceID)
		  If InStr(1,PNPDeviceIDCDRom,"cdromvbox") > 0 Then
		     WScript.Echo "Win32_CDROMDrive ==> CDRom.PNPDeviceID ==> " & CDRom.PNPDeviceID
			 VBoxFound = True
		  End If
	   End If		   
	End If
Next


'WIN32_NetworkClient aka netclient
Set NetClientQ = objX.ExecQuery("SELECT * FROM WIN32_NetworkClient")
For Each NetClient in NetClientQ
    If Not IsNull(NetClient) Then
	   If Not IsNull(NetClient.Description) Then
	      DescNetClient = LCase(NetClient.Description)
		  If DescNetClient = "vboxsf" Then
		     WScript.Echo "WIN32_NetworkClient ==> NetClient.Description ==> " & NetClient.Description
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(NetClient.Manufacturer) Then
	      ManufacturerNetClient = LCase(NetClient.Manufacturer)
		  If ManufacturerNetClient = "oracle corporation" Then
		     WScript.Echo "WIN32_NetworkClient ==> NetClient.Manufacturer ==> " & NetClient.Manufacturer
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(NetClient.Name) Then
	      NameNetClient = LCase(NetClient.Name)
		  If NameNetClient = "virtualbox shared folders" Then
		     WScript.Echo "WIN32_NetworkClient ==> NetClient.Name ==> " & NetClient.Name
			 VBoxFound = True
		  End If
	   End If
	End If
Next

'Win32_ComputerSystemProduct aka csproduct
Set CSProductQ = objX.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")
For Each CSProduct in CSProductQ
    If Not IsNull(CSProduct) Then
	   If Not IsNull(CSProduct.Name) Then
	      NameCSProduct = LCase(CSProduct.Name)
		  If NameCSProduct = "virtualbox" Then
		     WScript.Echo "Win32_ComputerSystemProduct ==> CSProduct.Name ==> " & CSProduct.Name
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(CSProduct.Vendor) Then
	      VendorCSProduct = LCase(CSProduct.Vendor)
		  If VendorCSProduct = "innotek gmbh" Then
		     WScript.Echo "Win32_ComputerSystemProduct ==> CSProduct.Vendor ==> " & CSProduct.Vendor
			 VBoxFound = True
		  End If
	   End If
	End If
Next

'Win32_VideoController
Set VideoControllerQ = objX.ExecQuery("SELECT * FROM Win32_VideoController")
For Each VideoController in VideoControllerQ
    If Not IsNull(VideoController) Then
	   If Not IsNull(VideoController.Name) Then
	      NameVideoController = LCase(VideoController.Name)
		  If NameVideoController = "virtualbox graphics adapter" Then
		     WScript.Echo "Win32_VideoController ==> VideoController.Name ==> " & VideoController.Name
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(VideoController.Description) Then
	      DescVideoController = LCase(VideoController.Description)
		  If DescVideoController = "virtualbox graphics adapter" Then
		     WScript.Echo "Win32_VideoController ==> VideoController.Description ==> " & VideoController.Description
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(VideoController.Caption) Then
	      CaptionVideoController = LCase(VideoController.Caption)
		  If CaptionVideoController = "virtualbox graphics adapter" Then
		     WScript.Echo "Win32_VideoController ==> VideoController.Caption ==> " & VideoController.Caption
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(VideoController.VideoProcessor) Then
	      VideoProcessorVideoController = LCase(VideoController.VideoProcessor)
		  If VideoProcessorVideoController = "vbox" Then
		     WScript.Echo "Win32_VideoController ==> VideoController.VideoProcessor ==> " & VideoController.VideoProcessor
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(VideoController.InstalledDisplayDrivers) Then
	      InstalledDisplayDriversVideoController = LCase(VideoController.InstalledDisplayDrivers)
		  If InstalledDisplayDriversVideoController = "vboxdisp.sys" Then
		     WScript.Echo "Win32_VideoController ==> VideoController.InstalledDisplayDrivers ==> " & VideoController.InstalledDisplayDrivers
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(VideoController.InfSection) Then
	      InfSectionVideoController = LCase(VideoController.InfSection)
		  If InfSectionVideoController = "vboxvideo" Then
		     WScript.Echo "Win32_VideoController ==> VideoController.InfSection ==> " & VideoController.InfSection
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(VideoController.AdapterCompatibility) Then
	      AdapterCompatibilityVideoController = LCase(VideoController.AdapterCompatibility)
		  If AdapterCompatibilityVideoController = "oracle corporation" Then
		     WScript.Echo "Win32_VideoController ==> VideoController.AdapterCompatibility ==> " & VideoController.AdapterCompatibility
			 VBoxFound = True
		  End If
	   End If
	End If
Next


'Win32_PnPEntity
Set PnPEntityQ = objX.ExecQuery("SELECT * FROM Win32_PnPEntity")
For Each PnPEntity in PnPEntityQ
    If Not IsNull(PnPEntity) Then
	   If Not IsNull(PnPEntity.Name) Then
	      NamePnPEntity = LCase(PnPEntity.Name)
		  If NamePnPEntity = "virtualbox device" Or NamePnPEntity = "vbox harddisk" Or NamePnPEntity = "vbox cd-rom" Or NamePnPEntity = "virtualbox graphics adapter" Then
		     WScript.Echo "Win32_PnPEntity ==> PnPEntity.Name ==> " & PnPEntity.Name
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(PnPEntity.Caption) Then
	      CaptionPnPEntity = LCase(PnPEntity.Caption)
		  If CaptionPnPEntity = "virtualbox device" Or CaptionPnPEntity = "vbox harddisk" Or CaptionPnPEntity = "vbox cd-rom" Or CaptionPnPEntity = "virtualbox graphics adapter" Then
		     WScript.Echo "Win32_PnPEntity ==> PnPEntity.Caption ==> " & PnPEntity.Caption
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(PnPEntity.Description) Then
	      DescPnPEntity = LCase(PnPEntity.Description)
		  If DescPnPEntity = "virtualbox device" Or DescPnPEntity = "virtualbox graphics adapter" Then
		     WScript.Echo "Win32_PnPEntity ==> PnPEntity.Description ==> " & PnPEntity.Description
			 VBoxFound = True
		  End If
	   End If
	   'Had to remove .Manufacturer as it detects Host as well
	   'If Not IsNull(PnPEntity.Manufacturer) Then
	      'ManuPnPEntity = LCase(PnPEntity.Manufacturer)
		  'If ManuPnPEntity = "oracle corporation" Then
		     'WScript.Echo "Win32_PnPEntity ==> PnPEntity.Manufacturer ==> " & PnPEntity.Manufacturer
			 'VBoxFound = True
		  'End If
	   'End If
	   If Not IsNull(PnPEntity.Service) Then
	      SrvPnPEntity = LCase(PnPEntity.Service)
		  If SrvPnPEntity = "vboxguest" Or SrvPnPEntity = "vboxvideo" Then
		     WScript.Echo "Win32_PnPEntity ==> PnPEntity.Service ==> " & PnPEntity.Service
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(PnPEntity.DeviceID) Then
	      DeviceIDPnPEntity = LCase(PnPEntity.DeviceID)
		  If InStr(1,DeviceIDPnPEntity,"diskvbox_") > 0 Or InStr(1,DeviceIDPnPEntity,"cdromvbox_") > 0 Then
		     WScript.Echo "Win32_PnPEntity ==> PnPEntity.DeviceID ==> " & PnPEntity.DeviceID
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(PnPEntity.PNPDeviceID) Then
	      PNPDeviceIDPnPEntity = LCase(PnPEntity.PNPDeviceID)
		  If InStr(1,PNPDeviceIDPnPEntity,"diskvbox_") > 0 Or InStr(1,PNPDeviceIDPnPEntity,"cdromvbox_") > 0 Then
		     WScript.Echo "Win32_PnPEntity ==> PnPEntity.PNPDeviceID ==> " & PnPEntity.PNPDeviceID
			 VBoxFound = True
		  End If
	   End If
	End If
Next

'Win32_NetworkConnection aka NetUse
Set NetUseQ = objX.ExecQuery("SELECT * FROM Win32_NetworkConnection")
For Each NetUse in NetUseQ
    If Not IsNull(NetUse) Then
	   If Not IsNull(NetUse.Name) Then
	      NameNetUse = LCase(NetUse.Name)
		  If InStr(1,NameNetUse,"vboxsvr") > 0 Then
		     WScript.Echo "Win32_NetworkConnection ==> NetUse.Name ==> " & NetUse.Name
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(NetUse.Description) Then
	      DescNetUse = LCase(NetUse.Description)
		  If  InStr(1,DescNetUse,"virtualbox shared folders") > 0 Then
		     WScript.Echo "Win32_NetworkConnection ==> NetUse.Description ==> " & NetUse.Description
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(NetUse.ProviderName) Then
	      PrvNameNetUse = LCase(NetUse.ProviderName)
		  If PrvNameNetUse = "virtualbox shared folders" Then
		     WScript.Echo "Win32_NetworkConnection ==> NetUse.ProviderName ==> " & NetUse.ProviderName
			 VBoxFound = True
		  End If
	   End If

	   If Not IsNull(NetUse.RemoteName) Then
	      RemoteNameNetUse = LCase(NetUse.RemoteName)
		  If InStr(1,RemoteNameNetUse,"vboxsvr") > 0 Then
		     WScript.Echo "Win32_NetworkConnection ==> NetUse.RemoteName ==> " & NetUse.RemoteName
			 VBoxFound = True
		  End If
	   End If
	   If Not IsNull(NetUse.RemotePath) Then
	      RemotePathNetUse = LCase(NetUse.RemotePath)
		  If InStr(1,RemotePathNetUse,"vboxsvr") > 0 Then
		     WScript.Echo "Win32_NetworkConnection ==> NetUse.RemotePath ==> " & NetUse.RemotePath
			 VBoxFound = True
		  End If
	   End If
	End If
Next

If VBoxFound = False Then
   WScript.Echo "VirtualBox Was Not Found"
End If