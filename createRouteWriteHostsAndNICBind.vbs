'Vars for setting the route and reading and writing the host file
Dim oWMI, Instances, Instance, Adapter, AdapterConfig, wmiObject
Dim pathToHostFile
Dim objFileSystem, objOutputFile
Dim inputData, strData
Dim ssbHp2Found, ssbHp2String, ssbHp2Ip
Dim svnSsbDisoDeFound, svnSsbDisoDeString, svnSsbDisoDeIp
Dim ssbsHp2SsbDisoDeFound, ssbsHp2SsbDisoDeString, ssbsHp2SsbDisoDeIp
Dim nexusSsbDisoDeFound, nexusSsbDisoDeString, nexusSsbDisoDeIp
Dim jupiterFound, jupiterString, jupitereIp 
Dim appsProdynaComFound,appsProdynaComString, appsProdynaComIp
Const OPEN_FILE_FOR_APPENDING = 8
Const OPEN_FILE_FOR_READING = 1
Const HKEY_LOCAL_MACHINE = &H80000002
Dim counter 
Dim regContent 
Dim bertelsmannDNSDomain

' registry location for nic binding order
strKeyPath = "SYSTEM\Currentcontrolset\Services\TCPIP\Linkage"
strComputer = "."
'tmp var for saving original registry
originalRegistry = Array("")
'new registry
newRegistry = Array("")
Set regContent = CreateObject("Scripting.Dictionary")

'for ip lookup dnsdomain 
bertelsmannDNSDomain = "bertelsmann.de"

'Settings for the hosts file
ssbHp2Found = 1
ssbHp2String = "ssbs-hp2"
ssbHp2Ip = "10.2.137.65"

svnSsbDisoDeFound = 1
svnSsbDisoDeString = "svn.ssb-diso.de"
svnSsbDisoDeIp = "10.2.137.65"

ssbsHp2SsbDisoDeFound = 1
ssbsHp2SsbDisoDeString = "svn.ssb-diso.de"
ssbsHp2SsbDisoDeIp = "10.2.137.65"

nexusSsbDisoDeFound = 1
nexusSsbDisoDeString = "nexus.ssb-diso.de"
nexusSsbDisoDeIp = "10.2.137.65"

jupiterFound = 1
jupiterString = "jupiter"
jupitereIp = "10.2.137.8"

appsProdynaComFound = 1
appsProdynaComString = "apps.prodyna.com"
appsProdynaComIp = "192.168.200.41"

Dim tmpId 

Set oReg=GetObject( _
	"winmgmts:{impersonationLevel=impersonate}!\\" & _
	strComputer & "\root\default:StdRegProv")
bindRegValue = "bind"

'Get base WMI object, "." means computer name (local)
Set oWMI = GetObject("WINMGMTS:\\.\ROOT\cimv2")

'Get instances of Win32_NetworkAdapterSetting 
Set Instances = oWMI.InstancesOf("Win32_NetworkAdapterSetting")

'Set the path to the host file
Set WshShell = CreateObject("WScript.shell") 
'System root 
SysRoot = WshShell.ExpandEnvironmentStrings("%SystemRoot%") 
'File Path to hosts
pathToHostFile = SysRoot & "\system32\drivers\etc\hosts" 

'Enumerate instances  
For Each Instance In Instances 
	'Get adapter config and adapter
	Set AdapterConfig = oWMI.Get(Instance.Setting)
	Set Adapter = oWMI.Get(Instance.Element)
	if AdapterConfig.DNSDomain = bertelsmannDNSDomain Then
		'WScript.Echo "Der Adaper ist die: " + AdapterConfig.DNSDomain + " Name: " + Adapter.Name + " IPAddress: " + AdapterConfig.IPAddress(0) + "IndexId: " + Adapter.DeviceID
		oReg.GetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,_
			bindRegValue,originalRegistry
		For Each strValue In originalRegistry
			if isStringAtEnd(strValue, AdapterConfig.SettingID) = 0 then
				tmpId = strValue
			else
				regContent.Add counter, strValue
				counter = counter + 1
			end if
		Next
		regContent.Add counter, tmpId
		
		ReDim newRegistry(regContent.Count-1)
		Dim items
		items = regContent.Items
		For i = 0 To regContent.Count -1 ' Iterate the array.
			newRegistry(i) = items(i)
		Next
		Set WshShell = WScript.CreateObject("WScript.Shell") 
		oReg.SetMultiStringValue HKEY_LOCAL_MACHINE,strKeyPath,_
			bindRegValue, newRegistry
		WshShell.Run "%comspec% /c route delete 10.2.0.0 &  " & _
			"route add 10.2.0.0 mask 255.255.0.0 10.3.0.1 & " & _
			"route delete 192.168.200.42 & " & _
			"route add 192.168.200.42 mask 255.255.255.255 10.3.0.1 & " & _
			"route delete 192.168.200.41 & " & _
			"route add 192.168.200.41 mask 255.255.255.255 10.3.0.1 & " & _
			"route delete 10.2.137.0 & " & _
			"route add 10.2.137.0 mask 255.255.255.0 " + AdapterConfig.IPAddress(0) + " & " & _
			"netsh interface ip set interface interface=" & Adapter.InterfaceIndex  & " metric=40"
	End if  
Next 


Set objFileSystem = CreateObject("Scripting.fileSystemObject")
Set objOutputFile = objFileSystem.OpenTextFile(pathToHostFile, _
  OPEN_FILE_FOR_READING)
  
 ' read everything in an array out of hosts file
inputData = Split(objOutputFile.ReadAll, vbNewline)

' search for existing entries 
For each strData In inputData
	if isStringAtBeginning(strData, "#") <> 0 and Len(strData) > 0 then
		if isStringAtEnd(strData, ssbHp2String) = 0 and isStringAtBeginning(strData, ssbHp2Ip) = 0 then
			ssbHp2Found = 0
		end if
		if isStringAtEnd(strData, svnSsbDisoDeString) = 0 and isStringAtBeginning(strData, svnSsbDisoDeIp) = 0 then
			svnSsbDisoDeFound = 0
		end if
		if isStringAtEnd(strData, ssbsHp2SsbDisoDeString) = 0 and isStringAtBeginning(strData, ssbsHp2SsbDisoDe2Ip) = 0 then
			ssbsHp2SsbDisoDeFound = 0
		end if
		if isStringAtEnd(strData, nexusSsbDisoDeString) = 0 and isStringAtBeginning(strData, nexusSsbDisoDeIp) = 0 then
			nexusSsbDisoDeFound = 0
		end if
		if isStringAtEnd(strData, jupiterString) = 0 and isStringAtBeginning(strData, jupitereIp) = 0 then
			jupiterFound = 0
		end if
		if isStringAtEnd(strData, appsProdynaComString) = 0 and isStringAtBeginning(strData, appsProdynaComIp) = 0 then
			appsProdynaComFound = 0
		end if
	end if
Next

'open hosts file for adding missing enteries
Set objOutputFile = objFileSystem.OpenTextFile(pathToHostFile, _
 OPEN_FILE_FOR_APPENDING)

objOutputFile.WriteLine("")
if ssbHp2Found = 1 then
	objOutputFile.WriteLine(ssbHp2Ip & " " & ssbHp2String)
end if
if svnSsbDisoDeFound = 1 then
	objOutputFile.WriteLine(svnSsbDisoDeIp & " " & svnSsbDisoDeString)
end if
if ssbsHp2SsbDisoDeFound = 1 then
	objOutputFile.WriteLine(ssbsHp2SsbDisoDeIp & " " & ssbsHp2SsbDisoDeString)
end if
if nexusSsbDisoDeFound = 1 then
	objOutputFile.WriteLine(nexusSsbDisoDeIp & " " & nexusSsbDisoDeString)
end if
if jupiterFound = 1 then
	objOutputFile.WriteLine(jupitereIp & " " & jupiterString)
end if
if appsProdynaComFound = 1 then
	objOutputFile.WriteLine(appsProdynaComIp & " " & appsProdynaComString)
end if
objOutputFile.Close

Set objFileSystem = Nothing

WScript.Quit(0)


'Helper functions .....
function isStringAtEnd(Byval totalString, byval searchString)

	totalString = LTrim(totalString)
	totalString = RTrim(totalString)
	searchString = LTrim(searchString)
	searchString = RTrim(searchString)
	foundEndString = Right(totalString, Len(searchString)) 
	
	isStringAtEnd = StrComp(searchString, foundEndString)
End function

function isStringAtBeginning(Byval totalString, byval searchString)

	totalString = LTrim(totalString)
	totalString = RTrim(totalString)
	searchString = LTrim(searchString)
	searchString = RTrim(searchString)
	foundEndString = Left(totalString, Len(searchString)) 
	
	isStringAtBeginning = StrComp(searchString, foundEndString)
End function