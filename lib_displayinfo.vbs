'Библиотека получения информации о мониторах из реестра
Const DISPLAY_REGKEY="HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY\"

'this function formats the parsed array for display
'this is where the final output is generated
'it is the one you will most likely want to
'customize to suit your needs
Function GetFormattedMonitorInfo(arrParsedMonitorInfo)
	for tmpctr=0 to ubound(arrParsedMonitorInfo)
		tmpResult=split(arrParsedMonitorInfo(tmpctr),"|||")
		if not (tmpResult(1) = "Bad EDID") then 
			if (Len(tmpOutput)>0) then tmpOutput=tmpOutput & ","
			tmpOutput=tmpOutput & "{""Monitor"":{" 
			'tmpOutput=tmpOutput & "EDID_VESAManufacturerID=" & tmpResult(1) & vbcrlf
			tmpOutput=tmpOutput & """DeviceID"":""" & tmpResult(3) & ""","
			tmpOutput=tmpOutput & """ManufactureDate"":""" & tmpResult(2) & ""","
			tmpOutput=tmpOutput & """SerialNumber"":""" & tmpResult(0) & ""","
			tmpOutput=tmpOutput & """ModelName"":""" & tmpResult(4) & ""","
			tmpOutput=tmpOutput & """Version"":""" & tmpResult(5) & ""","
			tmpOutput=tmpOutput & """VESAID"":""" & tmpResult(6) & ""","
			tmpOutput=tmpOutput & """PNPID"":""" & tmpResult(7) & """}}"
		end if
	next
	GetFormattedMonitorInfo=tmpOutput
End Function

'This is the main function. It calls everything else
'in the correct order.
Function GetMonitorInfo()
	debugMsg "Getting all display devices"
	arrAllDisplays=GetAllDisplayDevicesInReg()
	debugMsg "Filtering display devices to monitors"
	arrAllMonitors=GetAllMonitorsFromAllDisplays(arrAllDisplays)
	'debugMsg "Filtering monitors to active monitors"
	'arrActiveMonitors=GetActiveMonitorsFromAllMonitors(arrAllMonitors)
	'у меня на 10ке показывало что нет ни одного активного монитора
	arrActiveMonitors=arrAllMonitors
	if ubound(arrActiveMonitors)=0 and arrActiveMonitors(0)="{ERROR}" then
		debugMsg "No active monitors found"
		strFormattedMonitorInfo=""
	else
		debugMsg "Found active monitors"
		debugMsg "Retrieving EDID for all active monitors"
		arrActiveEDID=GetEDIDFromActiveMonitors(arrActiveMonitors)
		debugMsg "Parsing EDID/Windows data"
		arrParsedMonitorInfo=GetParsedMonitorInfo(arrActiveEDID,arrActiveMonitors)
		debugMsg "Formatting parsed data"
		strFormattedMonitorInfo=GetFormattedMonitorInfo(arrParsedMonitorInfo)
	end if
	debugMsg "Data retrieval completed"
	GetMonitorInfo=strFormattedMonitorInfo
end function








'This function returns an array of all subkeys of the 
'regkey defined by DISPLAY_REGKEY
'(typically this should be "HKLM\SYSTEM\CurrentControlSet\Enum\DISPLAY")
Function GetAllDisplayDevicesInReg()
	dim arrResult()
	redim arrResult(0)
	intArrResultIndex=-1
	arrtmpkeys=RegEnumKeys(DISPLAY_REGKEY)
	if vartype(arrtmpkeys)<>8204 then
		arrResult(0)="{ERROR}"
		GetAllDisplayDevicesInReg=false
		debugMsg "Display=Can't enum subkeys of display regkey"
	else
		debugMsg "Display=Can enum subkeys of display regkey"
		for tmpctr=0 to ubound(arrtmpkeys)
			arrtmpkeys2=RegEnumKeys(DISPLAY_REGKEY & arrtmpkeys(tmpctr))
			for tmpctr2 = 0 to ubound(arrtmpkeys2)
				intArrResultIndex=intArrResultIndex+1
				redim preserve arrResult(intArrResultIndex)
				arrResult(intArrResultIndex)=DISPLAY_REGKEY & arrtmpkeys(tmpctr) & "\" & arrtmpkeys2(tmpctr2)
				debugMsg "Display=" & arrResult(intArrResultIndex)
			next 
		next
	end if
	GetAllDisplayDevicesInReg=arrResult
End Function





'This function is passed an array of regkeys as strings
'and returns an array containing only those that have a
'hardware id value appropriate to a monitor.
Function GetAllMonitorsFromAllDisplays(arrRegKeys)
	dim arrResult()
	redim arrResult(0)
	intArrResultIndex=-1
	for tmpctr=0 to ubound(arrRegKeys)
		if IsDisplayDeviceAMonitor(arrRegKeys(tmpctr)) then
			intArrResultIndex=intArrResultIndex+1
			redim preserve arrResult(intArrResultIndex)
			arrResult(intArrResultIndex)=arrRegKeys(tmpctr)
			debugMsg "Monitor=" & arrResult(intArrResultIndex)
		end if
	next
	if intArrResultIndex=-1 then
		arrResult(0)="{ERROR}"
		debugMsg "Monitor=Unable to locate any monitors"
	end if
	GetAllMonitorsFromAllDisplays=arrResult
End Function


'this function is passed a regsubkey as a string
'and determines if it is a monitor
'returns boolean
Function IsDisplayDeviceAMonitor(strDisplayRegKey)
	'DEBUGMODE = 1
	dim arrtmpResult, strtmpResult
	arrtmpResult=RegGetMultiStringValue(strDisplayRegKey,"HardwareID")
	if (isarray(arrtmpResult)) then
		strtmpResult="|||" & join(arrtmpResult,"|||") & "|||"
	else
		strtmpResult=arrtmpResult
	end if
	
	if instr(lcase(strtmpResult),"|||monitor\")=0 then
		debugMsg "MonitorCheck='" & strDisplayRegKey & "'|||is not a monitor"
		IsDisplayDeviceAMonitor=false
	else
		debugMsg "MonitorCheck='" & strDisplayRegKey & "'|||is a monitor"
		IsDisplayDeviceAMonitor=true
	end if
End Function


'This function is passed an array of regkeys as strings
'and returns an array containing only those that have a
'subkey named "Control"...establishing that they are current.
Function GetActiveMonitorsFromAllMonitors(arrRegKeys)
	dim arrResult()
	redim arrResult(0)
	intArrResultIndex=-1
	for tmpctr=0 to ubound(arrRegKeys)
		if IsMonitorActive(arrRegKeys(tmpctr)) then
			intArrResultIndex=intArrResultIndex+1
			redim preserve arrResult(intArrResultIndex)
			arrResult(intArrResultIndex)=arrRegKeys(tmpctr)
			debugMsg "ActiveMonitor=" & arrResult(intArrResultIndex)
		end if
	next
	if intArrResultIndex=-1 then
		arrResult(0)="{ERROR}"
		debugMsg "ActiveMonitor=Unable to locate any active monitors"
	end if
	GetActiveMonitorsFromAllMonitors=arrResult
End Function

'this function is passed a regsubkey as a string
'and determines if it is an active monitor
'returns boolean
Function IsMonitorActive(strMonitorRegKey)
	arrtmpResult=RegEnumKeys(strMonitorRegKey)
	strtmpResult="|||" & join(arrtmpResult,"|||") & "|||"
	if instr(lcase(strtmpResult),"|||control|||")=0 then
		debugMsg "ActiveMonitorCheck='" & strMonitorRegKey & "'|||is not active"
		IsMonitorActive=true 'false
	else
		debugMsg "ActiveMonitorCheck='" & strMonitorRegKey & "'|||is active"
		IsMonitorActive=true
	end if
End Function

'This function is passed an array of regkeys as strings
'and returns an array containing the corresponding contents
'of the EDID value (in string format) for the "Device Parameters" 
'subkey of the specified key
Function GetEDIDFromActiveMonitors(arrRegKeys)
	dim arrResult()
	redim arrResult(0)
	intArrResultIndex=-1
	for tmpctr=0 to ubound(arrRegKeys)
		strtmpResult=GetEDIDForMonitor(arrRegKeys(tmpctr))
		intArrResultIndex=intArrResultIndex+1
	redim preserve arrResult(intArrResultIndex)
		arrResult(intArrResultIndex)=strtmpResult
		debugMsg "GETEDID=" & arrRegKeys(tmpctr) & "|||EDID,Yes"
	next
	if intArrResultIndex=-1 then
		arrResult(0)="{ERROR}"
		debugMsg "EDID=Unable to retrieve any edid"
	end if
	GetEDIDFromActiveMonitors=arrResult
End Function

'given the regkey of a specific monitor
'this function returns the EDID info
'in string format
Function GetEDIDForMonitor(strMonitorRegKey)
	arrtmpResult=RegGetBinaryValue(strMonitorRegKey & "\Device Parameters","EDID")
	if vartype(arrtmpResult) <> 8204 then
		debugMsg "GetEDID=No EDID Found|||" & strMonitorRegKey
		GetEDIDForMonitor="{ERROR}"
	else
		for each bytevalue in arrtmpResult
			strtmpResult=strtmpResult & chr(bytevalue)
		next
		debugMsg "GetEDID=EDID Found|||" & strMonitorRegKey
		debugMsg "GetEDID_Result=" & GetHexFromString(strtmpResult)
		GetEDIDForMonitor=strtmpResult
	end if
End Function

'passed a given string this function 
'returns comma seperated hex values 
'for each byte
Function GetHexFromString(strText)
	for tmpctr=1 to len(strText)
		tmpresult=tmpresult & right( "0" & hex(asc(mid(strText,tmpctr,1))),2) & ","
	next
	GetHexFromString=left(tmpresult,len(tmpresult)-1)
End Function

'this function should be passed two arrays with the same
'number of elements. array 1 should contain the
'edid information that corresponds to the active monitor
'regkey found in the same element of array 2
'Why not use a 2D array or a dictionary object?.
'I guess I'm just lazy
Function GetParsedMonitorInfo(arrActiveEDID,arrActiveMonitors)
	dim arrResult()
	for tmpctr=0 to ubound(arrActiveEDID)
		strSerial=GetSerialFromEDID(arrActiveEDID(tmpctr))
		strMfg=GetMfgFromEDID(arrActiveEDID(tmpctr))
		strMfgDate=GetMfgDateFromEDID(arrActiveEDID(tmpctr))
		strDev=GetDevFromEDID(arrActiveEDID(tmpctr))
		strModel=GetModelFromEDID(arrActiveEDID(tmpctr))
		strEDIDVer=GetEDIDVerFromEDID(arrActiveEDID(tmpctr))
		strWinVesaID=GetWinVESAIDFromRegKey(arrActiveMonitors(tmpctr))
		strWinPNPID=GetWinPNPFromRegKey(arrActiveMonitors(tmpctr))
		redim preserve arrResult(tmpctr)
		arrResult(tmpctr)=arrResult(tmpctr) & strSerial & "|||"
		arrResult(tmpctr)=arrResult(tmpctr) & strMfg & "|||"
		arrResult(tmpctr)=arrResult(tmpctr) & strMfgDate & "|||"
		arrResult(tmpctr)=arrResult(tmpctr) & strDev & "|||"
		arrResult(tmpctr)=arrResult(tmpctr) & strModel & "|||"
		arrResult(tmpctr)=arrResult(tmpctr) & strEDIDVer & "|||"
		arrResult(tmpctr)=arrResult(tmpctr) & strWinVesaID & "|||"
		arrResult(tmpctr)=arrResult(tmpctr) & strWinPNPID
		debugMsg arrResult(tmpctr)
	next
	GetParsedMonitorInfo=arrResult
End Function

'this is a simple string function to break the VESA monitor ID
'from the registry key
Function GetWinVESAIDFromRegKey(strRegKey)
	if strRegKey="{ERROR}" then
		GetWinVESAIDFromRegKey="Bad Registry Info"
		exit function
	end if
	strtmpResult=right(strRegKey,len(strRegkey)-len(DISPLAY_REGKEY))
	strtmpResult=left(strtmpResult,instr(strtmpResult,"\")-1) 
	GetWinVESAIDFromRegKey=strtmpResult
End Function

'this is a simple string function to break windows PNP device id
'from the registry key
Function GetWinPNPFromRegKey(strRegKey)
	if strRegKey="{ERROR}" then
		GetWinPNPFromRegKey="Bad Registry Info"
		exit function
	end if 
	strtmpResult=right(strRegKey,len(strRegkey)-len(DISPLAY_REGKEY))
	strtmpResult=right(strtmpResult,len(strtmpResult)-instr(strtmpResult,"\"))
	GetWinPNPFromRegKey=strtmpResult
End Function

'utilizes the GetDescriptorBlockFromEDID function
'to retrieve the serial number block
'from the EDID data
Function GetSerialFromEDID(strEDID)
	'a serial number descriptor will start with &H00 00 00 ff
	strTag=chr(&H00) & chr(&H00) & chr(&H00) & chr(&Hff)
	GetSerialFromEDID=GetDescriptorBlockFromEDID(strEDID,strTag)
End Function

'utilizes the GetDescriptorBlockFromEDID function
'to retrieve the model description block
'from the EDID data
Function GetModelFromEDID(strEDID)
	'a model number descriptor will start with &H00 00 00 fc
	strTag=chr(&H00) & chr(&H00) & chr(&H00) & chr(&Hfc)
	GetModelFromEDID=GetDescriptorBlockFromEDID(strEDID,strTag)
End Function

'This function parses a string containing EDID data
'and returns the information contained in one of the
'4 custom "descriptor blocks" providing the data in the
'block is tagged wit a certain prefix
'if no descriptor is tagged with the specified prefix then
'function returns "Not Present in EDID"
'otherwise it returns the data found in the descriptor
'trimmed of its prefix tag and also trimmed of
'leading NULLs (chr(0)) and trailing linefeeds (chr(10))
Function GetDescriptorBlockFromEDID(strEDID,strTag)
	if strEDID="{ERROR}" then
		GetDescriptorBlockFromEDID="Bad EDID"
		exit function
	end if

	'*********************************************************************
	'There are 4 descriptor blocks in edid at offset locations
	'&H36 &H48 &H5a and &H6c each block is 18 bytes long
	'the model and serial numbers are stored in the vesa descriptor
	'blocks in the edid.
	'*********************************************************************
	dim arrDescriptorBlock(3)
	arrDescriptorBlock(0)=mid(strEDID,&H36+1,18)
	arrDescriptorBlock(1)=mid(strEDID,&H48+1,18)
	arrDescriptorBlock(2)=mid(strEDID,&H5a+1,18)
	arrDescriptorBlock(3)=mid(strEDID,&H6c+1,18)

	if instr(arrDescriptorBlock(0),strTag)>0 then
		strFoundBlock=arrDescriptorBlock(0)
	elseif instr(arrDescriptorBlock(1),strTag)>0 then
		strFoundBlock=arrDescriptorBlock(1)
	elseif instr(arrDescriptorBlock(2),strTag)>0 then
		strFoundBlock=arrDescriptorBlock(2)
	elseif instr(arrDescriptorBlock(3),strTag)>0 then
		strFoundBlock=arrDescriptorBlock(3)
	else
		GetDescriptorBlockFromEDID="Not Present in EDID"
		exit function
	end if

	strResult=right(strFoundBlock,14)
	'the data in the descriptor block will either fill the 
	'block completely or be terminated with a linefeed (&h0a)
	if instr(strResult,chr(&H0a))>0 then
		strResult=trim(left(strResult,instr(strResult,chr(&H0a))-1))
	else
		strResult=trim(strResult)
	end if

	'although it is not part of the edid spec (as far as i can tell) it seems as though the
	'information in the descriptor will frequently be preceeded by &H00, this
	'compensates for that
	if left(strResult,1)=chr(0) then strResult=right(strResult,len(strResult)-1)

	GetDescriptorBlockFromEDID=strResult
End Function

'This function parses a string containing EDID data
'and returns the VESA manufacturer ID as a string
'the manufacturer ID is a 3 character identifier
'assigned to device manufacturers by VESA
'I guess that means you're not allowed to make an EDID
'compliant monitor unless you belong to VESA.
Function GetMfgFromEDID(strEDID)
	if strEDID="{ERROR}" then
		GetMfgFromEDID="Bad EDID"
		exit function
	end if

	'the mfg id is 2 bytes starting at EDID offset &H08
	'the id is three characters long. using 5 bits to represent
	'each character. the bits are used so that 1=A 2=B etc..
	'
	'get the data
	tmpEDIDMfg=mid(strEDID,&H08+1,2) 
	Char1=0 : Char2=0 : Char3=0 
	Byte1=asc(left(tmpEDIDMfg,1)) 'get the first half of the string 
	Byte2=asc(right(tmpEDIDMfg,1)) 'get the first half of the string
	'now shift the bits
	'shift the 64 bit to the 16 bit
	if (Byte1 and 64) > 0 then Char1=Char1+16 
	'shift the 32 bit to the 8 bit
	if (Byte1 and 32) > 0 then Char1=Char1+8 
	'etc....
	if (Byte1 and 16) > 0 then Char1=Char1+4 
	if (Byte1 and 8) > 0 then Char1=Char1+2 
	if (Byte1 and 4) > 0 then Char1=Char1+1 

	'the 2nd character uses the 2 bit and the 1 bit of the 1st byte
	if (Byte1 and 2) > 0 then Char2=Char2+16 
	if (Byte1 and 1) > 0 then Char2=Char2+8 
	'and the 128,64 and 32 bits of the 2nd byte
	if (Byte2 and 128) > 0 then Char2=Char2+4 
	if (Byte2 and 64) > 0 then Char2=Char2+2 
	if (Byte2 and 32) > 0 then Char2=Char2+1 

	'the bits for the 3rd character don't need shifting
	'we can use them as they are
	Char3=Char3+(Byte2 and 16) 
	Char3=Char3+(Byte2 and 8) 
	Char3=Char3+(Byte2 and 4) 
	Char3=Char3+(Byte2 and 2) 
	Char3=Char3+(Byte2 and 1) 
	tmpmfg=chr(Char1+64) & chr(Char2+64) & chr(Char3+64)
	GetMfgFromEDID=tmpmfg
End Function

'This function parses a string containing EDID data
'and returns the manufacture date in mm/yyyy format
Function GetMfgDateFromEDID(strEDID)
	if strEDID="{ERROR}" then
		GetMfgDateFromEDID="Bad EDID"
		exit function
	end if

	'the week of manufacture is stored at EDID offset &H10
	tmpmfgweek=asc(mid(strEDID,&H10+1,1))

	'the year of manufacture is stored at EDID offset &H11
	'and is the current year -1990
	tmpmfgyear=(asc(mid(strEDID,&H11+1,1)))+1990

	'store it in month/year format 
	tmpmdt=month(dateadd("ww",tmpmfgweek,datevalue("1/1/" & tmpmfgyear))) & "/" & tmpmfgyear
	GetMfgDateFromEDID=tmpmdt
End Function

'This function parses a string containing EDID data
'and returns the device ID as a string
Function GetDevFromEDID(strEDID)
	if strEDID="{ERROR}" then
		GetDevFromEDID="Bad EDID"
		exit function
	end if
	'the device id is 2bytes starting at EDID offset &H0a
	'the bytes are in reverse order.
	'this code is not text. it is just a 2 byte code assigned
	'by the manufacturer. they should be unique to a model
	tmpEDIDDev1=hex(asc(mid(strEDID,&H0a+1,1)))
	tmpEDIDDev2=hex(asc(mid(strEDID,&H0b+1,1)))
	if len(tmpEDIDDev1)=1 then tmpEDIDDev1="0" & tmpEDIDDev1
	if len(tmpEDIDDev2)=1 then tmpEDIDDev2="0" & tmpEDIDDev2
	tmpdev=tmpEDIDDev2 & tmpEDIDDev1
	GetDevFromEDID=tmpdev
End Function

'This function parses a string containing EDID data
'and returns the EDID version number as a string
'I should probably do this first and then not return any other data
'if the edid version exceeds 1.3 since most if this code probably
'won't work right if they change the spec drastically enough (which they probably
'won't do for backward compatability reasons thus negating my need to check and
'making this comment somewhat redundant)
Function GetEDIDVerFromEDID(strEDID)
	if strEDID="{ERROR}" then
		GetEDIDVerFromEDID="Bad EDID"
		exit function
	end if

	'the version is at EDID offset &H12
	tmpEDIDMajorVer=asc(mid(strEDID,&H12+1,1))

	'the revision level is at EDID offset &H13
	tmpEDIDRev=asc(mid(strEDID,&H13+1,1))

	tmpver=chr(48+tmpEDIDMajorVer) & "." & chr(48+tmpEDIDRev)
	GetEDIDVerFromEDID=tmpver
End Function



'' SIG '' Begin signature block
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' erRNYrSZSFOayAfZYJFcFruEkdR6j37IXsKndYLaVeGg
'' SIG '' ggWcMIIFmDCCA4CgAwIBAgIBAzANBgkqhkiG9w0BAQsF
'' SIG '' ADBtMQswCQYDVQQGEwJSVTENMAsGA1UECAwEVXJhbDEU
'' SIG '' MBIGA1UEBwwLQ2hlbHlhYmluc2sxETAPBgNVBAoMCFJl
'' SIG '' dmlha2luMQswCQYDVQQLDAJJVDEZMBcGA1UEAwwQcmV2
'' SIG '' aWFraW4tcm9vdC1DQTAeFw0yMzA1MjUxNTM3MDBaFw0y
'' SIG '' NDA2MDMxNTM3MDBaMGMxCzAJBgNVBAYTAlJVMQ0wCwYD
'' SIG '' VQQIDARVcmFsMQ0wCwYDVQQHDARDaGVsMREwDwYDVQQK
'' SIG '' DAhSZXZpYWtpbjELMAkGA1UECwwCSVQxFjAUBgNVBAMM
'' SIG '' DXJldmlha2luLWNvZGUwggEiMA0GCSqGSIb3DQEBAQUA
'' SIG '' A4IBDwAwggEKAoIBAQCtsuYd7CVRsLwbN6ybLrnCr72O
'' SIG '' nqGhfdASM37B9yC8+b5nnbw6EqDEN2IHpy32wOoThAlg
'' SIG '' zPna/D5/VX/TYuLR/1vjW+vRQPKbJi8m97BMr8PemMWl
'' SIG '' w6mjl9x4qW0x4irIwXra/Z4R34BgrY8ZACZRah0riiWY
'' SIG '' GXPvCw3ZjNYMXRJF4rVKJ6c/PNg1bNlML1Q8oHcy3MPC
'' SIG '' CVCHF/Qf3Bl/l76GKJhylViC5/ZiX34LfzCopdK1xnnY
'' SIG '' 45cP1c83pQH2IE3ucjGMwzWDYCwTNAeYi69aaK40fGHC
'' SIG '' Z9EJg6sS1RnEyCpp+Sj23T/GOJyTxM4kaiPmlMDZoCAq
'' SIG '' UndLk6HVAgMBAAGjggFLMIIBRzAJBgNVHRMEAjAAMBEG
'' SIG '' CWCGSAGG+EIBAQQEAwIFoDAzBglghkgBhvhCAQ0EJhYk
'' SIG '' T3BlblNTTCBHZW5lcmF0ZWQgQ2xpZW50IENlcnRpZmlj
'' SIG '' YXRlMB0GA1UdDgQWBBSXtltT7BkMs4W7USOsFdk+mc0S
'' SIG '' HjAfBgNVHSMEGDAWgBSNQkTnQD4Z5d3UogsBh0kUyrwl
'' SIG '' pzAOBgNVHQ8BAf8EBAMCBeAwJwYDVR0lBCAwHgYIKwYB
'' SIG '' BQUHAwIGCCsGAQUFBwMEBggrBgEFBQcDAzA4BgNVHR8E
'' SIG '' MTAvMC2gK6AphidodHRwOi8vcGtpLnJldmlha2luLm5l
'' SIG '' dC9jcmwvcm9vdC1jYS5jcmwwPwYIKwYBBQUHAQEEMzAx
'' SIG '' MC8GCCsGAQUFBzAChiNodHRwOi8vcGtpLnJldmlha2lu
'' SIG '' L25ldC9yb290LWNhLmNydDANBgkqhkiG9w0BAQsFAAOC
'' SIG '' AgEAix6Hc2aULCO6RiT4W5PIiB9zQgA4BGT3W5YdSttn
'' SIG '' gGhnmWDEfT2bhB/ZnRLkrtZeL/sYDj94FIfKZMvFTsNN
'' SIG '' CUeDNiV9hQyJrsrI9Gq3nkgcnCOGc/9mqqL7ItS33s1M
'' SIG '' ltSXVA7sLhoQ65yPrP70kd3681COUsCYOq7hroIR3Th4
'' SIG '' L8INGLvUR+Xll1sunIHrnuiTD/GZFNemDec0f3n8mNKp
'' SIG '' 5KiWuYlNYv0Zg//rTvCZfk2Y74Mk/2lCeABVKcQoJai+
'' SIG '' XiSN0mq1b6RlFmfbiuzU3iudZ3SKHKEd3reGBXZxD7b1
'' SIG '' QubveA17QKbgzwjT6DX9ISFjbIOuB9HUo3Bl7VLZ4DyH
'' SIG '' 2mt0z+UC1zpE9DLFzoawf4f5/KN6mixGX9Q7tSQQCOKo
'' SIG '' Jiyk7Y+0aLXhK7RmJdDK3vIieJkXSx0ip1SXdRYgr0sQ
'' SIG '' VsNq2D2SYJ0A1r2wWJ4sNuiHnDuxWuxLsAdC0rZTlKis
'' SIG '' 21i4uOIr3BCj2MFdTTdkeX5xB979r/8MLBdrDlzoVxMz
'' SIG '' tEWwXdNlqiCQosIMVq44bJF1zjFPD6pYk0JgEF9y8wTd
'' SIG '' G2LyGFjTqJYyCrKrWFkQa8GX6pazj4EarEpNjdVC6IXJ
'' SIG '' YRa4vRqUEWfS9WeTGlIR9hJyqtHKAc9N82lwrhTlPhh+
'' SIG '' lkL15ZPRXnnd5aICNgQpndNfyBIxggIbMIICFwIBATBy
'' SIG '' MG0xCzAJBgNVBAYTAlJVMQ0wCwYDVQQIDARVcmFsMRQw
'' SIG '' EgYDVQQHDAtDaGVseWFiaW5zazERMA8GA1UECgwIUmV2
'' SIG '' aWFraW4xCzAJBgNVBAsMAklUMRkwFwYDVQQDDBByZXZp
'' SIG '' YWtpbi1yb290LUNBAgEDMA0GCWCGSAFlAwQCAQUAoHww
'' SIG '' EAYKKwYBBAGCNwIBDDECMAAwGQYJKoZIhvcNAQkDMQwG
'' SIG '' CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisG
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIO06IS7kurTt
'' SIG '' bpRUB4cB4QavrFtrxstUp398VEFwxY+7MA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBADJ6QShVlP4y7WxwYzrBfpQHf/jbxvPv
'' SIG '' veIzn7V0GlgtcCAcFiOZuAFPs0B9CjQRtb0x2GlwYGUR
'' SIG '' A9ExAYSoa+GvFlHXQr5WzluhobjV5fMWKrrvpjO7ZQL8
'' SIG '' MOFsCEkpNrkE8w5i/sCDz+e+5r9nrGdWIieDpMbqpCBF
'' SIG '' wkBcoJVv+xOag3wzTConm9HSN4ioosp4vb/EwA7DOt/w
'' SIG '' l6JkM2bXdGPutLvUEyz0e2TDTsfb9yW0aG8i7nafJbdr
'' SIG '' UQFXGSRHDDAJ8h6yhmSndWMb3kNVdl4RAyCg74Xe4lrp
'' SIG '' IiL+QnMq8hXz4zpDg3zCoP/jeVdQC+9ZV4QHsYlKIvFNB9c=
'' SIG '' End signature block
