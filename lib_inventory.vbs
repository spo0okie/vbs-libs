option explicit

const inventoryCompIDRegStorage="HKEY_LOCAL_MACHINE\SOFTWARE\Reviakin\Inventory\compID"

dim inventory_user: inventory_user=""
dim inventory_password: inventory_password=""

function fetchCurrentCompId()
	dim savedCompID
	savedCompID=regRead(inventoryCompIDRegStorage)

	if (Len(savedCompID)>0) then
		savedCompID=CLng(savedCompID)
	else
		savedCompId=-1
	end if

	if (savedCompID > 0) then
		msg "INFO: saved CompID found ("&savedCompID&"). Checking in database..."
		compID=invChkCompId(savedCompID)
	else
		compID=-1
	end if


	if (compID > 0) then
		msg "INFO: got CompID=" & compID
	else
		msg "INFO: no valid saved CompID. Searching in database..."
		compID=invGetMyCompId
		msg compID
	end if
	fetchCurrentCOmpId=compID
end function


'обновлеят запись компьютера в БД, если в качестве compID передано -1, то создает новую
'возвращает ID созданной/обновленной записи
function invUpdateComp(byVal name, byVal OS, byVal hw, byVal sw, byVal ip, byVal mac)
	'datetime = Replace(timeGetUtcTimestamp,"/","-")
	dim data : data=_
	"&name="&Url.Encode(name)&_
	"&os="&Url.Encode(OS)&_
	"&raw_hw="&Url.Encode(hw)&_
	"&raw_soft="&Url.Encode(sw)&_
	"&raw_version="&Url.Encode(scrVer)&_
	"&ip="&Url.Encode(ip)&_
	"&mac="&Url.Encode(mac) '&_
	'"&updated_at="&Url.Encode(datetime)

	'сохраняем данные в отдельный файлик. На случай если с сервером связаться не удастся
	writeFile WorkDir & scrName & ".dat", data

	dim res,compID
	compID=fetchCurrentCompID
	
	if (compID>-1) then
		res=putAuthXmlData(inventory_apihost & "/web/api/comps/"&compID, data, inventory_user, inventory_password)
	else
		res=postAuthXmlData(inventory_apihost & "/web/api/comps", data, inventory_user, inventory_password)
	end if

	dim id : id = getXmlResponseID(res)

	if (id > -1) then
		Msg "Got actual compID " & id & "; Database updated"
	else
		Msg "ERR: failed send data to database: " & vbCrLf &_
			"data=" & data & vbCrLf &_
			"response=" & res
	end if

	invUpdateComp = id	
end function

'отправляет данные на сервер
function invSendCompData(data)
	dim res : res=postAuthXmlData(inventory_apihost & "/web/api/comps/push", data, inventory_user, inventory_password)

	dim id : id = getXmlResponseID(res)

	if (id > -1) then
		Msg "Got actual compID " & id & "; Database updated"
	else
		Msg "ERR: failed send data to database: " & vbCrLf &_
			"data=" & data & vbCrLf &_
			"response=" & res
	end if
	invSendCompData = id
end function


'обновлеят запись компьютера в БД, если в качестве compID передано -1, то создает новую
'возвращает ID созданной/обновленной записи
function invPushComp(byVal name, byVal OS, byVal hw, byVal sw, byVal ip, byVal mac)
	'datetime = Replace(timeGetUtcTimestamp,"/","-")
	dim data : data=_
	"&name="&Url.Encode(name)&_
	"&os="&Url.Encode(OS)&_
	"&raw_hw="&Url.Encode(hw)&_
	"&raw_soft="&Url.Encode(sw)&_
	"&raw_version="&Url.Encode(scrVer)&_
	"&ip="&Url.Encode(ip)&_
	"&mac="&Url.Encode(mac) '&_
	'"&updated_at="&Url.Encode(datetime)

	dim savedCompID : savedCompID=regRead(inventoryCompIDRegStorage)

	if (Len(savedCompID)>0) then
		data=data&"&id="&savedCompID
	end if

	'сохраняем данные в отдельный файлик. На случай если с сервером связаться не удастся
	writeFile WorkDir & scrName & ".dat", data

	invPushComp = invSendCompData (data)
end function

function invAuthorizedAccess()
	invAuthorizedAccess=false
	if (Len(inventory_user) > 0 and Len(inventory_password) > 0) then
		invAuthorizedAccess=true
	end if
end function


'возвращает элементы групп лицензий (типов лицензий)
function invGetProductLicGroups(byVal productId, byVal compName, byVal userLogin)
	invGetProductLicGroups=getAuthXmlData(inventory_apihost & "/web/api/lic-groups/search?product_id="&productId&"&user_login="&userLogin&"&comp_name="&compName, inventory_user, inventory_password)
end function

'лицензирован ли продукт на ОС
function invIsProductLicensedOnComp(byVal productId, byVal compName)
	invIsProductLicensedOnComp=countXmlItems(getAuthXmlData(inventory_apihost & "/web/api/lic-groups/search?product_id="&productId&"&comp_name="&compName, inventory_user, inventory_password),"descr")
end function

'лицензирован ли продукт для пользователя
function invIsProductLicensedForUser(byVal productId, byVal userLogin)
	invIsProductLicensedForUser=countXmlItems(getAuthXmlData(inventory_apihost & "/web/api/lic-groups/search?product_id="&productId&"&user_login="&userLogin, inventory_user, inventory_password),"descr")
end function

'возвращает ID компьютера по домену, имени или -1 если не найден
function invGetCompId(byVal domain, byVal name)
	dim compData
	compData=getAuthXMLData(inventory_apihost&"/web/api/comps/search?name="&name&"."&domain, inventory_user, inventory_password)
	invGetCompId=getXmlResponseID(compData)
	debugMsg "Got CompID="&invGetCompId&" from "&inventory_apihost&"/web/api/comps/search?name="&name&"."&domain
end function

'возвращает ID этого компьютера
function invGetMyCompId()
	debugMsg computerDomain & "\" & computerName
	invGetMyCompId=invGetCompId(computerDomain,computerName)
end function

'возвращает ID этого пользователя
function invGetMyUserId()
	invGetMyUserId=getXmlResponseID(getAuthXMLData(inventory_apihost&"/web/api/users/search?login="&userName, inventory_user, inventory_password))
end function


'возвращает ID домена по имени или -1 если домен не найден
function invGetDomainId(byVal domain)
	invGetDomainId=getXmlResponseID(getAuthXMLData(inventory_apihost&"/web/api/domains/"&domain, inventory_user, inventory_password))
	debugMsg "Got DomainID="&invGetDomainId&" from "&inventory_apihost&"/web/api/domains/"&domain	
end function

'возвращает ID компьютера по ID или -1 если не найден 
'для валидации сохраненного ID - проверка что ID есть в БД
function invChkCompId(byVal id)
	invChkCompId=getXmlResponseID(getAuthXMLData(inventory_apihost&"/web/api/comps/"&id, inventory_user, inventory_password))
	debugMsg "Got CompID="&invChkCompId&" from "&inventory_apihost&"/web/api/comps/"&id
end function


'' SIG '' Begin signature block
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' ImRqV1naI8dEGFg5ssAracJfTfGLvvAAaZ38L8fTVwyg
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
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIC6q2P6Lpk4w
'' SIG '' uZiHvdIH/y4ZWxQkhfBiZj8D3/hBpEKLMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAGNLWIpvhI4kHsUpQjY4w5GPzeGmjPsl
'' SIG '' qCkuSFeFab4L4+gqTBlfPABRwlzUGahWkEsiy5VYajki
'' SIG '' D78kjbHKMcnNCoC7IdxLzuRRU0s85C7ilafYawAKG6+f
'' SIG '' Ggveh/41W1CJ7Se0aDjMWk7mFAuU5ke8m2MjsgQEhdaQ
'' SIG '' u+fHf99CPWTevg48h7NF5cUzFeWHDQmh3DWBcKQVliCv
'' SIG '' +b7bFjrmdbec7Pw3WEZIatG3SgZe+lR+gxkcuiljvCVq
'' SIG '' zdZC69tqOM+v0kLOXDywBnUGPdKQuaIR1cmC2Of5+2li
'' SIG '' 5LBTfP6fTQzDoWtSWaUm49sbOtuzt7wixwTzs+izsV7HALs=
'' SIG '' End signature block
