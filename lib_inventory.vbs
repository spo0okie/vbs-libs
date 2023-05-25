'обновлеят запись компьютера в БД, если в качестве compID передано -1, то создает новую
'возвращает ID созданной/обновленной записи
function invUpdateComp(byVal compID, byVal domainID, byVal name, byVal OS, byVal hw, byVal sw, byVal ip, byVal mac)
	datetime = Replace(timeGetUtcTimestamp,"/","-")
	data=_
	"domain_id="&domainID&_
	"&name="&Url.Encode(name)&_
	"&os="&Url.Encode(OS)&_
	"&raw_hw="&Url.Encode(hw)&_
	"&raw_soft="&Url.Encode(sw)&_
	"&raw_version="&Url.Encode(scrVer)&_
	"&ip="&Url.Encode(ip)&_
	"&mac="&Url.Encode(mac)&_
	"&updated_at="&Url.Encode(datetime)

	'сохраняем данные в отдельный файлик. На случай если с сервером связаться не удастся
	writeFile WorkDir & scrName & ".dat", data

	if (compID>-1) then
		res=putXmlData(inventory_apihost & "/web/api/comps/"&compID,data)
	else
		res=postXmlData(inventory_apihost & "/web/api/comps",data)
	end if

	id = getXmlResponseID(res)

	if (id > -1) then
		Msg "Got actual compID " & id & "; Database updated"
	else
		Msg "ERR: failed send data to database: " & vbCrLf &_
			"data=" & data & vbCrLf &_
			"response=" & res
	end if

	invUpdateComp = id	
end function


'возвращает элементы групп лицензий (типов лицензий)
function invGetProductLicGroups(byVal productId, byVal compName, byVal userLogin)
	invGetProductLicGroups=getXmlData(inventory_apihost & "/web/api/lic-groups/search?product_id="&productId&"&user_login="&userLogin&"&comp_name="&compName)
end function


'лицензирован ли продукт на ОС
function invIsProductLicensedOnComp(byVal productId, byVal compName)
	invIsProductLicensedOnComp=countXmlItems(getXmlData(inventory_apihost & "/web/api/lic-groups/search?product_id="&productId&"&comp_name="&compName),"descr")
end function

'лицензирован ли продукт для пользователя
function invIsProductLicensedForUser(byVal productId, byVal userLogin)
	invIsProductLicensedForUser=countXmlItems(getXmlData(inventory_apihost & "/web/api/lic-groups/search?product_id="&productId&"&user_login="&userLogin),"descr")
end function

function invGetMyCompId()
	invGetMyCompId=getXmlResponseID(getXMLData(inventory_apihost&"/web/api/comps/"&computerDomain&"."&computerName))
end function

function invGetMyUserId()
	invGetMyUserId=getXmlResponseID(getXMLData(inventory_apihost&"/web/api/users/view?login="&userName))
end function
'' SIG '' Begin signature block
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' d2AOsj6xIP8E31Gj+Wer6o1AGIfLFeL7eIHm9ypA/1mg
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
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIFFHcm2oRH0n
'' SIG '' Be//Et42zPbRyaCG0BGspaR+qAkF0Iz9MA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAIUkjpzipJehoTkVAWHdufvviyoO4h14
'' SIG '' cFFmKrDjTiOcHZiJTBKsSj07wfCoYe34hUIy6i2p19TL
'' SIG '' LNXP5Xw+qy+C1mAXM1WefCd+BFtG1F7H6xr9ui/m1nCz
'' SIG '' Jy67mj86RwdgL3AbsqQkpSbV0uUZ/A+VlrQKNl5ZkwwZ
'' SIG '' eCS9MfbXYxA5KKd4GZDoAnxK9aH6Dqf3ggLB8snclxZb
'' SIG '' NWnZ56NbWUi1L+UFOUZNQkf+YshxEaFy0E1W1Ue4xYeT
'' SIG '' /R0u8EVzCkHHlEGlhXZiRsyW56uHrPPDAjbGw7Re0WTH
'' SIG '' +KfISRhT1JenEGoYqAD9ig96DW+eY4G4ctOOLZc3ErWX7oY=
'' SIG '' End signature block
