'Библиотека HTTP REST запросов
option explicit

'Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")
Dim  xmlHTTP: Set xmlHTTP = CreateObject("Msxml2.ServerXMLHTTP")
'SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056 : Ignore All certificate errors.
xmlHTTP.SetOption 2,13056 

'отправить данные POST запросом //нужно для новых записей в БД
function postXmlData(byVal url, byVal data)
	debugMsg "HTTP POST-ing " & url

	xmlHTTP.Open "POST", url, false
   	xmlHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    	xmlHTTP.setRequestHeader "Content-Length", CStr(Len(data))
	xmlHTTP.SetRequestHeader "accept","application/xml"
    	xmlHTTP.Send CStr(data)

	Do While xmlHTTP.readystate <> 4: WScript.Sleep 200: Loop
    	postXmlData = xmlHTTP.responseText 
End Function

'отправить данные PUT запросом //нужно для обновления записей в БД
function putXmlData(byVal url, byVal data)
	debugMsg "HTTP PUT-ing " & url

	xmlHTTP.Open "PUT", url, false
   	xmlHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    	xmlHTTP.setRequestHeader "Content-Length", CStr(Len(data))
	xmlHTTP.SetRequestHeader "accept","application/xml"
   	xmlHTTP.Send CStr(data)

	Do While xmlHTTP.readystate <> 4: WScript.Sleep 200: Loop
    	putXmlData = xmlHTTP.responseText 
End Function

'получить данные GET запросом
function getXmlData(byVal url)
	debugMsg "HTTP GET-ting data from " & url & "..."

	xmlHTTP.Open "GET", url , false
	xmlHTTP.SetRequestHeader "accept","application/xml"
	xmlHTTP.Send

	if (xmlHTTP.status=200) then
		debugMsg "HTTP Got data sucessfully"
		getXMLData = xmlHTTP.responseText
	else
		Msg "HTTP GET-ting error: status " & xmlHTTP.status & "(" & xmlHTTP.responseText & ")"
		getXMLData = "error"
	end if
End function

'получить данные GET запросом
function getJsonData(byVal url)
	debugMsg "HTTP GET-ting data from " & url & "..."

	xmlHTTP.Open "GET", url , false
	xmlHTTP.SetRequestHeader "accept","application/json"
	xmlHTTP.Send

	if (xmlHTTP.status=200) then
		debugMsg "HTTP Got data sucessfully"
		getJsonData = xmlHTTP.responseText
	else
		Msg "HTTP GET-ting error: status " & xmlHTTP.status & "(" & xmlHTTP.responseText & ")"
		getJsonData = "error"
	end if
End function

'получает значение поля ID из XML ответа сервера или -1 если ID не найден
function getXmlResponseID(byVal Response)
	dim doc: Set doc = CreateObject("MSXML2.DOMDocument") 
	doc.loadXML(Response)
	dim nodes, node
	Set nodes = doc.getElementsByTagName("id")
	for each node in nodes
		getXmlResponseID=node.text
	next
	if (Len(getXmlResponseID)) then
		getXmlResponseID=CInt(getXmlResponseID)
	else
		getXmlResponseID=-1
	end if
end function

function countXmlItems(byVal data, byVal token)
	dim xml,nodes,node,count
	set xml = CreateObject("MSXML2.DOMDocument")
	count=0
	if (data <> "error" ) then
		xml.loadXML(data)
		set nodes = xml.getElementsByTagName(token)
		for each node in nodes
			msg "got item " & node.text
			count=count+1
		next
	end if
	countXmlItems=count
end function

function getXmlItem(byVal data, byVal token)
	getXmlItem=""
	dim xml,nodes,node,count
	set xml = CreateObject("MSXML2.DOMDocument")
	if (data <> "error" ) then
		xml.loadXML(data)
		set nodes = xml.getElementsByTagName(token)
		for each node in nodes
			getXmlItem=node.text
			exit function
		next
	end if
end function
'' SIG '' Begin signature block
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' yqArAAcmvmXNt5AumMR4pfBNTXvqgc2MA2pS9PNp7W+g
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
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIEMd3Fs+dT5U
'' SIG '' eKOcRfRsksqXRu36ngSVHdOFKtXIASDSMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAJfJxjkC3naSBM+92RnXLYdCWsyj8w4H
'' SIG '' V2bRr/fbgPB5UMGEdwNBcEZKa7YPA5MRq/kZ7XtE4Uxl
'' SIG '' /7P9KBtHVe1GV1xZz8abMoL4x7s+gVH/sMbmTMvpt6mI
'' SIG '' 8j5r33ueyXDYMs2YtZN7SVIgE6sRkWWHJQ2PVrvhaupA
'' SIG '' xJqEFfttW2BHkNfUrK6qTy4bGVFO0L+Yr5QwTM5Fedh+
'' SIG '' UXS4El3iJ4QgtKpbc/ft0gnMfRAQ3n6cd3IwFCNe6d6m
'' SIG '' huqL+DgZFtaaD5fMPNjScu8/ymwg3ttVWVkposElXDJj
'' SIG '' uIr73An1wsxNCMulPDjiO1U/3uVuQpTggoUuvmx4lHj2mNc=
'' SIG '' End signature block
