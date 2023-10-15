'Библиотека HTTP REST запросов
option explicit

'Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")
Dim  xmlHTTP: Set xmlHTTP = CreateObject("Msxml2.ServerXMLHTTP")
'SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056 : Ignore All certificate errors.
xmlHTTP.SetOption 2,13056 

function getParamFromData(ByVal data, byVal param, byVal default)
	getParamFromData=default
	dim strings,str,tokens : strings = split(data,"&")
	for each str in strings
    	tokens=split(str,"=")
		'debugMsg "Testing " & str
		if (param = tokens(0)) then
			debugMsg "Got " & tokens(0) & " => " & tokens(1)
			getParamFromData=tokens(1)
		end if
	next
end function

function sendXmlData(byVal mode, byVal url, byVal data, byVal user, byVal password)
	debugMsg "HTTP " & mode & "-ing " & url

	xmlHTTP.Open mode, url, false
	xmlHTTP.SetRequestHeader "accept","application/xml"

	if (Len(user)>0 and Len(password)>0) then
		xmlHTTP.setRequestHeader "Authorization", "Basic " & Base64Encode(user & ":" & password)
	end if

    if (mode="PUT" or mode="POST" or mode="PATCH") then
		xmlHTTP.setRequestHeader "Content-Length", CStr(Len(data))
		xmlHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xmlHTTP.Send CStr(data)
	else
		xmlHTTP.Send
	end if

	Do While xmlHTTP.readystate <> 4: WScript.Sleep 200: Loop

	if (xmlHTTP.status>=200 and xmlHTTP.status<300) then
		debugMsg "HTTP Got data sucessfully"
		sendXmlData = xmlHTTP.responseText
	else
		Msg "HTTP " & mode & "-ting error: status " & xmlHTTP.status & "(" & xmlHTTP.responseText & ")"
		sendXmlData = "error"
	end if
	
end Function


'отправить данные POST запросом //нужно для новых записей в БД
function postXmlData(byVal url, byVal data)
	postXmlData = sendXmlData ("POST", url, data, "", "")
End Function

'отправить данные PUT запросом //нужно для обновления записей в БД
function putXmlData(byVal url, byVal data)
	putXmlData = sendXmlData ("PUT", url, data, "", "")
End Function

'получить данные GET запросом
function getXmlData(byVal url)
	getXmlData = sendXmlData ("GET", url, "", "", "")
End function

'отправить данные POST запросом //нужно для новых записей в БД
function postAuthXmlData(byVal url, byVal data, byVal user, byVal password)
	postAuthXmlData = sendXmlData ("POST", url, data, user, password)
End Function

'отправить данные PUT запросом //нужно для обновления записей в БД
function putAuthXmlData(byVal url, byVal data, byVal user, byVal password)
	putAuthXmlData = sendXmlData ("PUT", url, data, user, password)
End Function

'получить данные GET запросом
function getAuthXmlData(byVal url, byVal user, byVal password)
	getAuthXmlData = sendXmlData ("GET", url, "", user, password)
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
		getXmlResponseID=CLng(getXmlResponseID)
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


Function Base64Encode(sText)
    Dim oXML, oNode

    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue =Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function
'' SIG '' Begin signature block
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' tqxCOXyTHjGJE6mPpIuYkQwP5VpPBGE2o1QzZcjE0cqg
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
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEII7m52fEpvz3
'' SIG '' I+TDSDkT7KVYEA+oVRR8kfqQnotxuwNCMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAKlQ2OE/HBIQHuYwhEoXN92d3zS43rKk
'' SIG '' nxWp3di4T7lGZ4Z8Vh26B780a+cVB7l6f/1A3Q1+A4I6
'' SIG '' Yyba6xMuiKtdBhEhbqw5ZuaAk8BHSZWvF8b1mYNNSlon
'' SIG '' sqMPYUea/RmW3G2iogHfLMX5ygrQn7O2mEpmza872BSn
'' SIG '' oWysC+lwSvhwICHUooYHrtOe5TsNJchm1rfk3eqw55Ld
'' SIG '' /z1a63ryznwo5JMx1XYr+rek3faLuigpTVQ8vrqLaOx2
'' SIG '' CPlyTBhAgA2qjAcDvqm3WnofmtTVwzqPxAiOuW+TbysR
'' SIG '' srsGqx+kjh+zbeR5Ao4oN1GRdaptI5klaQzWGZzo35pbbpQ=
'' SIG '' End signature block
