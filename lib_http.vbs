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
