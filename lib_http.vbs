'Библиотека HTTP REST запросов

'Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")
Dim  xmlHTTP: Set xmlHTTP = CreateObject("Msxml2.ServerXMLHTTP")
'SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056 : Ignore All certificate errors.
xmlHTTP.SetOption 2,13056 

'отправить данные POST запросом //нужно для новых записей в БД
function postXmlData(byVal url, byVal data)
	msg "HTTP POST-ing " & url

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
	msg "HTTP PUT-ing " & url

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
	Msg "HTTP GET-ting data from " & url & "..."

	xmlHTTP.Open "GET", url , false
	xmlHTTP.SetRequestHeader "accept","application/xml"
	xmlHTTP.Send

	if (xmlHTTP.status=200) then
		Msg "HTTP Got data sucessfully"
		getXMLData = xmlHTTP.responseText
	else
		Msg "HTTP GET-ting error: status " & xmlHTTP.status & "(" & xmlHTTP.responseText & ")"
		getXMLData = "error"
	end if
End function

'получает значение поля ID из XML ответа сервера или -1 если ID не найден
function getXmlResponseID(byVal Response)
	getXmlResponseID=-1

	dim doc: Set doc = CreateObject("MSXML2.DOMDocument") 
	doc.loadXML(Response)
	Set nodes = doc.getElementsByTagName("id")
	for each node in nodes
		getXmlResponseID=node.text
	next
end function
