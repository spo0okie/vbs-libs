'���������� HTTP REST ��������

'Set xmlHTTP = CreateObject("Microsoft.XMLHTTP")
Dim  xmlHTTP: Set xmlHTTP = CreateObject("Msxml2.ServerXMLHTTP")
'SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056 : Ignore All certificate errors.
xmlHTTP.SetOption 2,13056 

'��������� ������ POST �������� //����� ��� ����� ������� � ��
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

'��������� ������ PUT �������� //����� ��� ���������� ������� � ��
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

'�������� ������ GET ��������
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

'�������� ������ GET ��������
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

'�������� �������� ���� ID �� XML ������ ������� ��� -1 ���� ID �� ������
function getXmlResponseID(byVal Response)
	getXmlResponseID=-1

	dim doc: Set doc = CreateObject("MSXML2.DOMDocument") 
	doc.loadXML(Response)
	Set nodes = doc.getElementsByTagName("id")
	for each node in nodes
		getXmlResponseID=node.text
	next
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