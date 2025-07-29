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
	dim count
	count=0
	count=count+countXmlItems(getAuthXmlData(inventory_apihost & "/web/api/lic-groups/search?product_id="&productId&"&comp_name="&compName, inventory_user, inventory_password),"descr")
	count=count+countXmlItems(getAuthXmlData(inventory_apihost & "/web/api/lic-items/search?product_id="&productId&"&comp_name="&compName, inventory_user, inventory_password),"descr")
	count=count+countXmlItems(getAuthXmlData(inventory_apihost & "/web/api/lic-keys/search?product_id="&productId&"&comp_name="&compName, inventory_user, inventory_password),"descr")
	invIsProductLicensedOnComp=count
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


