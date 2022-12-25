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