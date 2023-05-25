'библиотека с функциями определения принадлежности к группе

'Надо пояснить что собственно тут происходит, т.к. я сначала думал что все понятно, а потом оказалось что не все так очевидно

'это глобальная переменная список групп пользователей/компьютеров в формате
'COMP1\group2
'COMP1\group2
'COMP1\		- наличие такой записи означает, что для этого объекта все группы уже загружены в словарь
'USER1\group1
'USER1\group2
'USER1\		- наличие такой записи означает, что для этого объекта все группы уже загружены в словарь
'он нужен для того чтобы сократить количество запросов к ЛДАП, т.к. для полного списка групп приходится долго
'обходить все дерево вложенных групп. И дапы избежать повторной подобной операции при проверке другой группы
'у польщзователя - мы сразу все кэшируем
Option explicit
Dim objGroupList: Set objGroupList = CreateObject("Scripting.Dictionary")
objGroupList.CompareMode = vbTextCompare 'это означает про при поиске в словаре регистр не будет иметь значения

Dim objSysInfo : Set objSysInfo = CreateObject("ADSystemInfo")
'Сразу проинициализируем все переменные, раз либу подключили - значит пригодится
' Escape any forward slash characters, "/", with the backslash
' escape character. All other characters that should be escaped are.
on error resume next
Dim strUserDN : strUserDN = Replace(objSysInfo.userName, "/", "\/")
Dim strComputerDN : strComputerDN = Replace(objSysInfo.computerName, "/", "\/")
on error goto 0

' Bind to the user and computer objects with the LDAP provider.
Dim objUser : if (Len(strUserDN)) then Set objUser = getADObject (strUserDN)
Dim objComputer : if (Len(strComputerDN)) then Set objComputer = getADObject (strComputerDN)

'Получаем объект из АД сначала через ГК, если не выщло то через ЛДАП
Function getADObject(ByVal strObjectDN)
	'MsgBox "GetAdObject: " & strObjectDN
	'Добавляем обратные слэши для экранирования
	strObjectDN = Replace(strObjectDN, "/", "\/")
	' Используем пути через GC:// что означает Глобальный Каталог 
	' (изначально было LDAP:// и жутко тормозило на RODC)
	' Кое где выскакивали ошибки "К указанному домену невозможно подключиться"
	' кто бы знал почему, но предположим, что проблема в GC
	On Error Resume Next
		Set getADObject = GetObject("GC://" & strObjectDN)
		If Err Then
			Msg("Got err accessing GC://" & strObjectDN)
			Msg("Switchin to LDAP://" & strObjectDN)
			Set getADObject = GetObject("LDAP://" & strObjectDN)
		End If
	On Error Goto 0 
	'Set getADObject = objResult
End Function

' Function to test for group membership.
' objADObject is a user or computer object.
' strGroup - группа, принадлежность к которой проверяем.
' Returns True if the user or computer is a member of the group.
' Subroutine LoadGroups is called once for each different objADObject.
Function IsMember(ByVal objADObject, ByVal strGroup)
    
	DebugMsg "Checking membership of " & objADObject.sAMAccountName & " in " & strGroup
   
	'Если в словаре нет записи USERNAME\ - значит для него еще не загружали группы
	If (objGroupList.Exists(objADObject.sAMAccountName & "\") = False) Then
		'грузим группы для этого объекта
		Call LoadGroups(objADObject, objADObject)
		'ставим флажок USERNAME\
		objGroupList.Add objADObject.sAMAccountName & "\", True
	End If
	'после загрузки словаря просто напросто проверяем, что в словаре есть искомая комбинация
	'пользователя и пароля USERNANE\groupname
	IsMember = objGroupList.Exists(objADObject.sAMAccountName & "\" & strGroup)
End Function



' Recursive subroutine to populate dictionary object with group
' memberships (objGroupList). When this subroutine is first called
' by Function IsMember, both objPriADObject and objSubADObject are the
' user or computer object. On recursive calls objPriADObject still refers
' to the user or computer object being tested, but objSubADObject will be
' a group object. The dictionary object objGroupList keeps track of group
' memberships for each user or computer separately. For each group in
' the MemberOf collection, first check to see if the group is already in
' the dictionary object. If it is not, add the group to the dictionary
' object and recursively call this subroutine again to enumerate any
' groups the group might be a member of (nested groups). It is necessary
' to first check if the group is already in the dictionary object to
' prevent an infinite loop if the group nesting is "circular".
Sub LoadGroups(ByVal objPriADObject, ByVal objSubADObject)

	Dim colstrGroups, objGroup, j

        'группы первого уровня, в которые входит объект (без вложенности)
	colstrGroups = objSubADObject.memberOf
	' если объект никуда не входит то и проверять нечего
	If (IsEmpty(colstrGroups) = True) Then
		Exit Sub
	End If
	
	'если нашлась одна (строка)
	If (TypeName(colstrGroups) = "String") Then
		DebugMsg "Loadgroups:" & colStrGroups
		Set objGroup = getADObject (colstrGroups)
		' если мы эту группу еще не загружали в словарь - грузим
		If (objGroupList.Exists(objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
			objGroupList.Add objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
			Call LoadGroups(objPriADObject, objGroup)
        	End If
		Exit Sub
	End If

	' видимо в любом противном случае получается массив.
	For j = 0 To UBound(colstrGroups)
		DebugMsg "Loadgroups:" & colStrGroups(j)
		Set objGroup = getADObject (colstrGroups(j))
		' если мы эту группу еще не загружали в словарь - грузим
		If (objGroupList.Exists(objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
			objGroupList.Add objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
			Call LoadGroups(objPriADObject, objGroup)
		End If
	Next
End Sub



'' SIG '' Begin signature block
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' NRLCxSieCoLwZTFZKVoMogRdMBEp41ifrSi4b8uHiEqg
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
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIFD3XPzUM4lN
'' SIG '' Ga3hYzbfipYPH0gQPMb/YQ49xSttFSlcMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBADvJg66sqbekwqieG31ICFpBnijAwRWd
'' SIG '' Gu3LaHvX8K8ArUZpRe5vlyy4UMIoBi3LPtU24EqvM91G
'' SIG '' dmwYYi341eDY5Sudcua+2wvdZYjAh/0jgLIaT77aOEJH
'' SIG '' yuMww+nM9W5zyxwOvYX0A/bcazXYXbgo9uIxsg4WPRb4
'' SIG '' JrfRaOK1YLyOUhDOM4AhoE08t2qa3XMT0ZRB3L+NPnmG
'' SIG '' Isg+w405wMjFjMjT1wsg8ZSApTGLjIHrgSttCtd+iX1C
'' SIG '' JRMXNP+cVO4q45D4pG+pf4IdXSW/Z5tWY1JZx85hauhb
'' SIG '' uePvK//53i1UV+nuMVKFLYorV8EX/q5Ga7f1LC38cZnY26Q=
'' SIG '' End signature block
