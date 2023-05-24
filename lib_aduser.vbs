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
'' SIG '' MIIH0QYJKoZIhvcNAQcCoIIHwjCCB74CAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' NRLCxSieCoLwZTFZKVoMogRdMBEp41ifrSi4b8uHiEqg
'' SIG '' ggUsMIIFKDCCAxCgAwIBAgIBADANBgkqhkiG9w0BAQsF
'' SIG '' ADBtMQswCQYDVQQGEwJSVTENMAsGA1UECAwEVXJhbDEU
'' SIG '' MBIGA1UEBwwLQ2hlbHlhYmluc2sxETAPBgNVBAoMCFJl
'' SIG '' dmlha2luMQswCQYDVQQLDAJJVDEZMBcGA1UEAwwQcmV2
'' SIG '' aWFraW4tcm9vdC1DQTAeFw0yMzA1MjQwNDU2NTdaFw0y
'' SIG '' NDA2MDIwNDU2NTdaMGAxCzAJBgNVBAYTAlJVMQ0wCwYD
'' SIG '' VQQIDARVcmFsMQ0wCwYDVQQHDARDaGVsMREwDwYDVQQK
'' SIG '' DAhSZXZpYWtpbjELMAkGA1UECwwCSVQxEzARBgNVBAMM
'' SIG '' CnNjcmlwdHNpZ24wggEiMA0GCSqGSIb3DQEBAQUAA4IB
'' SIG '' DwAwggEKAoIBAQDBTtnKwGde6qQttu1TOo/JIGTZ2hoa
'' SIG '' HIGDBFKgexeDT8choad2DXRQzxGyu2y9w7djwuthEODY
'' SIG '' KLVf12PcofOKnowAoSIqQ7VW77I8I4VLI7hi0VDGZ9V9
'' SIG '' W4pC/mcJjkaEMSAFj6/CST5tpeczI2KxYpM1f+mEWGiu
'' SIG '' TkB3K3jVhsaDCuWZYZoszAJkUgp3SevPyqA6JuqzwpHD
'' SIG '' aDbNG5ohd1MwcwvRKab6HNwkEprYyTiX6uWZ8rBGyIGE
'' SIG '' 4ZtshlAt6yyn6U/tYREG9+pA9CzoPHfB3gh6taeR0/25
'' SIG '' oeZ5WYHuGMNeHaHYeeIXKS9gfPh3ND/fJGQaTljVSGX5
'' SIG '' e3StAgMBAAGjgd8wgdwwHQYDVR0OBBYEFDc+8unMGviq
'' SIG '' cvfVA1vXi3LqheoJMB8GA1UdIwQYMBaAFKJJoRQ/bOk/
'' SIG '' S1B2wDmCrQ0ZzJbKMA8GA1UdEwEB/wQFMAMBAf8wDgYD
'' SIG '' VR0PAQH/BAQDAgGGMDgGA1UdHwQxMC8wLaAroCmGJ2h0
'' SIG '' dHA6Ly9wa2kucmV2aWFraW4ubmV0L2NybC9yb290LWNh
'' SIG '' LmNybDA/BggrBgEFBQcBAQQzMDEwLwYIKwYBBQUHMAKG
'' SIG '' I2h0dHA6Ly9wa2kucmV2aWFraW4vbmV0L3Jvb3QtY2Eu
'' SIG '' Y3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCyB0c0PKF0ffSX
'' SIG '' RmTBaqNWVOEpokgkdJbUNhVhKL4d7MR2wF1GX6rTeGTD
'' SIG '' hF4p1R3N6wRR0AAFVfp63st3w51XoQbJmGInJ7IFgrB2
'' SIG '' 7G6XzFVkp0llNu/1ygiqHm8v7JZEhdiqCun+JDd0ata4
'' SIG '' HKz2lca85tg2wnDfm0n3N7cdI56UkB+dKAzMLINVNT9X
'' SIG '' GSF70kXtCSPeLPDorVge0oWLxDvUiYAzlLvXk2+MTlrJ
'' SIG '' ka3R/s84X5W6CP9JJptIuzVuSd5ETB+tz/6xid2ELhNK
'' SIG '' ihkETnTViqdKp0CFGS/tRSDnfQ7Kp+Udr/SL7V/cg6Kh
'' SIG '' y8tXMCW+EJQBhrAGhudOvnIcFtTrUmhjupqMUuaLqDVY
'' SIG '' ACSwtmuihx7RAKREee0d8DJ99P3unNqfThtTPfHCzgeU
'' SIG '' Yk+i505Y8Op7G286bAwMv+m6SvnOT8vexSzJ3c77Vuyv
'' SIG '' HEU49MkgZAhpajQjTeOq2Kj3o1m+jxQ3OkWgMD6EMoJ8
'' SIG '' PIQS1XPhXcZ1N81uheeUf9EX13m32CulsDHmOnhQcT56
'' SIG '' jKt/9dn4jqHodqEodaz2Jb/tu7u6uIHmuaB2g6DTRxAO
'' SIG '' v33V/0yI40FG0SPAoNsWNVFySO5UwnewXA6H1hWEFezZ
'' SIG '' UPWWnqWb+F2uNUC8gl7Uguc2q3pJ5RhoJX+TxgBIt3oW
'' SIG '' SrZ8foMC3jGCAf0wggH5AgEBMHIwbTELMAkGA1UEBhMC
'' SIG '' UlUxDTALBgNVBAgMBFVyYWwxFDASBgNVBAcMC0NoZWx5
'' SIG '' YWJpbnNrMREwDwYDVQQKDAhSZXZpYWtpbjELMAkGA1UE
'' SIG '' CwwCSVQxGTAXBgNVBAMMEHJldmlha2luLXJvb3QtQ0EC
'' SIG '' AQAwDQYJYIZIAWUDBAIBBQCgXjAQBgorBgEEAYI3AgEM
'' SIG '' MQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAv
'' SIG '' BgkqhkiG9w0BCQQxIgQgUPdc/NQziU0ZreFjNt+Klg8f
'' SIG '' SBA8xv9hDj3FK20VKVwwDQYJKoZIhvcNAQEBBQAEggEA
'' SIG '' dy98klyRNkxZZrd2HVrtEOGIoPNAHsCaSOgLpSao8rx2
'' SIG '' 7ge+IDs4L0tw3lPVvKVss2OkqtFlglexeF412UjVJEfs
'' SIG '' geyMtNbAZ5W88sfgxfGjXi9nlYFz/fElVEMd9lPY5zc5
'' SIG '' k6JShmKwvGuTv8knpjPLswPUHmpL/bwhGn6wgUdAw/gs
'' SIG '' 39c+jD/LTjUcM3LR9NlWQQBKH1/8ArbEWn/y5UwA2iMc
'' SIG '' hByvAN+xcb9xszsSHcq3ESn/DFJMcbIKPzFY+gm66srO
'' SIG '' 6M83yKCu7v1L///PJoxicasz8RHJ/Llpz44oavq2Gr3q
'' SIG '' QF6eUlzLlUxGKvNUeIzmcHfqiZz0plTqXg==
'' SIG '' End signature block
