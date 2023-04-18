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


