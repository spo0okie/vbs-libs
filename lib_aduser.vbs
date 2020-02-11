'библиотека с функциями определения принадлежности к группе

Dim objGroupList
Dim objSysInfo : Set objSysInfo = CreateObject("ADSystemInfo")

'Сразу проинициализируем все переменные, раз либу подключили - значит пригодится
' Escape any forward slash characters, "/", with the backslash
' escape character. All other characters that should be escaped are.
Dim strUserDN : strUserDN = Replace(objSysInfo.userName, "/", "\/")
Dim strComputerDN : strComputerDN = Replace(objSysInfo.computerName, "/", "\/")

' Bind to the user and computer objects with the LDAP provider.
Dim objUser : Set objUser = GetObject("LDAP://" & strUserDN)
Dim objComputer : Set objComputer = GetObject("LDAP://" & strComputerDN)

Function IsMember(ByVal objADObject, ByVal strGroup)
    ' Function to test for group membership.
    ' objADObject is a user or computer object.
    ' strGroup is the NT Name of the group to test.
    ' strGroup - группа, принадлежность к которой проверяем.
    ' objGroupList is a dictionary object with global scope.
    ' Returns True if the user or computer is a member of the group.
    ' Subroutine LoadGroups is called once for each different objADObject.

    If (IsEmpty(objGroupList) = True) Then
        Set objGroupList = CreateObject("Scripting.Dictionary")
    End If
    If (objGroupList.Exists(objADObject.sAMAccountName & "\") = False) Then
        Call LoadGroups(objADObject, objADObject)
        objGroupList.Add objADObject.sAMAccountName & "\", True
    End If
    IsMember = objGroupList.Exists(objADObject.sAMAccountName & "\" & strGroup)
End Function

Sub LoadGroups(ByVal objPriADObject, ByVal objSubADObject)
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

    Dim colstrGroups, objGroup, j

    objGroupList.CompareMode = vbTextCompare
    colstrGroups = objSubADObject.memberOf
    If (IsEmpty(colstrGroups) = True) Then
        Exit Sub
    End If
    If (TypeName(colstrGroups) = "String") Then
        ' Escape any forward slash characters, "/", with the backslash
        ' escape character. All other characters that should be escaped are.
        colstrGroups = Replace(colstrGroups, "/", "\/")
        Set objGroup = GetObject("LDAP://" & colstrGroups)
        If (objGroupList.Exists(objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
            objGroupList.Add objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
            Call LoadGroups(objPriADObject, objGroup)
        End If
        Exit Sub
    End If
    For j = 0 To UBound(colstrGroups)
        ' Escape any forward slash characters, "/", with the backslash
        ' escape character. All other characters that should be escaped are.
        colstrGroups(j) = Replace(colstrGroups(j), "/", "\/")
        Set objGroup = GetObject("LDAP://" & colstrGroups(j))
        If (objGroupList.Exists(objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
            objGroupList.Add objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
            Call LoadGroups(objPriADObject, objGroup)
        End If
    Next
End Sub


