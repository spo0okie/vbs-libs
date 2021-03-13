'���������� � ��������� ����������� �������������� � ������

'���� �������� ��� ���������� ��� ����������, �.�. � ������� ����� ��� ��� �������, � ����� ��������� ��� �� ��� ��� ��������

'��� ���������� ���������� ������ ����� �������������/����������� � �������
'COMP1\group2
'COMP1\group2
'COMP1\		- ������� ����� ������ ��������, ��� ��� ����� ������� ��� ������ ��� ��������� � �������
'USER1\group1
'USER1\group2
'USER1\		- ������� ����� ������ ��������, ��� ��� ����� ������� ��� ������ ��� ��������� � �������
'�� ����� ��� ���� ����� ��������� ���������� �������� � ����, �.�. ��� ������� ������ ����� ���������� �����
'�������� ��� ������ ��������� �����. � ���� �������� ��������� �������� �������� ��� �������� ������ ������
'� ������������� - �� ����� ��� ��������
Dim objGroupList: Set objGroupList = CreateObject("Scripting.Dictionary")
objGroupList.CompareMode = vbTextCompare '��� �������� ��� ��� ������ � ������� ������� �� ����� ����� ��������

Dim objSysInfo : Set objSysInfo = CreateObject("ADSystemInfo")

'����� ����������������� ��� ����������, ��� ���� ���������� - ������ ����������
' Escape any forward slash characters, "/", with the backslash
' escape character. All other characters that should be escaped are.
Dim strUserDN : strUserDN = Replace(objSysInfo.userName, "/", "\/")
Dim strComputerDN : strComputerDN = Replace(objSysInfo.computerName, "/", "\/")

' Bind to the user and computer objects with the LDAP provider.
' ���������� ���� ����� GC:// ��� �������� ���������� ������� 
' (���������� ���� LDAP:// � ����� ��������� �� RODC)
Dim objUser : Set objUser = GetObject("GC://" & strUserDN)
Dim objComputer : Set objComputer = GetObject("GC://" & strComputerDN)




' Function to test for group membership.
' objADObject is a user or computer object.
' strGroup - ������, �������������� � ������� ���������.
' Returns True if the user or computer is a member of the group.
' Subroutine LoadGroups is called once for each different objADObject.
Function IsMember(ByVal objADObject, ByVal strGroup)
    
	DebugMsg "Checking membership of " & objADObject.sAMAccountName & " in " & strGroup
   
	'���� � ������� ��� ������ USERNAME\ - ������ ��� ���� ��� �� ��������� ������
	If (objGroupList.Exists(objADObject.sAMAccountName & "\") = False) Then
		'������ ������ ��� ����� �������
		Call LoadGroups(objADObject, objADObject)
		'������ ������ USERNAME\
		objGroupList.Add objADObject.sAMAccountName & "\", True
	End If
	'����� �������� ������� ������ �������� ���������, ��� � ������� ���� ������� ����������
	'������������ � ������ USERNANE\groupname
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

        '������ ������� ������, � ������� ������ ������ (��� �����������)
	colstrGroups = objSubADObject.memberOf
	' ���� ������ ������ �� ������ �� � ��������� ������
	If (IsEmpty(colstrGroups) = True) Then
		Exit Sub
	End If
	
	'���� ������� ���� (������)
	If (TypeName(colstrGroups) = "String") Then
		DebugMsg "Loadgroups:" & colStrGroups
	        ' Escape any forward slash characters, "/", with the backslash
		' escape character. All other characters that should be escaped are.
		' ���������� �����
		colstrGroups = Replace(colstrGroups, "/", "\/")
		' ������ ������ ������ �� LDAP (GC)
		Set objGroup = GetObject("GC://" & colstrGroups)
		' ���� �� ��� ������ ��� �� ��������� � ������� - ������
		If (objGroupList.Exists(objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
			objGroupList.Add objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
			Call LoadGroups(objPriADObject, objGroup)
        	End If
		Exit Sub
	End If

	' ������ � ����� ��������� ������ ���������� ������.
	For j = 0 To UBound(colstrGroups)
		DebugMsg "Loadgroups:" & colStrGroups(j)
		' Escape any forward slash characters, "/", with the backslash
		' escape character. All other characters that should be escaped are.
		' ���������� �����
		colstrGroups(j) = Replace(colstrGroups(j), "/", "\/")
		' ������ ������ ������ �� LDAP (GC)
		Set objGroup = GetObject("GC://" & colstrGroups(j))
		' ���� �� ��� ������ ��� �� ��������� � ������� - ������
		If (objGroupList.Exists(objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
			objGroupList.Add objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
			Call LoadGroups(objPriADObject, objGroup)
		End If
	Next
End Sub


