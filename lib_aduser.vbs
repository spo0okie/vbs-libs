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
Option explicit
Dim objGroupList: Set objGroupList = CreateObject("Scripting.Dictionary")
objGroupList.CompareMode = vbTextCompare '��� �������� ��� ��� ������ � ������� ������� �� ����� ����� ��������

Dim objSysInfo : Set objSysInfo = CreateObject("ADSystemInfo")
'����� ����������������� ��� ����������, ��� ���� ���������� - ������ ����������
' Escape any forward slash characters, "/", with the backslash
' escape character. All other characters that should be escaped are.
on error resume next
Dim strUserDN : strUserDN = Replace(objSysInfo.userName, "/", "\/")
Dim strComputerDN : strComputerDN = Replace(objSysInfo.computerName, "/", "\/")
on error goto 0

' Bind to the user and computer objects with the LDAP provider.
Dim objUser : if (Len(strUserDN)) then Set objUser = getADObject (strUserDN)
Dim objComputer : if (Len(strComputerDN)) then Set objComputer = getADObject (strComputerDN)

'�������� ������ �� �� ������� ����� ��, ���� �� ����� �� ����� ����
Function getADObject(ByVal strObjectDN)
	'MsgBox "GetAdObject: " & strObjectDN
	'��������� �������� ����� ��� �������������
	strObjectDN = Replace(strObjectDN, "/", "\/")
	' ���������� ���� ����� GC:// ��� �������� ���������� ������� 
	' (���������� ���� LDAP:// � ����� ��������� �� RODC)
	' ��� ��� ����������� ������ "� ���������� ������ ���������� ������������"
	' ��� �� ���� ������, �� �����������, ��� �������� � GC
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
		Set objGroup = getADObject (colstrGroups)
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
		Set objGroup = getADObject (colstrGroups(j))
		' ���� �� ��� ������ ��� �� ��������� � ������� - ������
		If (objGroupList.Exists(objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
			objGroupList.Add objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
			Call LoadGroups(objPriADObject, objGroup)
		End If
	Next
End Sub


