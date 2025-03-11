Option Explicit
'v.1	�������� ��������� ���������� ��� ������ � WMI

Dim WmiErr

'�������������� WMI ������
on error resume next
	Err.Clear
	Dim objWmiService: Set objWmiService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	HaltTextIfError "GetObject(""winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"")"
on error goto 0

'������ ������ �� WMI � ����������� ������ � ��� ���������
function getWmiQuery(ByVal Query)
	Err.Clear
	on error resume next
	set getWmiQuery=objWMIService.ExecQuery (Query)
	on error goto 0
end function

'������ ������� ������ �� WMI � ����������� ������ � ��� ���������
function getWmiQueryArray(ByVal Query)
	Dim colItems, objItem
	Dim arrResults()
	Dim i, strValue
	set colItems= getWmiQuery(Query)
	on error resume next

	' �������� ���������� ����������� ��� ����������� ������� �������
	i = 0
	For Each objItem In colItems
		i = i + 1
	Next

	' �������������� ������ ����������� �������
	ReDim arrResults(i-1)
    
	' ��������� ������ �������� �� �������
	i = 0
	For Each objItem In colItems
		Set arrResults(i) = objItem
		i = i + 1
	Next
    
	' ���������� ���������
	getWmiQueryArray = arrResults
	on error goto 0
end function


'������ ������ �� WMI � ��������� ������� � ������ ������
function getWmiQueryCrit(ByVal Query)
	set getWmiQueryCrit = getWmiQuery(Query)
	HaltTextIfError "getWmiQueryCrit: objWMIService.ExecQuery(" & Query & ")"
end function

'������ ������ �� WMI � ��������� ������� � ������ ������
function getWmiQueryArrayCrit(ByVal Query)
	getWmiQueryArrayCrit = getWmiQueryArray(Query)
	HaltTextIfError "getWmiQueryArrayCrit: objWMIService.ExecQuery(" & Query & ")"
end function



function getOsCaption()
	dim oss : Set oss = getWmiQueryCrit ("Select * from Win32_OperatingSystem")
	dim os
	For Each os in oss
	    getOsCaption=os.Caption
	next
	if (Platform="AMD64") then
		getOsCaption=getOsCaption & " (x64)"
	end if
end function


'�������� �������� ��
Function GetOS    
    GetOS="UNKNOWN"
    dim colOS:  Set colOS = getWmiQueryCrit("Select * from Win32_OperatingSystem")
    dim objOS
    For Each objOS in colOS
	'wscript.echo objOS.Caption
        If instr(objOS.Caption, "Windows 10") Then
        	GetOS = "Windows 10"
        elseIf instr(objOS.Caption, "Windows 8") Then
        	GetOS = "Windows 8"    
        elseIf instr(objOS.Caption, "Windows 7") Then
        	GetOS = "Windows 7"    
        elseIf instr(objOS.Caption, "Vista") Then
        	GetOS = "Windows Vista"
        elseIf instr(objOS.Caption, "Windows XP") Then
      		GetOS = "Windows XP"
        elseIf instr(objOS.Caption, "Windows Server 2012 R2") Then
      		GetOS = "Windows Server 2012 R2"
        elseIf instr(objOS.Caption, "Windows Server 2012") Then
      		GetOS = "Windows Server 2012"
        elseIf instr(objOS.Caption, "Windows Server 2008 R2") Then
      		GetOS = "Windows Server 2008 R2"
        elseIf instr(objOS.Caption, "Windows Server 2008") Then
      		GetOS = "Windows Server 2008"
        End If
	Next
End Function
