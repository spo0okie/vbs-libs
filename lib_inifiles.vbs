option Explicit
'-----------------------------------------------------------------------
'INI FILES LIBRARY           -------------------------------------------
'-----------------------------------------------------------------------
'������ ���������� �� ������ ������ ��������� ������ (INI��������, �� �� ���)
'���������� �������� ��� ����������� ������� �����
'eq		: ����� ���������  					= : => := ...
'secL, secR	: ������������ ���������� ������ 	[, ]
'ComDelims	: � ���� ���������� ������� 		# rem :: ' // ...
'Cst, Cend	: ������ ������ � ����� ����� �������� /* */ <!-- --!> �� �������������� (��������� �� �������)
'CaSense	: ���������������� ���������� � ������ � �������� true\false
'defCrLf	: ������� ������ �� ��������� (���������� ���, ���� �� ������ ����� � ����� ��� ������������)
'Section	: C ����� ������� ����� ��������
'KeyName	: � ����� ����������
'Value		: �������� (�������� �� ��������� ��� ������, ���������� ��� ������)
'req		: read\write

'�������� ����������� ����� ������ ��� ���������� ��������� � ����� �����


'���������� ������ ������� � ������ (ComDelims - ������ �������� ������ �������)
'��� ����� ������+1 ���� ��� �������
function CommentAt(byVal INIString, byVal ComDelims, byVal Cst, byVal Cend, byVal CaSense)
	Dim startPos, testPos, Delim
	if not CaSense then
		INIString=UCase(INIString)
	end if
	startPos=Len(INIString)+1	'�� ��������� ������� ��� � �� �� ������ ������
	for each Delim in ComDelims
		if not CaSense then
			Delim=UCase(Delim)
		end if
		testPos = InStr(1, INIString, Delim, vbTextCompare)
		if testPos>0 then
			startPos=min(startPos,testPos)
		end if
	next
	CommentAt = startPos
end function

'���������� �������� ����� ������, ��� �������� � ������ �� �����
function GetUsefulPart(INIString, ComDelims, Cst, Cend, CaSense)
	GetUsefulPart=TrimWithTabs(Left(INIString,CommentAt(INIString, ComDelims, Cst, Cend, CaSense)-1))
end function

'���� ������ �������� ����������� ������, �� ���������� �� ���, ����� ""
function GetSectionName(INIString,ldelim,rdelim)
	If len(ldelim)>0 and len(rdelim)>0 and Left(INIString, len(ldelim)) = ldelim and Right(INIString, len(rdelim)) = rdelim then
		GetSectionName=TrimWithTabs(Mid(INIString,len(ldelim)+1,Len(INIString)-len(ldelim)-len(rdelim)))
	else
		GetSectionName=""
	end if
end function

'���������� ������-��������� ������, ���� ��� ���������� �� �������
function declareSectionName(CurSection,Section,ldelim,rdelim)
	if Len(Section)>0 and Section<>CurSection and Len(ldelim)>0 and Len(rdelim)>0 then
		declareSectionName=ldelim&Section&rdelim	'���������� ������
	else
		declareSectionName=""	'������ �� ���� ���������
	end if
end function

'���������� ������-��������� ����������
function declareVariable(varName,eq,varValue)
	declareVariable=varName&eq&varValue
end function

'���� � ������ ��� ����������
function GetVariableName(INIString,eq)
	dim eqPos
	eqPos = InStr(1, INIString, eq, vbTextCompare)
	If eqPos>0 Then
		GetVariableName=TrimWithTabs(Left(INIString,eqPos-len(eq)))
	else
		'wscript.echo eq & " not found in " & INIString
		GetVariableName=""
	end if
end function

'���������� ����� ������, ���� ������ �� ������
function optionalCrLf(somedata,CrLf)
	if Len(somedata)>0 then
		optionalCrLf=CrLf
	else
		optionalCrLf=""
	end if
end function

'���������� ������-�����������
function commentLine(Comment,ComDelims,CrLf)
	commentLine=""
	'wscript.echo "comline: " & UBound(ComDelims) & " " & Len(Comment)
	if IsArrayDimmed(ComDelims) and Len(Comment)>0 then
		commentLine = ComDelims(0) & Comment & CrLf
	end if
end function

'���� � ������ �������� ����������
function GetVariableVal(INIString,eq)
	dim eqPos
	eqPos = InStr(1, INIString, eq, vbTextCompare)
	If eqPos>0 Then
		GetVariableVal=TrimWithTabs(Right(INIString,Len(INIString)-eqPos-len(eq)+1))
	else
		GetVariableVal=""
	end if
end function

'���� � ����� ����� �������� ������������ ������� ������
function GetFileCrLf(INIContents,default)
	'���� �� �������� ����� ������� �� ���������, �� ��� ���� ����� �����
	if Len(default)=0 then
		default=vbCrLf
	end if

	'������� ����� ������� � ��� ��� ����������� � �����
	dim crPos,lfPos
	crPos = InStr(1, INIContents, vbCr, vbTextCompare)
	lfPos = InStr(1, INIContents, vbLf, vbTextCompare)
	If crPos>0 or lfPos>0 Then
		if (crPos>0) and ((crPos<lfPos) or (lfPos=0)) then '���������� �� CrLf
			if lfPos=crPos+1 then
				'wscript.echo "GetFileCrLf: dected Windows CRLF at " & crPos
				GetFileCrLf=vbCr&vbLf 	'Windows
			else
				'wscript.echo "GetFileCrLf: dected UNIX CR at " & crPos
				GetFileCrLf=vbCr		'UNIX
			end if
		elseif (lfPos>0) and ((lfPos<crPos) or (crPos=0)) then '���������� �� LfCr
			if lfPos=crPos-1 then
				'wscript.echo "GetFileCrLf: dected LFCR WTF!!!??? " & lfPos
				GetFileCrLf=vbLf&vbCr'���� ����� � �� ��������� �����???
			else
				'wscript.echo "GetFileCrLf: dected Mac LF at " & lfPos
				GetFileCrLf=vbLf		'Mac
			end if
		else '???WTF???
			GetFileCrLf=default '�� ���������
		end if
	else ' �� ����� �� cr �� lf
		wscript.echo "GetFileCrLf: Newline symbol not detected, using default: "&default
		GetFileCrLf=default '�� ���������
	end if
end function


'����� ���� ���� ����������, �� ����� �������, ������� ������� ��������� ��� ��� �������� ���� �������� ��������������� � �������


'������������� ������ ������ ������ (INI��������, �� �� ���) - ������ ��� ����� ������->����->��������
'������ ���� �� ��������� ���������������:
'eq			: ����� ���������  					= : => := ...
'secL, secR	: ������������ ���������� ������ 	[, ]
'ComDelims	: � ���� ���������� ������� 		# rem :: ' // ...
'Cst, Cend	: ������ ������ � ����� ����� �������� /* */ <!-- --!> �� �������������� (��������� �� �������)
'CaSense	: ���������������� ����������� ������ � �������� true\false
'defCrLf	: ������� ������ �� ��������� (���������� �����, ���� �� ������ ����� � ����� ��� ������������)
'Section	: C ����� ������� ����� ��������
'KeyName	: � ����� ����������
'Value		: �������� (�������� �� ��������� ��� ������, ���������� ��� ������)
'req		: read\write
Function parseINIString(FileName, eq, secL, secR, ComDelims, Cst, Cend, CaSense, defCrLf, Section, KeyName, Value, req, Comment)
	Dim i, INIContents, INIStrings, INIString, UsefulPart, CurSection, CrLf
	Dim testSec, testVar, testVal, jobSec, jobVar, jobDone

	'Get contents of the INI file As a string
	DebugMsg "reading file : "&FileName
	INIContents = GetFile(FileName)
	CrLf=GetFileCrLf(INIContents,defCrLf)
	INIStrings = Split (INIContents, CrLf)
	DebugMsg "found lines : "&ubound(INIStrings)&CrLf

	CurSection=""		'�� ������ ����� ������ ��� (����� � �� ����� �� ����� �����)
	jobDone=false		'������ �� �������
	jobSec = Section	'���� ����� ������
	jobVar = KeyName	'� ����� ����������
	if not CaSense then	'���� ��������� �� ������������ � ��������, �� �������� ��� ������� � ������ ��������
		jobSec=UCase(jobSec)
		jobVar=UCase(jobVar)
	end if

	for i = 0 to UBound(INIStrings)
		INIString=INIStrings(i)
		UsefulPart=GetUsefulPart(INIString, ComDelims, Cst, Cend, CaSense) '�������� �������� ����� ������
		testSec=GetSectionName(UsefulPart,secL,secR)	'������� ���� �� � ��� ����� ������
		testVar=GetVariableName(UsefulPart,eq)	'��� ����������
		if not CaSense then	'���� ��������� �� ������������ � ��������, �� �������� ��� ������� � ������ ��������
			testSec=UCase(testSec)
			testVar=UCase(testVar)
		end if
		if Len(testSec)>0 then
			DebugMsg "Section detected: " & testSec & " vs " & jobSec
			if CurSection=jobSec then	'��������� ������ � ������� �� ������ ����������
				if req="write" and not jobDone then '� ��� ������ ���� ���� �������� ����������, � �� �� ����
					'������� � ����� ���������� ������
					DebugMsg "INSERTING BEFORE SECTION :" &i
					INIStrings(i)=	commentLine(Comment,ComDelims,CrLf)& _
									declareVariable(KeyName,eq,Value)&_
									CrLf & INIStrings(i)
					jobDone=true
				end if
			end if
			CurSection=testSec
		elseif len(testVar)>0 then
			testVal=GetVariableVal(UsefulPart,eq)
			DebugMsg "Variable detected: " & testVar & " => " & testVal
			if CurSection=jobSec and testVar=jobVar then
			'�� ����� ���� ����������
				if req="write" then '������
					'������ ������ �� ��������� ����� '������������ ������� ����� '����� ����
					if testVal<>Value then
						DebugMsg "CHANGING CURRENT :" &i
						INIStrings(i)=	commentLine(INIStrings(i),ComDelims,CrLf)& _
										commentLine(Comment,ComDelims,CrLf)& _
										declareVariable(KeyName,eq,Value)
					end if
					jobDone=true
				elseif req="delete" then '�������� �������� '������������ ������� ����� 
						DebugMsg "CHANGING CURRENT :" &i
						INIStrings(i)=	commentLine(Comment,ComDelims,CrLf)& _
										commentLine(INIStrings(i),ComDelims,"")
				
					jobDone=true
				else '������
					parseINIString=testVal
					jobDone=true
				end if
			end if
		end if
	next
	if not jobDone then	'���������� �� �����
		if req="write" then '������
			'��������� ��� ���� �����
			redim preserve INIStrings(UBound(INIStrings)+1)
			if (UBound(INIStrings) > i+1) then
				dim j:	for j=UBound(INIStrings) to i+1
					DebugMsg(j)
					INIStrings(j)=INIStrings(j-1)
				next
			end if
			'i=i-1
			DebugMsg "Inserting in string " &i
			INIStrings(i)=	declareSectionName(CurSection, Section, secL, secR)& _
							optionalCrLf(declareSectionName(CurSection, Section, secL, secR),CrLf)&_
							commentLine(Comment,ComDelims,CrLf)& _
							declareVariable(KeyName,eq,Value) '&CrLf
			DebugMsg INIStrings(i)
			jobDone=true
		else '������
			parseINIString=Value
		end if
	end if
	if req="write" or req="delete" then '������
		if jobDone then WriteFile FileName, Join(INIStrings,CrLf)
		parseINIString=jobDone
	end if

End Function


'������������� ������ ��������� ������, ���� ������, ��������� ������ �� ��� ����� ��������
'������ ���� �� ��������� ���������������:
'eq			: ����� ���������  					= : => := ...
'secL, secR	: ������������ ���������� ������ 	[, ]
'ComDelims	: � ���� ���������� ������� 		# rem :: ' // ...
'Cst, Cend	: ������ ������ � ����� ����� �������� /* */ <!-- --!> �� �������������� (��������� �� �������)
'CaSense	: ���������������� ����������� ������ � �������� true\false
'defCrLf	: ������� ������ �� ��������� (���������� �����, ���� �� ������ ����� � ����� ��� ������������)
'Section	: C ����� ������� ����� ��������
'Keystring	: ����� ������ �������
'Addition	: ��� ��������
'Position	: before\after
'Many		: �������� ��� ��� ���������
Function parseTXTString(FileName, eq, secL, secR, ComDelims, Cst, Cend, CaSense, defCrLf, Section, Keystring, Addition, Position, Many)
	Dim i, INIContents, INIStrings, INIString, UsefulPart, CurSection, CrLf
	Dim testSec, jobSec, jobDone

	'Get contents of the INI file As a string
	'wscript.echo "reading file : "&FileName
	INIContents = GetFile(FileName)
	CrLf=GetFileCrLf(INIContents,defCrLf)
	INIStrings = Split (INIContents, CrLf)
	'wscript.echo "parseTXTString: found " & ubound(INIStrings) & " lines with delim of " & Len(CrLf) &" bytes"

	CurSection=""		'�� ������ ����� ������ ��� (����� � �� ����� �� ����� �����)
	jobDone=false		'������ �� �������
	jobSec = Section	'���� ����� ������
	if not CaSense then	'���� ��������� �� ������������ � ��������, �� �������� ��� ������� � ������ ��������
		jobSec=UCase(jobSec)
		'msg ("parseTXTString: case insensetive mode")
	end if
	for i = 0 to UBound(INIStrings)
		INIString=INIStrings(i)
		'msg("at first it is "&INIString	)
		testSec=GetSectionName(GetUsefulPart(INIString, ComDelims, Cst, Cend, CaSense),secL,secR)	'������� ���� �� ����� ������
		if not CaSense then	'���� ��������� �� ������������ � ��������, �� �������� ��� ������� � ������ ��������
			testSec=UCase(testSec)
		end if
		if Len(testSec)>0 then
			CurSection=testSec
		elseif INIString=Keystring then
			'msg ("parseTXTString: found position for insert!")
			if Ucase(Position)="BEFORE" then
				INIStrings(i) =	Addition & CrLf & INIStrings(i)
				msg ("parseTXTString: found position for insert!")
			else
				INIStrings(i) =	INIStrings(i) & CrLf & Addition
			end if
			jobDone=true
			if not Many then
				exit for
			end if
		else
			'msg (INIString & " != " & Keystring)
		end if
	next
	if jobDone then WriteFile FileName, Join(INIStrings,CrLf)
	parseTXTString=jobDone
End Function


'����� ���� ��� �������, ������� ��������� ������ � ����� 2�� ��� ����


'��������� ��� ���������� ���������� - ������������� ��������� ����������� ��� �����
Function CheckFileTypeDescr(ByVal FType)
	'eq			: ����� ���������  					= : => := ...
	'secL, secR	: ������������ ���������� ������ 	[, ]
	'ComDelims	: � ���� ���������� ������� 		# rem :: ' // ...
	'Cst, Cend	: ������ ������ � ����� ����� �������� /* */ <!-- --!> �� �������������� (��������� �� �������)
	'CaSense	: ���������������� ����������� ������ � �������� true\false
	'defCrLf	: ������� ������ �� ��������� (���������� �����, ���� �� ������ ����� � ����� ��� ������������)
	'Msg typeName (FType)
	CheckFileTypeDescr=true
	dim flds,fld
	flds=array("eq","secL","secR","ComDelims","Cst","Cend","CaSense","defCrLf")
	for each fld in flds
		if not FType.exists(fld) then
			msg "CheckFileTypeDescr err: " & fld & " not set"
			CheckFileTypeDescr=false
		end if
	next
End Function


'��� ���������� ���� ���� ���� � �������� ��� �����������

'�������� �������� �� INI �����
Function conffile_get(ByVal FPath, ByVal FType, ByVal Section, ByVal Key, ByVal Default)
	if CheckFileTypeDescr(FType) then
		conffile_get=parseINIString(FPath, FType("eq"), FType("secL"), FType("secR"), FType("ComDelims"), FType("Cst"), FType("Cend"), FType("CaSense"), FType("defCrLf"), Section, Key, Default, "read", "")
	else
		conffile_get=Default
	end if
end Function


'�������� �������� � INI ����
Function conffile_set(ByVal FPath, ByVal FType, ByVal Section, ByVal Key, ByVal Value, ByVal Comment)
	dim mode
	if Value=unset_me then
		mode="delete"
	else
		mode="write"
	end if
	if CheckFileTypeDescr(FType) then
		conffile_set=parseINIString(FPath, FType("eq"), FType("secL"), FType("secR"), FType("ComDelims"), FType("Cst"), FType("Cend"), FType("CaSense"), FType("defCrLf"), Section, Key, Value, mode, Comment)
	else
		conffile_set=false
	end if
end Function


'��������� �������� � INI ����
'���������� ������� ������������� � ����� ���������
Function conffile_fix(ByVal FPath, ByVal FType, ByVal Section, ByVal Key, ByVal Value, ByVal Comment)
	if CheckFileTypeDescr(FType) then
		if (Value=conffile_get(FPath,FType,Section,Key,"value not found marker")) then
			conffile_fix=false
		else
			conffile_set FPath,FType,Section,Key,Value,Comment
			conffile_fix=true
		end if
	else
		conffile_fix=false
	end if
end Function




'��������� ���������� ����� ������ ������
Function textfile_set_after(ByVal FPath, ByVal FType, ByVal Section, ByVal Key, ByVal Value, ByVal Comment, ByVal after)
	if CheckFileTypeDescr(FType) then
		if not conffile_get(FPath, FType, Section, Key, "un1Que_deFFault") = "un1Que_deFFault" then '���� ��� ���� �����-�� �������� - ������
			textfile_set_after=parseINIString(FPath, FType("eq"), FType("secL"), FType("secR"), FType("ComDelims"), FType("Cst"), FType("Cend"), FType("CaSense"), FType("defCrLf"), Section, Key, Value, "write", Comment)
		else
			dim CrLf, Addition
			CrLf=GetFileCrLf(GetFile(FPath),FType("defCrLf"))
			Addition= commentLine(Comment,FType("ComDelims"),CrLf) & declareVariable(Key,FType("eq"),Value)
			'msg("trying to add " & Addition & " after " & after)
			textfile_set_after=parseTXTString(FPath, FType("eq"), FType("secL"), FType("secR"), FType("ComDelims"), FType("Cst"), FType("Cend"), FType("CaSense"), FType("defCrLf"), Section, after, Addition, "after", true)
		end if
	else
		conffile_set=false
	end if
end Function




'����������� ����� ������ ��� ����������
'���� .bat/.cmd
dim ftype_bat : set ftype_bat = CreateObject("Scripting.Dictionary")
	ftype_bat.add "eq",		"="
	ftype_bat.add "secL",	""
	ftype_bat.add "secR",	""
	ftype_bat.add "ComDelims",array("rem ","::")
	ftype_bat.add "Cst",	""
	ftype_bat.add "Cend",	""
	ftype_bat.add "CaSense",false
	ftype_bat.add "defCrLf",vbCrLf

'���� .properties �� ��
dim ftype_tc_conf : set ftype_tc_conf = CreateObject("Scripting.Dictionary")
	ftype_tc_conf.add "eq",		"="
	ftype_tc_conf.add "secL",	""
	ftype_tc_conf.add "secR",	""
	ftype_tc_conf.add "ComDelims",array("#")
	ftype_tc_conf.add "Cst",	""
	ftype_tc_conf.add "Cend",	""
	ftype_tc_conf.add "CaSense",true
	ftype_tc_conf.add "defCrLf",vbCrLf

'���� etc\hosts
dim ftype_etc_hosts : set ftype_etc_hosts = CreateObject("Scripting.Dictionary")
	ftype_etc_hosts.add "eq",		" "
	ftype_etc_hosts.add "secL",	""
	ftype_etc_hosts.add "secR",	""
	ftype_etc_hosts.add "ComDelims",array("#")
	ftype_etc_hosts.add "Cst",	""
	ftype_etc_hosts.add "Cend",	""
	ftype_etc_hosts.add "CaSense",false
	ftype_etc_hosts.add "defCrLf",vbCrLf

'���� .ini
dim ftype_tc_ini : set ftype_tc_ini = CreateObject("Scripting.Dictionary")
	ftype_tc_ini.add "eq",		"="
	ftype_tc_ini.add "secL",	"["
	ftype_tc_ini.add "secR",	"]"
	ftype_tc_ini.add "ComDelims",array(";")
	ftype_tc_ini.add "Cst",	""
	ftype_tc_ini.add "Cend",	""
	ftype_tc_ini.add "CaSense",true
	ftype_tc_ini.add "defCrLf",vbCrLf

'���� .ini
dim ftype_ini : set ftype_ini = CreateObject("Scripting.Dictionary")
	ftype_ini.add "eq",		"="
	ftype_ini.add "secL",	"["
	ftype_ini.add "secR",	"]"
	ftype_ini.add "ComDelims",array("#")
	ftype_ini.add "Cst",	""
	ftype_ini.add "Cend",	""
	ftype_ini.add "CaSense",true
	ftype_ini.add "defCrLf",vbCrLf

'' SIG '' Begin signature block
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' RUoK2rWWTbwKmcV+4WXW7mD7H1j6qN3tv/r3iQoAa6eg
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
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIA+j8d/jVt8y
'' SIG '' e69yUin0T4fbVYYdMi/JnI+aMq37qI8lMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAGshuRR2FyrAH6J1X5UpkFTJh4XbNGWe
'' SIG '' VkqhgxRGBdm3h+G90fVg4ux1bAVZ4A9nrVkvqXPOse41
'' SIG '' sng5Ac8RfsxxkcImfnGGsbNII8JW/9r4zfXclm5WvPEF
'' SIG '' Yy/skcAPlqeVg9b/92jJlHcAVTBiLA5oMQok5xsrVSq/
'' SIG '' M9flCL8XUwktLcfHPClrYqZ3MhwsTRJUVQ2UdFPxzrwV
'' SIG '' jIPIPlzDkkO7EvG7X0KeDvDMkF7HEXpodV0SHRVepLZw
'' SIG '' 5wb+1n7FjBTMB4wdESZVSvJMlTU4tpuhyAHS5COVt0Y6
'' SIG '' 9YVoAYagTtuVOhvdjfxSBV9vk39sSlt3OtZJuX3G7R5ZmV4=
'' SIG '' End signature block
