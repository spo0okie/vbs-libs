'����������� � ���������� ��������� ���������, ������� ����������� ����������
'��� ����� �� ����� �� ��������
Option Explicit
Const coreLibVer="2.13"
'v2.13 + ComputerDomain (�������� �� ������� ����� ��� ComputerName)
'v2.12 * regDeleteRecursive � ������ ������� ������� ��������� ������ � � ������
'v2.11 * unset_me ���������� � ����, �.�. ������ ������������ � ��� ������� � ��� INI
'v2.10   computerName ������ �������� ������ ��� ����������
'        ������ ��������� �� 15�� NETBIOS ��� ������ ����� � compName
'v2.9    launchPad ��������� �������� ��������� �������
'v2.8    isProcRunning �������� � lib_procs
'v2.7  + �������� ������ objWmi
'v2.6  + ��������� ���������� Platform, Systemdrive
'v2.5  ! ��������������� ������� launchpad, ������� ����������� ������ �� ��������������������
'        ������������. ���� ������ ������ �� Win10: ������ ��� ����������� ����� �������� %username%,
'        �.�. ���� ����������, ��� ������ ���� ������� ���, ��� ��������� ������ �� ������ ������������
'        ������ �������� ����� �� �� ������� ���� ���������� ���� � ���������� �������.
'v2.4  * ������� Msg ������ ������������ ���������� ���������� LogFile � ��� ������ � ��� ������
'        ���� �������� ������, �� ����� ����� � ����, ���� ������ - �� ��������� ����, ����� � ���������
'        ��� ������ �������� � ��� ������� ��� ������, ����� ���������� ���������� ������� ����� 
'        launchpad
'      + �������� ������ objReg ��� ������ � �������� ����� WMI (����� ����� �������� �-� wshShell)


const unset_me=		"#UNSET_me#" '��� �������� ������� � ���������� ������� ���� ������

Dim wshShell	: Set wshShell = WScript.CreateObject("WScript.Shell")
Dim objUserEnv	: Set objUserEnv = wshShell.Environment("USER")
Dim objSystemEnv: Set objSystemEnv = wshShell.Environment("SYSTEM")
Dim objProcessEnv:Set objSystemEnv = wshShell.Environment("PROCESS")
Dim objVolatileEnv:Set objSystemEnv = wshShell.Environment("VOLATILE")
Dim objShell	: Set objShell = CreateObject("Shell.Application")
Dim objFSO	: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objReg	: Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
Dim objWmi	: Set objWmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Dim WorkDir : WorkDir =	WshShell.ExpandEnvironmentStrings("%TEMP%") & "\"
Dim WindowsDir : WindowsDir = objFSO.GetSpecialFolder(0)

Dim CompName : ComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
Dim UserName : UserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
Dim UserDomain : UserDomain = wshShell.ExpandEnvironmentStrings( "%USERDOMAIN%" )
'������ ��� �������� - �� ����� ��� ����������� ������. � �������� ��� ������� ���������� 15������� NETBIOS ���
Dim Computername : Computername=wshShell.RegRead ("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\HostName")
Dim ComputerDomain : ComputerDomain = wshShell.RegRead ("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Domain")
Dim UserProfile : UserProfile = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
Dim Platform : Platform = wshShell.ExpandEnvironmentStrings( "%PROCESSOR_ARCHITECTURE%" )
Dim SystemDrive : SystemDrive = wshShell.ExpandEnvironmentStrings( "%SYSTEMDRIVE%" )
Dim SystemRoot : SystemRoot = wshShell.ExpandEnvironmentStrings( "%SYSTEMROOT%" )

Dim DEBUGMODE : DEBUGMODE = 0

'
' ������ ��������� ������:
'
Const LocalSystemSID	= "S-1-5-18"
Const LocalAdminsSID	= "S-1-5-32-544"
Const LocalUsersSID	= "S-1-5-32-545"

' Constants (taken from WinReg.h)
'
Const HKEY_CLASSES_ROOT   = &H80000000
Const HKEY_CURRENT_USER   = &H80000001
Const HKEY_LOCAL_MACHINE  = &H80000002
Const HKEY_USERS          = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005

Const REG_SZ        = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY    = 3
Const REG_DWORD     = 4
Const REG_MULTI_SZ  = 7


Dim SessionName: SessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
if ( SessionName = "%SESSIONNAME%" ) then
	Dim arrSubkeys
	Dim counter
	objReg.EnumKey HKEY_CURRENT_USER, "Volatile Environment", arrSubKeys
	If IsArray(arrSubKeys) then
		if Ubound(arrSubKeys)>0 Then
			counter=arrSubKeys(0)
			objReg.GetStringValue HKEY_CURRENT_USER, "Volatile Environment\" & counter, "SESSIONNAME", SessionName
			SessionName=SessionName & " "
		End If
	End If
End if


function Max(a,b)
    Max = a
    If b > a then Max = b
end function

function Min(a,b)
    Min = a
    If b < a then Min = b
end function


Sub LogMsg(ByVal logtext)

	dim logType
	logType="undefined"
	on error resume next
	logType= TypeName(logFile)
	on error goto 0

	if logType = "undefined" then
		exit sub
	end if
	if (isObject(logFile)) then
		'���� ���� ��� ���� - ������� � ����
		on error resume next
		logFile.Write(logtext)
		on error goto 0
	else
		Dim logFileObj
		on error resume next
		Set logFileObj = objFSO.OpenTextFile(logFile, 8, True)
		logFileObj.Write(logtext)
		logFileObj.close
		on error goto 0
	end if
End Sub


'������ ����� � ������� (���� �� � �������) � � ������� (���� �� ��������)
Sub Msg(ByVal text)
	dim logtext

	if (text="") then '���� ������ ����� �� � ��� ����� ������ ������ ������ ��� ����/�������
		logtext=""
	elseif (text="-") then '����������� (� � ������� � � ���)
		text="-----------------------------------------------------------------------------------"
		logtext=text
	else '�� ��������� � ��� ��������� ��������
		logtext=Date&" "&Time&" "&text
	end if

	If LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then 
		'���� �� �������� � ���������� ������ - ������� � �������
		wscript.echo(text)
	End if

	LogMsg logtext & vbCrLf
End Sub


'������ ����� � ������� (���� �� � �������) � � ������� (���� �� ��������)
'��� �������� ������
Sub Msg_(ByVal text)

	If LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then 
		'���� �� �������� � ���������� ������ - ������� � �������
		wscript.stdout.write(text)
	End if

	LogMsg Date&" "&Time&" "&text
End Sub

'������ ����� � ������� (���� �� � �������) � � ������� (���� �� ��������)
'��� �������� ������ � ��������� ������� � ���
Sub Msg__(ByVal text)
	If LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then 
		'���� �� �������� � ���������� ������ - ������� � �������
		wscript.stdout.write(text)
	End if

	LogMsg text
End Sub

'������ ����� � ������� (���� �� � �������) � � ������� (���� �� ��������)
'� ��������� ������, �� ��� ��������� ������� � ���
Sub Msg_n(ByVal text)
	If LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then 
		'���� �� �������� � ���������� ������ - ������� � �������
		wscript.echo(text)
	End if

	LogMsg text & vbCrLf
End Sub


'if debugmode=1 the writes dubug info to the specified
'file and if running under cscript also writes it to screen.
Sub DebugMsg(strDebugInfo)
	if not DEBUGMODE=1 then exit sub	
	Msg "[debug]: "& strDebugInfo
End Sub 

Sub DebugMsg_(strDebugInfo)
	if not DEBUGMODE=1 then exit sub	
	Msg_ "[debug]: "& strDebugInfo
End Sub 

Sub DebugMsg__(strDebugInfo)
	if not DEBUGMODE=1 then exit sub	
	Msg__ "[debug]: "& strDebugInfo
End Sub 

Sub DebugMsg_n(strDebugInfo)
	if not DEBUGMODE=1 then exit sub	
	Msg_n "[debug]: "& strDebugInfo
End Sub 


Sub MsgIf (ByVal Text, ByVal Condition)
	if not Condition Then exit sub
	Msg Text
End Sub

Sub MsgIf_ (ByVal Text, ByVal Condition)
	if not Condition Then exit sub
	Msg_ Text
End Sub

Sub MsgIf__ (ByVal Text, ByVal Condition)
	if not Condition Then exit sub
	Msg__ Text
End Sub

Sub MsgIf_n (ByVal Text, ByVal Condition)
	if not Condition Then exit sub
	Msg_n Text
End Sub


'this sub forces execution under cscript
'it can be useful for debugging if your machine's
'default script engine is set to wscript
Sub ForceCScript
	dim strCurrScriptHost, strExecCmdLine
	strCurrScriptHost=lcase(right(wscript.fullname,len(wscript.fullname)-len(wscript.path)-1))
	if strCurrScriptHost<>"cscript.exe" then
		strExecCmdLine=wscript.path & "\cscript.exe //nologo " & objfso.getfile(wscript.scriptfullname).shortpath
		wshShell.run strExecCmdLine & getArgsStr,1,false
		wscript.quit
	end if
End Sub


'���������� ������ str �������� symbol �� ����� maxlen
function stringPrependTo (str,symbol,maxLen)
	dim testString
	testString = str
	do while (Len(testString)<maxLen)
		testString=symbol+testString
	loop
	stringPrependTo=testString
end function


'allows for a pause at the end of execution
'currently used only for debugging
Sub Pause
	wscript.stdout.write "Press [ENTER] to continue..."
	wscript.stdin.readline
end Sub

Sub enterToExit
	wscript.stdout.write "Press [ENTER] to exit..."
	wscript.stdin.readline
	wscript.Quit
end Sub

'������� ����������
Function unset(ByRef val)
    If isObject(val) Then
        set val = Nothing
    Else
        val = null
    End If
End Function

'������� ��������� � �������
Sub Halt(ByVal text)
	Msg("HALT: "&text)
	WScript.Quit(10)
End Sub

'������� ��������� � ������� ���� ��������� �������
Sub HaltIf(ByVal condition,ByVal text)
	if (condition) then
		Halt(text)
	end if
End Sub


'������� ��������� � ������� ���� ������ ����������
Sub HaltTextIfError(ByVal text)
	If Err.Number <> 0 Then 
		Halt text & vbCrLf &_ 
			"Err code: " & Err.Number & vbCrLf &_ 
			"Description: " & Err.Description & vbCrLf &_ 
			"Source: " & Err.Source 
	end if
End Sub

'������� ��������� ���� ������ ����������
Sub HaltIfError()
	HaltTextIfError "Runtime error!"
End Sub


'������� ��������� ���� ������ ����������
Sub MsgIfError()
	If Err.Number <> 0 Then 
		Msg "ERR: Runtime error!" & vbCrLf &_ 
			"Err code: " & Err.Number & vbCrLf &_ 
			"Description: " & Err.Description & vbCrLf &_ 
			"Source: " & Err.Source 
	end if
End Sub

function getOsCaption()
	dim objWMIService : Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	dim oss : Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	dim os
	For Each os in oss
	    getOsCaption=os.Caption
	next
	if (Platform="AMD64") then
		getOsCaption=getOsCaption & " (x64)"
	end if
end function


'������ ������� ��������� � ���������� ���� ������
Sub safeRun(ByVal cmd)
	msg_ "Running: " & cmd
	on error resume next
	dim ret : ret=wshShell.run(cmd,1,true)
	msg_n " - return code: "&ret
	on error goto 0
End sub

'������ ������� ��������� � ���������� ���� ������
Sub safeRunSilent(ByVal cmd)
	debugMsg "Running: " & cmd
	on error resume next
	dim ret : ret=wshShell.run(cmd,0,true)
	debugMsg " - return code: "&ret
	on error goto 0
End sub

Sub safeFork(ByVal cmd)
	msg_ "Running: " & cmd
	on error resume next
	dim ret : ret=wshShell.run(cmd,1,false)
	msg_n " - return code: "&ret
	on error goto 0
End sub

Sub safeExec(ByVal cmd, ByVal params, ByVal path)
	msg "Executing: " & cmd&" "&params &" @"& path
	on error resume next
	wshShell.ShellExecute cmd, params, path, "runas", 1
	on error goto 0
End sub

function exitCode(ByVal cmd)
	msg_ "Running: " & cmd
	on error resume next
	dim ret : ret=wshShell.run(cmd,1,true)
	msg_n " - return code: "&ret
	on error goto 0
	exitCode = ret
End function

function execStdout(ByVal cmd)
	debugMsg "Running: " & cmd
	on error resume next
	dim ret : set ret=wshShell.exec(cmd)
	on error goto 0
	execStdout = ret.StdOut.ReadAll()
End function


'��������� ���� � ������� ���� � ��� ���� �������
function quotePath(ByVal Path)
	quotePath=Path
	if (len(Path)>0 and (not left(Path,1) = """") and (InStr(1, Path, " ", vbTextCompare)>0)) then
		quotePath=""""&Path&""""
	end if
end function

'������� ������� ���� ��� ����
function unquotePath(ByVal Path)
	unquotePath=Path
	if (len(Path)>2) and (left(Path,1) = """") and (right(Path,1)="""") then
		unquotePath=mid(Path,2,Len(Path)-2)
	end if
end function

'��������� ��� ������ ��������
Function IsArrayDimmed(arr)
	IsArrayDimmed = False
	If IsArray(arr) Then
		On Error Resume Next
		Dim ub : ub = UBound(arr)
		If (Err.Number = 0) And (ub >= 0) Then IsArrayDimmed = True
	End If
End Function

'������� �� ����� ������� � ������
Function TrimWithTabs(trimme)
	dim lead,tail
	lead=false
	tail=false
    Do Until lead
    	If Left(trimme, 1) = Chr(32) Or Left(trimme, 1) = Chr(9) then
    		trimme = Right(trimme, Len(trimme) - 1)
    	Else
    		lead = true
    	End If
    Loop
    Do Until tail
    	If Right(trimme, 1) = Chr(32) Or Right(trimme, 1) = Chr(9) then
    		trimme = Left(trimme, Len(trimme) - 1)
    	Else
    		tail = true
    	End If
    Loop
    TrimWithTabs = trimme
End Function

'CLI ARGUMENTS ROUTINE -------------------------------------------------
function argName(ByVal argument)
'���������� ��� ��������� �� ���� ��������:��������
	dim tokens
	tokens=Split(argument,":")
	argName=LCase(tokens(0))
	'msg "argName: Return "& LCase(tokens(0)) & " from " & argument
end function


function argVal(ByVal argument)
'���������� ��� ��������� �� ���� ��������:��������
	dim tokens
	tokens=Split(argument,":")
	if (Ubound(tokens)<2) then
		argVal=true
	else
		argVal=tokens(1)
	end if
end function


function arg(ByVal needle)
'������ ����������
'���� ������ ������� ���������� ����� ���������� ����������, �� ���������� true
'���� � ��� ���� ������� ������� �� ���������� ��������
'����� false
	arg=false
	needle=lcase(needle)
	'msg "Searching " & needle & " ... "
	if (WScript.Arguments.Count=0) then
		exit function
	end if
	dim i
	for i = 0 to WScript.Arguments.Count-1
		if (argName(WScript.Arguments(i)) = needle) then
			arg=argVal(WScript.Arguments(i))
			exit Function
		end if
	next
	arg=false
end function

'������� ������ ����������, ����������� ������� glue
function argList(ByVal glue)
	dim i,list
	list=""

	for i = 0 to WScript.Arguments.Count-1
		if (i>0) then 
			list=list&glue
		end if
		list=list&WScript.Arguments(i)
	next
	argList=list
end function

'�������� ���������� �����
Function GetFile(ByVal FileName)
	'default
	GetFile = ""
	if (objFSO.FileExists(FileName)) then
		dim f : set f=objFSO.OpenTextFile(FileName,1) '1=ForReading
		'��������� ��� ��������� ����� �� ��������� � ��� ����� 
		'����� ��� ������ ������� ����� ����� �������� ������ "input past end of file"
		If Not f.AtEndOfStream Then 
			GetFile = f.ReadAll
		end if
		f.close
	end if
End Function

'���������� ���������� �����
Function WriteFile(ByVal FileName, ByVal Contents)
	On Error Resume Next
 	WriteFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 2, True).Write(Contents)
	HaltTextIfError "Error writing file " & FileName
	On Error Goto 0
End Function

'�������� ������������� ���������� �����
Function GetIntFile(ByVal FileName)
	GetIntFile = 0
	if (objFSO.FileExists(FileName)) then
		dim strData
		On Error Resume Next
			strData=objFSO.OpenTextFile(FileName).ReadLine
		On Error Goto 0		
		if err.number<>0 then
			Msg "Error while Reading " & FileName
			exit function
		end if

		On Error Resume Next
			GetIntFile=CLng(Trim(strData))	
		On Error Goto 0		
		if err.number<>0 then
			Msg "Error while parsing integer [" & strData & "]"
			exit function
		end if
		
	else    
		Msg(FileName & " not found")
	end if
End Function


'������ ����������
Function getArgsStr()
	getArgsStr=""
	if (WScript.Arguments.Count=0) then
		exit function
	end if
	dim i
	redim arrArgs(WScript.Arguments.Count-1)
	For i = 0 To WScript.Arguments.Count-1
		arrArgs(i) = WScript.Arguments(i)
	Next
	getArgsStr=" " & join(arrArgs," ")
End Function






'REGISTRY ROUTINE ------------------------------------------------------
'������ ������
Function regRead(ByVal Path)
	debugMsg "Reading " & Path & " ... "
	on error resume next
	RegRead = WshShell.RegRead (Path)
	if err.number<>0 then
		debugMsg "Error while Reading " & Path
		regRead = false
	end if
	on error goto 0
End Function

'����� � ������
sub regWrite (ByVal Path, ByVal varType, ByVal varVal)
	debugMsg "Writing " & Path & "=" & varVal & "(" & varType & ") ... "
	on error resume next
	WshShell.RegWrite Path, varVal, varType
	if err.number<>0 then
		debugMsg "Error while Writing " & Path
		regWrite = false
	end if
	on error goto 0
End Sub

'������� ���� � �������
sub regDelete (ByVal Path)
	debugMsg "Deleting " & Path & " ... "
	on error resume next
	WshShell.RegDelete Path
	if err.number<>0 then
		debugMsg "Error while Deleting " & Path & vbCrLf &_
			"Err code: " & Err.Number & vbCrLf &_ 
			"Description: " & Err.Description & vbCrLf &_ 
			"Source: " & Err.Source 
		regDelete = false
	end if
	on error goto 0
End Sub

'������� ���������� � ����������� ����������� � ������������ ������
sub regCheck (ByVal Path, ByVal varType, ByVal varVal)
	Dim tmp : tmp=regRead(Path)
	if tmp = varVal Then
		Msg tmp & " already set"
	else
		Msg "Got " &tmp& " instead of " & varVal
		if (varVal<>False) then
			regWrite Path, varType, varVal
		else
			regDelete Path
		end if
	end if
End Sub

'��������� ����� �� ������� ��������
function regExists (ByVal strKey)
	dim ssig: ssig="Unable to open registry key"
	on error resume next
	debugMsg "Searchin "&strKey
	dim present: present = WshShell.RegRead(strKey)
	if err.number<>0 then
		debugMsg "Got some error on "&strKey
	    	if right(strKey,1)="\" then    'strKey is a registry key
        		if instr(1,err.description,ssig,1)<>0 then
					regExists=true
        		else
            		regExists=false
        		end if
    		else    'strKey is a registry valuename
        		regExists=false
    		end if
    	err.clear
	else
    	regExists=true
	end if
	on error goto 0
end function


Sub regCleanFolder(hive, path)
	Msg "Cleaning reg folder " & hive & "," & path & "..."
	dim oReg, arrSubKeys, subkey
	Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	'Msg "1"
	oReg.EnumKey hive, path, arrSubKeys
	'Msg "2"
  	If Not IsNull(arrSubKeys) Then
    		For Each subkey In arrSubKeys
				Msg "Deleting reg folder " & path & "\" & subkey & "..."
      			oReg.DeleteKey hive, path & "\" & subkey
    		Next
	else
		Msg "Empty"
  	End If
End Sub


Sub regDeleteRecursive(RegPath)
	'������� ������������� �����
	if not (right(RegPath,1) = "\") then
		regPath=regPath & "\"
	end if

	if (not regExists (RegPath)) then
		Msg "Folder " & RegPath & " not exist (no need to delete)"
		exit sub
	end if
	
	Msg "Deleting reg folder " & RegPath & "..."
	dim arrSubKeys, subkey
	arrSubkeys=RegEnumKeys(RegPath)
  	If Not IsNull(arrSubKeys) Then
    		For Each subkey In arrSubKeys
				'Msg "Deleting reg folder " & path & "\" & subkey & "..."
      			call regDeleteRecursive(RegPath & subkey & "\")
    		Next
	else
		Msg "No subfolders"
  	End If
	regDelete RegPath
End Sub

'simple function to provide an
'easier interface to the wmi registry functions
Function RegEnumKeys(RegKey)
	dim hive, strKeyPath, arrSubKeys
	hive=SetHive(RegKey)
	strKeyPath = right(RegKey,len(RegKey)-instr(RegKey,"\"))
	objReg.EnumKey Hive, strKeyPath, arrSubKeys
	RegEnumKeys=arrSubKeys
End Function

'simple function to provide an
'easier interface to the wmi registry functions
Function RegGetStringValue(RegKey,RegValueName)
	dim hive, strKeyPath, RegValue
	hive=SetHive(RegKey)
	strKeyPath = right(RegKey,len(RegKey)-instr(RegKey,"\"))
	tmpreturn=objReg.GetStringValue(Hive, strKeyPath, RegValueName, RegValue)
	if tmpreturn=0 then
		RegGetStringValue=RegValue
	else
		RegGetStringValue="~{{<ERROR>}}~"
	end if
End Function

'simple function to provide an
'easier interface to the wmi registry functions
Function RegGetMultiStringValue(RegKey,RegValueName)
	dim hive, strKeyPath, RegValue, tmpreturn
	hive=SetHive(RegKey)
	strKeyPath = right(RegKey,len(RegKey)-instr(RegKey,"\"))
	tmpreturn=objReg.GetMultiStringValue(Hive, strKeyPath, RegValueName, RegValue)
	if tmpreturn=0 then
		RegGetMultiStringValue=RegValue
	else
		RegGetMultiStringValue="~{{<ERROR>}}~"
	end if
End Function

'simple function to provide an
'easier interface to the wmi registry functions
Function RegGetBinaryValue(RegKey,RegValueName)
	dim hive, strKeyPath, RegValue, tmpreturn
	hive=SetHive(RegKey)
	strKeyPath = right(RegKey,len(RegKey)-instr(RegKey,"\"))
	tmpreturn=objReg.GetBinaryValue(Hive, strKeyPath, RegValueName, RegValue)
	if tmpreturn=0 then
		RegGetBinaryValue=RegValue
	else
		RegGetBinaryValue="~{{<ERROR>}}~"
	end if
End Function


'function to parse the specified hive
'from the registry functions above
'to all the other registry functions (regenumkeys, reggetstringvalue, etc...)
Function SetHive(RegKey)
	dim strHive
	strHive=left(RegKey,instr(RegKey,"\"))
	if strHive="HKCR\" or strHive="HKR\" or strHive="HKEY_CLASSES_ROOT\" then SetHive=HKEY_CLASSES_ROOT
	if strHive="HKCU\" or strHive="HKEY_CURRENT_USER\" then SetHive=HKEY_CURRENT_USER
	if strHive="HKCC\" or strHive="HKEY_CURRENT_CONFIG\" then SetHive=HKEY_CURRENT_CONFIG
	if strHive="HKLM\" or strHive="HKEY_LOCAL_MACHINE\" then SetHive=HKEY_LOCAL_MACHINE
	if strHive="HKU\"  or strHive="HKEY_USERS\" then SetHive=HKEY_USERS
End Function


'USER RIGHTS ROUTINE -----------------------------------------------------------------

'�������� ��������� ���������� ����� whoami
'http://stackoverflow.com/questions/1599567/vbscript-check-if-the-script-has-administrative-permissions
Function UserPerms (PermissionQuery)
	UserPerms = False  ' False unless proven otherwise
	Dim CheckFor, CmdToRun

	Select Case Ucase(PermissionQuery)
	'Setup aliases here
	Case "ELEVATED"
		CheckFor =  "S-1-16-12288"
	Case "SYSTEM"
		CheckFor =  LocalSystemSID
	Case "ADMIN"
		CheckFor =  LocalAdminsSID
	Case "ADMINISTRATOR"
		CheckFor =  LocalAdminsSID
	Case Else
		CheckFor = PermissionQuery
	End Select

	CmdToRun = "%comspec% /c %systemroot%\system32\whoami.exe /all | %systemroot%\system32\findstr /I /C:""" & CheckFor & """"

	DebugMsg ("Checking " & PermissionQuery & " permissions ...")
	DebugMsg ("Running " & CmdToRun & " ...")
	If wshShell.Run(CmdToRun, 0, true) = 0 Then UserPerms = True
	DebugMsg ("Checking " & PermissionQuery & " permissions complete")
End Function

'�������� ������� �� UAC
Function UACTurnedOn ()
	If regExists("HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA") then
		if WshShell.RegRead("HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA") = 0 Then
			UACTurnedOn = false
		Else
			UACTurnedOn = true
		End If
	else
		Msg err.Number
		UACTurnedOn = false
	end if
End Function

'� ������� ��������� �������: �� ����� SYSTEM ������������� ELEVATED
'����� �� �����, � �� �����-������ ����� ������ ���� ������� UAC
function checkFullAdminRights()
	Msg "Checking admin permissions ..."
	If (UserPerms("System")) Then
		Msg " - System"
	Elseif (UserPerms("Admin")) Then
		Msg " - Admin"
		if UACTurnedOn then
			Msg "INIT: UAC is turned ON: Need Admin/Elevated permissions"
			if (UserPerms("Elevated")) Then
				Msg " - Elevated"
			Else
				Halt("ERROR: got perm: Non-Elevated. ")
			End if
		Else
			Msg "INIT: UAC is turned off"
		End if
	Else
		Halt("ERROR: got perm:Non-Admin. Need:Admin")
	End if
end function


'����������� ������� � ����� ����������, ��� ������� ����� �������� � ��������
'� ���������� UAC. ���� � ���, ��� ����� ������ ����������� � �����������������
'�������� � ������� � �������, � ��� ������������ � ����������������� ��������
'����� � �������� ����� �� �����. ��� ������� ����� ������� ������ ������� ���������
'� ������������������� ��������
function launchPad (ByVal strAppPath)
	const TriggerTypeRegistration = 7
	const ActionTypeExecutable = 0
	const FlagTaskCreate = 2
	const LogonTypeInteractive = 3

	Dim strTaskName, rootFolder, service, taskDefinition, triggers, trigger, Action, NewTask
	strTaskName = "Launch_As_" & UserName & "_unelevated_" & objFSO.GetBaseName(strAppPath)
	Set service = CreateObject("Schedule.Service")
	call service.Connect()
	Set rootFolder = service.GetFolder("\")

	On Error Resume Next
		call rootFolder.DeleteTask(strTaskName, 0)
		Err.Clear
	On Error goto 0

	Set taskDefinition = service.NewTask(0)
	Set triggers = taskDefinition.Triggers
	Set trigger = triggers.Create(TriggerTypeRegistration)
	Set Action = taskDefinition.Actions.Create( ActionTypeExecutable )
	Action.Path = strAppPath

	Msg "Task definition created. About to submit the task..."
	Msg "> " & strTaskName & ", taskDefinition, " & FlagTaskCreate & ",,, "& LogonTypeInteractive

	call rootFolder.RegisterTaskDefinition(strTaskName, taskDefinition, FlagTaskCreate,,,LogonTypeInteractive,NewTask)

	if isObject(NewTask) then
		Msg "Task submitted."
	else
		Msg "Err. submiting task."
	end if
end function

'��������� ������� ������ �� �������������������� ������������, ���� ������� ������� �� �����
sub unPrivelegeMe()
	if UACTurnedOn then
		if (UserPerms("Elevated")) Then
			'��� ���� ���������� NOLAUNCHPAD � ���������� ��� ������� ����������� ��������
			launchPad Wscript.ScriptFullName & getArgsStr & " unprivelege_me_forked"
			Msg Wscript.ScriptFullName
			Halt ( "Parent process exiting due to priveleged state" )
		Else
			Msg "Priveleged mode check passed - running unpriveleged process"
		End if
	else
		Msg "Priveleged mode check passed - UAC turned off"
	end if
end sub

'��������� ������� ������ �� �������������������� ������������, ���� ������� ������� �� �����
sub forceUnPrivelegeMe()
	if arg("privelege_me_forked") then
		Msg "Unpriveleged child process detected"
	else
		launchPad Wscript.ScriptFullName & getArgsStr & " unprivelege_me_forked"
		Msg Wscript.ScriptFullName&" unprivelege_me_forked"
		Halt ( "Parent process exiting due to priveleged state" )
	End if
end sub

'��������� ������� ������ �� ������������������ ������������, ���� ������� ������� �� �����
sub privelegeMe()
	Msg "Running as " & Username
	if UACTurnedOn then
		if (UserPerms("Elevated")) Then
			Msg "Priveleged mode check passed - running priveleged process"
		Else
			if (right(Username,1)="$") then
				Msg "Launchpad not working under SYSTEM user. Ignoring"
				exit sub
			end if
			Msg "Priveleged mode check failed"
			if arg("privelege_me_forked") then
				Halt ( "Child process failed to achieve elevated state" )
			else
				Msg " - forking priveleged process ..."
				dim shellCmd : shellCmd = Chr(34) & WScript.ScriptFullName & Chr(34) & getArgsStr & " privelege_me_forked"
				debugMsg "Running " & shellcmd & " ... "
				objShell.ShellExecute "cscript.exe", shellCmd , , "runas", 1 
				Halt ( "Parent process exiting due to unpriveleged state" )
			end if
		End if
	else
		Msg "Priveleged mode check passed - UAC turned off"
	end if
end sub


'�������� ������� �������� ������� � NTP ��������
function GetNtpDiff(server)
	GetNtpDiff = -1
	dim objProc : set objProc = WshShell.Exec("%SystemRoot%\System32\w32tm.exe /monitor /nowarn /computers:"&server)

	dim input: input = ""
	dim strOutput: strOutput = ""
	Do While Not objProc.StdOut.AtEndOfStream
		input = objProc.StdOut.ReadLine
		If InStr(input, "NTP") Then
			strOutput = strOutput & input
		End If
	Loop

	dim myRegExp: Set myRegExp = New RegExp
	myRegExp.IgnoreCase = True
	myRegExp.Global = True
	myRegExp.Pattern = " NTP: ([+-][0-9]+\.[0-9]+)s"
	dim myMatches: Set myMatches = myRegExp.Execute(strOutput)

	If myMatches(0).SubMatches(0) <> "" Then
		GetNtpDiff = myMatches(0).SubMatches(0)
	End If
end function

'��������� ��������� �� ����
function HostPings(host)
	if (WshShell.Run("ping -n 1 " & host, 0, True) = 0) then
		HostPings = true	
	else	
		HostPings = false
	end if
end function


'�������� �������� ��
Function GetOS    
    GetOS="UNKNOWN"
    dim objWMI: Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    dim colOS:  Set colOS = objWMI.ExecQuery("Select * from Win32_OperatingSystem")
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

'-----------------------------------------------------------------------------------
'ENVIRONMENT VARIABLES


function EnvironmentVariableName (ByVal setString)
	dim eqPos
	eqPos=instr(1,setString,"=",vbTextCompare)
	if (eqPos = 0) then
		EnvironmentVariableName=setString
	else
		EnvironmentVariableName=Left(setString,eqPos-1)
	end if
end function

'��������� ����������� �� ���������� � ������ ��������� � ������������� ���� ����� (��� �������) 
'System		� ��������� ����������_�����, 
'User		� ����������_����� ������������
'Volatile	� ���������_���������� (���� ������� ������ ����������� ���������� � ������ �����)
'Process	� ����������_����� �������� ��������
function EnvironmentVariableCorrect (ByVal Environment, ByVal varName, ByVal varVal)
	'��� ��������� ���������� � ���������� ������, ����� ���� �� ��������� ��� ���������, ������ SYSTEM
	'if (varVal<>unset_me) then
	'	regcheck "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\"&varName,"REG_SZ", varVal
	'else
	'	regcheck "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\"&varName,"REG_SZ", false
	'end if
	Dim objEnvironment,index
	Set objEnvironment = wshShell.Environment(Environment)
	EnvironmentVariableCorrect=false
	if (varVal<>unset_me) then
		Msg "Checking if " & Environment & " variable " & varName & " is set to " & varVal
		if (objEnvironment(varName) = varVal) then
			Msg " - yes"
		else
			objEnvironment(varName) = varVal
			EnvironmentVariableCorrect=true
			Msg " - No. Fixed"
		end if
	else
		Msg "Checking if " & Environment & " variable " & varName & " is unset "
		varName=UCase(varName)
		For Each index In objEnvironment
			'DebugMsg UCase(EnvironmentVariableName(index)) &" vs "& varName
			if UCase(EnvironmentVariableName(index)) = varName then
				objEnvironment.Remove(varName)
				EnvironmentVariableCorrect=true
				Msg " - No. Fixing"
				exit For
			end if
		Next 
		if (EnvironmentVariableCorrect = false) then
			Msg " - Yes"
		end if
	end if
	
end Function

sub EnvironmentVariableSet (ByVal Environment, ByVal varName, ByVal varVal)
	Dim objEnvironment
	Set objEnvironment = wshShell.Environment(Environment)
	objEnvironment(varName)=varVal
	unset(objEnvironment)
End Sub

function EnvironmentVariableGet (ByVal Environment, ByVal varName)
	Dim objEnvironment
	Set objEnvironment = wshShell.Environment(Environment)
	EnvironmentVariableGe = objEnvironment(varName)
	unset(objEnvironment)
end function

'������� ����������� ���������� � �����-���� ���������
function EnvironmentVariableUnset (ByVal Environment, ByVal varName)
	Dim objEnvironment,index
	Set objEnvironment = wshShell.Environment(Environment)
	EnvironmentVariableUnset=false
	varName=Ucase(varName)
	For Each index In objEnvironment
		if UCase(EnvironmentVariableName(index)) = varName then
			objEnvironment.Remove(varName)
			EnvironmentVariableUnset=true
			exit For
		end if
	Next 
	unset(objEnvironment)
End function


sub EnvVarCorrectNow (ByVal varName, ByVal varVal)
	call EnvVarCorrect(varName, varVal)
end sub

sub EnvVarCorrect (ByVal varName, ByVal varVal)
	if (EnvironmentVariableCorrect ("SYSTEM",varName, varVal)) then
		call EnvironmentVariableCorrect ("PROCESS",varName, varVal)
		call EnvironmentVariableCorrect ("VOLATILE",varName, varVal)
	end if
end sub

sub EnvUsrVarCorrect (ByVal varName, ByVal varVal)
	if (EnvironmentVariableCorrect ("USER",varName, varVal)) then
		call EnvironmentVariableCorrect ("PROCESS",varName, varVal)
		call EnvironmentVariableCorrect ("VOLATILE",varName, varVal)
	end if
end sub


function EnvVarCheck(ByVal varName, ByVal varVal)
	Msg "Checking environment variable " & varName & " ... "
	on error resume next
	dim current
	current = WshShell.ExpandEnvironmentStrings(varName)
	if err.number <> 0 then
		Msg "Error expanding variable " & varName
		if (varVal=unset_me) then
			EnvVarCheck=true
		else
			EnvVarCheck=false
		end if
	else
		if (LCase(current)<>LCase(varVal)) then
			EnvVarCheck=false
			Msg(varName & " set to """ & current & """ instead of """ & varVal & """")
		else
			EnvVarCheck=true
		End if
	end if
end function


function EnvPathCorrect(ByVal testPath)
'��������� ������� ����������� ���� � ���������� PATH, ��������� ���� ���
	dim dirs,found,i

	Msg "Checking path variable for " & testPath & " presence ... "
	EnvPathCorrect=false

	testPath=unquotePath(trim(testPath))
	dirs=split(EnvironmentVariableGet ("SYSTEM", "PATH"),";")
	found=false

	for i=0 to ubound(dirs)
		if UCase(trim(dirs(i)))=UCase(testPath) then
			found=true
		end if
	next

	if found then
		msg " - found"
	else
		msg " - not found. Adding"
		ReDim Preserve dirs(UBound(dirs) + 1)
		dirs(UBound(dirs)) = testPath
	end if

	if (not found) then
		msg " - saving changes..."
		call EnvironmentVariableSet ("SYSTEM","PATH", join(dirs,";"))
		call EnvironmentVariableSet ("PROCESS","PATH", join(dirs,";"))
		call EnvironmentVariableSet ("VOLATILE","PATH", join(dirs,";"))
		EnvPathCorrect=true
		msg " - done"
	end if

end function
