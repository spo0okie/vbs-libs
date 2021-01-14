'Библиотечка с маленькими полезными функциями, которые обязательно пригодятся
'без этого ну никак не обойтись
Option Explicit
Const coreLibVer="2.8"
'ver 2.8
' - isProcRunning вынесена в lib_procs
'ver 2.7
' + добавлен объект objWmi
'ver 2.6
' + добавлена переменные Platform, Systemdrive
'ver 2.5
' ! скорректирована функция launchpad, которая перепускала скрипт от непревилегированного
'   пользователя. Была ошибка работы на Win10: теперь имя создаваемой задчи включает %username%,
'   т.к. есть подозрения, что ошибка была вызвана тем, что созданная задача от одного пользователя
'   мешала созданию такой же от другого изза совпадения имен и отсутствия доступа.

'ver 2.4:
' * Функция Msg теперь обрабатывает глобальную переменную LogFile и как Обьект и как строку
'   Если обьявлен объект, то пишем прямо в него, если строка - то открываем файл, пишем и закрываем
'   Это решает проблему с уже занятым лог файлом, когда происходит перезапуск скрипта через 
'   launchpad
' + Добавлен объект objReg для работы с реестром через WMI (также можно работать ч-з wshShell)

Dim WshShell : Set WshShell = WScript.CreateObject("WScript.Shell")
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objReg : Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
Dim objWmi : Set objWmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Dim WorkDir : WorkDir =	WshShell.ExpandEnvironmentStrings("%TEMP%") & "\"
Dim WindowsDir : WindowsDir = objFSO.GetSpecialFolder(0)


Dim ComputerName : ComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
Dim UserName : UserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
Dim UserProfile : UserProfile = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
Dim Platform : Platform = wshShell.ExpandEnvironmentStrings( "%PROCESSOR_ARCHITECTURE%" )
Dim SystemDrive : SystemDrive = wshShell.ExpandEnvironmentStrings( "%SYSTEMDRIVE%" )

Dim DEBUGMODE : DEBUGMODE = 0

'
' Хорошо известные группы:
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
	If Not IsNull(arrSubKeys) Then
		counter=arrSubKeys(0)
		objReg.GetStringValue HKEY_CURRENT_USER, "Volatile Environment\" & counter, "SESSIONNAME", SessionName
		SessionName=SessionName & " "
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


'делает вывод в консоль (если мы в консоли) и в логфайл (если он объявлен)
Sub Msg(ByVal text)
	dim logtext
	if (text="") then 'если строка пуста то в лог пишем просто пустую строку без даты/времени
		logtext=""
	elseif (text="-") then 'разделитель (и в консоль и в лог)
		text="-----------------------------------------------------------------------------------"
		logtext=text
	else 'по умолчанию в лог добавляем префиксы
		logtext=Date&" "&Time&" "&text
	end if

	If LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then 
		'если мы работаем в консольном режиме - выводим в консоль
		wscript.echo(text)
	End if

	if (isObject(logFile)) then
		'если есть лог файл - выводим в него
		on error resume next
		logFile.WriteLine(logtext)
		on error goto 0
	elseif not isnull (logFile) then
		Dim logFileObj
		on error resume next
		Set logFileObj = objFSO.OpenTextFile(logFile, 8, True)
		logFileObj.WriteLine(logtext)
		logFileObj.close
		on error goto 0
	end if
End Sub

'if debugmode=1 the writes dubug info to the specified
'file and if running under cscript also writes it to screen.
Sub DebugMsg(strDebugInfo)
	if not DEBUGMODE=1 then exit sub
	
	Msg "[debug]: "& strDebugInfo
End Sub 



'this sub forces execution under cscript
'it can be useful for debugging if your machine's
'default script engine is set to wscript
Sub ForceCScript
	strCurrScriptHost=lcase(right(wscript.fullname,len(wscript.fullname)-len(wscript.path)-1))
	if strCurrScriptHost<>"cscript.exe" then
		set objFSO=CreateObject("Scripting.FileSystemObject")
		Set objShell = CreateObject("WScript.Shell")
		Set objArgs = WScript.Arguments
		strExecCmdLine=wscript.path & "\cscript.exe //nologo " & objfso.getfile(wscript.scriptfullname).shortpath
		For argctr = 0 to objArgs.Count - 1
			strExecArg=objArgs(argctr)
			if instr(strExecArg," ")>0 then strExecArg=chr(34) & strExecArg & chr(34)
			strExecAllArgs=strExecAllArgs & " " & strExecArg
		Next
		objShell.run strExecCmdLine & strExecAllArgs,1,false
		set objFSO = nothing
		Set objShell = nothing
		Set objArgs = nothing
		wscript.quit
	end if
End Sub

'allows for a pause at the end of execution
'currently used only for debugging
Sub Pause
	set objStdin=wscript.stdin
	set objStdout=wscript.stdout
	objStdout.write "Press ENTER to continue..."
	strtmp=objStdin.readline
end Sub



'останов программы с ошибкой
Sub Halt(ByVal text)
	Msg(text)
	WScript.Quit()
End Sub


'останов программы если ошибка выполнения
Sub HaltIfError()
	If Err.Number <> 0 Then 
		Halt "HALT: Runtime error!" & vbCrLf &_ 
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


'запуск внешней программы с обработкой кода выхода
Sub safeRun(ByVal cmd)
	msg "Running: " & cmd
	on error resume next
	dim ret : ret=wshShell.run(cmd,1,true)
	msg " - return code: "&ret
	on error goto 0
End sub

Sub safeFork(ByVal cmd)
	msg "Running: " & cmd
	on error resume next
	dim ret : ret=wshShell.run(cmd,1,false)
	msg " - return code: "&ret
	on error goto 0
End sub

Sub safeExec(ByVal cmd, ByVal params, ByVal path)
	msg "Executing: " & cmd&" "&params &" @"& path
	on error resume next
	wshShell.ShellExecute cmd, params, path, "runas", 1
	on error goto 0
End sub

function exitCode(ByVal cmd)
	msg "Running: " & cmd
	on error resume next
	dim ret : ret=wshShell.run(cmd,1,true)
	msg " - return code: "&ret
	on error goto 0
	exitCode = ret
End function

'заключает путь в кавычки если в нем есть пробелы
function quotePath(ByVal Path)
	quotePath=Path
	if (len(Path)>0 and (not left(Path,1) = """") and (InStr(1, Path, " ", vbTextCompare)>0)) then
		quotePath=""""&Path&""""
	end if
end function

'убирает ковычки если они есть
function unquotePath(ByVal Path)
	unquotePath=Path
	if (len(Path)>2) and (left(Path,1) = """") and (right(Path,1)="""") then
		unquotePath=mid(Path,2,Len(Path)-2)
	end if
end function

'проверяет что массив объявлен
Function IsArrayDimmed(arr)
	IsArrayDimmed = False
	If IsArray(arr) Then
		On Error Resume Next
		Dim ub : ub = UBound(arr)
		If (Err.Number = 0) And (ub >= 0) Then IsArrayDimmed = True
	End If
End Function

'убирает по краям пробелы с табами
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
'возвращает имя аргумента из пары аргумент:значение
	dim tokens
	tokens=Split(argument,":")
	argName=tokens(0)
	'msg "argName: Return "& LCase(tokens(0)) & " from " & argument
end function


function argVal(ByVal argument)
'возвращает имя аргумента из пары аргумент:значение
	dim tokens
	tokens=Split(argument,":")
	if (Ubound(tokens)<2) then
		argVal=true
	else
		argVal=tokens(1)
	end if
end function


function arg(ByVal needle)
'парсер аргументов
'если просто находит переменную среди переданных параметров, то возвращает true
'если у нее есть какоето знчение то возвращает значение
'иначе false
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


'получить содержимое файла
Function GetFile(ByVal FileName)
	if (objFSO.FileExists(FileName)) then
		GetFile = objFSO.OpenTextFile(FileName).ReadAll
	else
		GetFile = ""
	end if
End Function

'переписать содержимое файла
Function WriteFile(ByVal FileName, ByVal Contents)
	On Error Resume Next
 	WriteFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 2, True).Write(Contents)
	HaltIfError
	On Error Goto 0
End Function









'REGISTRY ROUTINE ------------------------------------------------------
'читает реестр
Function regRead(ByVal Path)
	Msg "Reading " & Path & " ... "
	on error resume next
	RegRead = WshShell.RegRead (Path)
	if err.number<>0 then
		Msg "Error while Reading " & Path
		RegRead = false
	end if
	on error goto 0
End Function

'пишет в реестр
sub regWrite (ByVal Path, ByVal varType, ByVal varVal)
	Msg "Writing " & Path & "=" & varVal & "(" & varType & ") ... "
	WshShell.RegWrite Path, varVal, varType
End Sub

'удаляет путь в реестре
sub regDelete (ByVal Path)
	Msg "Deleting " & Path & " ... "
	WshShell.RegDelete Path
End Sub

'сверяет содержимое с переданными параметрами и корректирует реестр
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

'проверяет ветку на наличие значения
function regExists (ByVal strKey)
	dim ssig: ssig="Unable to open registry key"
	on error resume next
	Msg "Searchin "&strKey
	dim present: present = WshShell.RegRead(strKey)
	if err.number<>0 then
		Msg "Got some error on "&strKey
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
	if strHive="HKCR\" or strHive="HKR\" then SetHive=HKEY_CLASSES_ROOT
	if strHive="HKCU\" then SetHive=HKEY_CURRENT_USER
	if strHive="HKCC\" then SetHive=HKEY_CURRENT_CONFIG
	if strHive="HKLM\" then SetHive=HKEY_LOCAL_MACHINE
	if strHive="HKU\" then SetHive=HKEY_USERS
End Function


'USER RIGHTS ROUTINE -----------------------------------------------------------------

'проверка админских привелегий через whoami
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

	If WshShell.Run(CmdToRun, 0, true) = 0 Then UserPerms = True
End Function

'проверка включен ли UAC
Function UACTurnedOn ()
	If regExists("HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA") then
		if WshShell.RegRead("HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA") = 0 Then
			UACTurnedOn = false
		Else
			UACTurnedOn = true
		End If
	else
		wscript.echo err.Number
		UACTurnedOn = false
	end if
End Function

'С правами следующая история: от имени SYSTEM дополнительно ELEVATED
'права не нужны, а от юзера-админа нужны только если включен UAC
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


'Стандартная функция с сайта майкрософт, для запуска логон скриптов в системах
'с включенным UAC. Суть в том, что логон скрипт запускается в привелигированном
'процессе а десктом в обычном, и все примапленные в привелигерованном процессе
'диски в десктопе потом не видны. Эта функция через шедулер пущает внешнюю программу
'в непривелигированном процессе
function launchPad (ByVal strAppPath)
	const TriggerTypeRegistration = 7
	const ActionTypeExecutable = 0
	const FlagTaskCreate = 2
	const LogonTypeInteractive = 3
	Dim strTaskName
	strTaskName = "Launch_As_" & UserName & "_unelevated_"  & objFSO.GetBaseName(strAppPath)
	Dim service
	Set service = CreateObject("Schedule.Service")
	call service.Connect()
	Dim rootFolder
	Set rootFolder = service.GetFolder("\")

	On Error Resume Next
		call rootFolder.DeleteTask(strTaskName, 0)
	Err.Clear
	On Error goto 0

	Dim taskDefinition
	Set taskDefinition = service.NewTask(0)

	Dim triggers
	Set triggers = taskDefinition.Triggers

	Dim trigger
	Set trigger = triggers.Create(TriggerTypeRegistration)

	Dim Action
	Set Action = taskDefinition.Actions.Create( ActionTypeExecutable )
	Action.Path = strAppPath

	Msg "Task definition created. About to submit the task..."
	Msg "> " & strTaskName & ", taskDefinition, " & FlagTaskCreate & ",,, "& LogonTypeInteractive

	Dim NewTask
	
	call rootFolder.RegisterTaskDefinition(strTaskName, taskDefinition, FlagTaskCreate,,,LogonTypeInteractive,NewTask)

	if isObject(NewTask) then
		Msg "Task submitted."
	else
		Msg "Err. submiting task."
	end if
end function

'запускает текущий скрипт от непривилегированного пользователя, если текущий процесс не такой
sub unPrivelegeMe()
	if UACTurnedOn then
		if (UserPerms("Elevated")) Then
			launchPad Wscript.ScriptFullName
			Msg Wscript.ScriptFullName
			Halt ( "Parent process exiting due to priveleged state" )
		Else
			Msg "Priveleged mode check passed - running unpriveleged process"
		End if
	else
		Msg "Priveleged mode check passed - UAC turned off"
	end if
end sub

'запускает текущий скрипт от непривилегированного пользователя, если текущий процесс не такой
sub forceUnPrivelegeMe()
	if arg("privelege_me_forked") then
		Msg "Unpriveleged child process detected"
	else
		launchPad Wscript.ScriptFullName&" unprivelege_me_forked"
		Msg Wscript.ScriptFullName&" unprivelege_me_forked"
		Halt ( "Parent process exiting due to priveleged state" )
	End if
end sub

'запускает текущий скрипт от привилегированного пользователя, если текущий процесс не такой
sub privelegeMe()
	if UACTurnedOn then
		if (UserPerms("Elevated")) Then
			Msg "Priveleged mode check passed - running priveleged process"
		Else
			Msg "Priveleged mode check failed"
			if arg("privelege_me_forked") then
				Halt ( "Child process failed to achieve elevated state" )
			else
				Msg " - forking priveleged process ..."
				dim objShell: Set objShell = CreateObject("Shell.Application")
				objShell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " privelege_me_forked", , "runas", 1 
				Halt ( "Parent process exiting due to unpriveleged state" )
			end if
		End if
	else
		Msg "Priveleged mode check passed - UAC turned off"
	end if
end sub


'замеряет разницу текущего времени с NTP сервером
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

'проверяет пингуется ли хост
function HostPings(host)
	HostPings = WshShell.Run("ping -n 1 " & host, 0, True)
end function


'получить название ОС
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
