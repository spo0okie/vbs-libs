'Библиотечка с маленькими полезными функциями, которые обязательно пригодятся
'без этого ну никак не обойтись
Option Explicit
Const coreLibVer="3.0"
'v3.0  * Код работы с WMI, Переменными окружения, UAC, временем и реестром вынесен в отдельный библиотеки
'        Добавлены DebugMsgTextIfError, DebugMsgIfError, MsgTextIfError
'        Добавлена простановка флажка ошибки в случае срабатывания этих функций
'v2.14 ! launchpad переделан немного. в него теперь передается скрипт с аргументами, а wscript/cscript, а все остальное в аргументы
'v2.13 + ComputerDomain (читается из реестра также как ComputerName)
'v2.12 * regDeleteRecursive и другие функции реестра принимают корень и в полном
'v2.11 * unset_me перенесено в ядро, т.к. теперь используется и для реестра и для INI
'v2.10   computerName теперь содержит полное имя компьютера
'        старое усеченное до 15зн NETBIOS имя теперь лежит в compName
'v2.9    launchPad учитывает исходные аргументы скрипта
'v2.8    isProcRunning вынесена в lib_procs
'v2.7  + добавлен объект objWmi
'v2.6  + добавлена переменные Platform, Systemdrive
'v2.5  ! скорректирована функция launchpad, которая перепускала скрипт от непревилегированного
'        пользователя. Была ошибка работы на Win10: теперь имя создаваемой задчи включает %username%,
'        т.к. есть подозрения, что ошибка была вызвана тем, что созданная задача от одного пользователя
'        мешала созданию такой же от другого изза совпадения имен и отсутствия доступа.
'v2.4  * Функция Msg теперь обрабатывает глобальную переменную LogFile и как Обьект и как строку
'        Если обьявлен объект, то пишем прямо в него, если строка - то открываем файл, пишем и закрываем
'        Это решает проблему с уже занятым лог файлом, когда происходит перезапуск скрипта через 
'        launchpad
'      + Добавлен объект objReg для работы с реестром через WMI (также можно работать ч-з wshShell)

on error resume next
Dim wshShell          : Set wshShell      = WScript.CreateObject("WScript.Shell")
Dim objShell          : Set objShell      = CreateObject("Shell.Application")
Dim objFSO            : Set objFSO        = CreateObject("Scripting.FileSystemObject")

Dim WorkDir    : WorkDir    = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\"
Dim WindowsDir : WindowsDir = objFSO.GetSpecialFolder(0)

'Полное имя получить - не такая уже тривиальная задача. в основном все способы возвращают 15значное NETBIOS имя
Dim Computername    : Computername   = wshShell.RegRead ("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\HostName")
Dim ComputerDomain  : ComputerDomain = wshShell.RegRead ("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Domain")

Dim CompName        : CompName    = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
Dim UserName        : UserName    = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
Dim UserDomain      : UserDomain  = wshShell.ExpandEnvironmentStrings( "%USERDOMAIN%" )
Dim UserProfile     : UserProfile = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
Dim Platform        : Platform    = wshShell.ExpandEnvironmentStrings( "%PROCESSOR_ARCHITECTURE%" )
Dim SystemDrive     : SystemDrive = wshShell.ExpandEnvironmentStrings( "%SYSTEMDRIVE%" )
Dim SystemRoot      : SystemRoot  = wshShell.ExpandEnvironmentStrings( "%SYSTEMROOT%" )
on error goto 0

Dim DEBUGMODE : DEBUGMODE = 0

' CONSOLE funcs ---------------------------

Dim CONSOLEMODE
CONSOLEMODE = false
If LCase( Right( WScript.FullName, 12 ) ) = "\cscript.exe" Then CONSOLEMODE = true

'Требует выполнения скрипта через cscript
Sub ForceCScript
	if CONSOLEMODE then exit sub
	dim strExecCmdLine
	strExecCmdLine=Wscript.path & "\cscript.exe //nologo " & objfso.getfile(Wscript.ScriptFullname).Shortpath
	wshShell.run strExecCmdLine & getArgsStr,1,false
	wscript.quit
End Sub


'allows for a pause at the end of execution
'currently used only for debugging
Sub Pause
	if CONSOLEMODE then exit sub
	wscript.stdout.write "Press [ENTER] to continue..."
	wscript.stdin.readline
end Sub

Sub enterToExit
	if CONSOLEMODE then exit sub
 	wscript.stdout.write "Press [ENTER] to exit..."
	wscript.stdin.readline
	wscript.Quit
end Sub

' ELAPSED TIME funcs ----------------------
Dim scriptStartedTime : scriptStartedTime = Timer()
Function ElapsedTime ()
	ElapsedTime = FormatNumber(Timer() - scriptStartedTime, 2)
End function


Function timePrefix()
	timePrefix = Date & " " & Time & " (" & ElapsedTime & "s) "
End function


' ERR funcs -------------------------------

Dim ErrorsCount : ErrorsCount = 0

'Возвращает развернутое описание ошибки в 3 строки
Function getErrorDescr()
	getErrorDescr="Err code: " & Err.Number & vbCrLf &_ 
		"Description: " & Err.Description & vbCrLf &_ 
		"Source: " & Err.Source 
End Function

'Записывает состояние ошибки в файл-флаг (если его путь определен в виде строки)
'Увеличивает счетчик ошибок
Sub setErrFlag(ByVal status)
	if (status>0) then ErrorsCount=ErrorsCount+1

	dim logType
	logType="undefined"
	on error resume next
	logType= TypeName(errFile)
	on error goto 0

	if logType = "undefined" then exit sub

	writeFile errFile, status
End Sub

'Очищает флаг ошибок если ошибок не было
Sub okErrFlag()
	if ErrorsCount=0 then setErrFlag(0)
End Sub




' LOG funcs -------------------------------

'записать сообщение в логфайл, если он определен
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
		'если есть лог файл - выводим в него
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

'делает вывод в консоль (если мы в консоли) и в логфайл (если он объявлен)
Sub Msg(ByVal text)
	dim logtext

	if (text="") then 'если строка пуста то в лог пишем просто пустую строку без даты/времени
		logtext=""
	elseif (text="-") then 'разделитель (и в консоль и в лог)
		text="-----------------------------------------------------------------------------------"
		logtext=text
	else 'по умолчанию в лог добавляем префиксы
		logtext=timePrefix & text
	end if

	If CONSOLEMODE Then wscript.echo(text)

	LogMsg logtext & vbCrLf
End Sub

'делает вывод в консоль (если мы в консоли) и в логфайл (если он объявлен)
'Без перевода строки
Sub Msg_(ByVal text)
	If CONSOLEMODE Then wscript.stdout.write(text)
	LogMsg timePrefix & text
End Sub

'делает вывод в консоль (если мы в консоли) и в логфайл (если он объявлен)
'Без перевода строки и временной отметки в лог
Sub Msg__(ByVal text)
	If CONSOLEMODE Then wscript.stdout.write(text)
	LogMsg text
End Sub

'делает вывод в консоль (если мы в консоли) и в логфайл (если он объявлен)
'С переводом строки, но без временной отметки в лог
Sub Msg_n(ByVal text)
	If CONSOLEMODE Then wscript.echo(text)
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

'Сообщение если ошибка выполнения
Sub MsgTextIfError(ByVal Text)
	If Err.Number <> 0 Then 
		Msg "ERR: " & Text & vbCrLf & getErrorDescr
		setErrFlag 1
	end if
End Sub

'с текстом по умолчанию
Sub MsgIfError()
	MsgTextIfError "Runtime error!"
End Sub

'Отладочное сообщение если ошибка выполнения
Sub DebugMsgTextIfError(ByVal Text)
	If Err.Number <> 0 Then 
		DebugMsg "ERR: " & Text & vbCrLf & getErrorDescr
		setErrFlag 1
	end if
End Sub

'с текстом по умолчанию
Sub DebugMsgIfError()
	DebugMsgTextIfError "Runtime error!"
End Sub

'Останов программы -----

'останов программы с ошибкой
Sub Halt(ByVal text)
	Msg("HALT: "&text)
	WScript.Quit(10)
End Sub

'останов программы с ошибкой если соблюдено условие
Sub HaltIf(ByVal condition,ByVal text)
	if (condition) then
		Halt(text)
	end if
End Sub

'останов программы с ошибкой если ошибка выполнения
Sub HaltTextIfError(ByVal text)
	If Err.Number <> 0 Then 
		setErrFlag 1
		Halt "ERR: " & Text & vbCrLf & getErrorDescr
	end if
End Sub

'останов программы если ошибка выполнения
Sub HaltIfError()
	HaltTextIfError "Runtime error!"
End Sub


'номинальный останов программы
Sub Done()
	Msg ("Script complete.")
	WScript.Quit(0)
End Sub


' RUN PROC -------------------

'запуск внешней программы с обработкой кода выхода
Sub safeRun(ByVal cmd)
	msg_ "Running: " & cmd
	on error resume next
	dim ret : ret=wshShell.run(cmd,1,true)
	msg_n " - return code: "&ret
	on error goto 0
End sub

'запуск внешней программы с обработкой кода выхода
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

'CLI ARGUMENTS ROUTINE -------------------------------------------------
function argName(ByVal argument)
'возвращает имя аргумента из пары аргумент:значение
	dim tokens
	tokens=Split(argument,":")
	argName=LCase(tokens(0))
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

'вывести список аргументов, разделенных строкой glue
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


'строка аргументов
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




' FILE IO ---------------------------

'получить содержимое файла
Function GetFile(ByVal FileName)
	'default
	GetFile = ""
	if (objFSO.FileExists(FileName)) then
		dim f : set f=objFSO.OpenTextFile(FileName,1) '1=ForReading
		'проверяем что указатель файла не находится в его конце 
		'иначе при чтении пустого файла будет вылетать ошибка "input past end of file"
		If Not f.AtEndOfStream Then 
			GetFile = f.ReadAll
		end if
		f.close
	end if
End Function

'переписать содержимое файла
Function WriteFile(ByVal FileName, ByVal Contents)
	On Error Resume Next
 	WriteFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 2, True).Write(Contents)
	HaltTextIfError "Error writing file " & FileName
	On Error Goto 0
End Function

'получить целочисленное содержимое файла
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



' VAR funcs --------------------------

'удалить переменную
Function unset(ByRef val)
    If isObject(val) Then
        set val = Nothing
    Else
        val = null
    End If
End Function

'проверяет что массив объявлен
Function IsArrayDimmed(arr)
	IsArrayDimmed = False
	If IsArray(arr) Then
		On Error Resume Next
		Dim ub : ub = UBound(arr)
		If (Err.Number = 0) And (ub >= 0) Then IsArrayDimmed = True
	End If
End Function

function getVariableType(byRef var)
	getVariableType = "undefined"
	on error resume next
	getVariableType = TypeName(var)
	on error goto 0
end function

' MATH funcs -----------------------------

function Max(a,b)
    Max = a
    If b > a then Max = b
end function

function Min(a,b)
    Min = a
    If b < a then Min = b
end function


' STRING funcs ---------------------------

'предваряет строку str сиволами symbol до длины maxlen
function stringPrependTo (str,symbol,maxLen)
	dim testString
	testString = str
	do while (Len(testString)<maxLen)
		testString=symbol+testString
	loop
	stringPrependTo=testString
end function


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




' MISC funcs -----------------------------

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
	if (WshShell.Run("ping -n 1 " & host, 0, True) = 0) then
		HostPings = true	
	else	
		HostPings = false
	end if
end function

