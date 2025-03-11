'Библиотечка работы с полномочиями пользователя/процесса
Option Explicit


' Хорошо известные SID:
Const LocalSystemSID	= "S-1-5-18"
Const LocalAdminsSID	= "S-1-5-32-544"
Const LocalUsersSID	= "S-1-5-32-545"
Const ElevatedSID	= "S-1-16-12288"


'USER RIGHTS ROUTINE -----------------------------------------------------------------

'проверка админских привелегий через whoami
'http://stackoverflow.com/questions/1599567/vbscript-check-if-the-script-has-administrative-permissions
Function UserPerms (PermissionQuery)
	UserPerms = False  ' False unless proven otherwise
	Dim CheckFor, CmdToRun

	Select Case Ucase(PermissionQuery)
	'Setup aliases here
	Case "ELEVATED"
		CheckFor =  ElevatedSID
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

'проверка включен ли UAC
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
	Action.Path = WScript.FullName
	Action.Arguments = strAppPath

	Msg "Task definition created. About to submit the task..."
	Msg "> " & strTaskName & ", taskDefinition, " & FlagTaskCreate & ",,, "& LogonTypeInteractive

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
			'тут надо обработать NOLAUNCHPAD в параметрах для запрета бесконечной рекурсии
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

'запускает текущий скрипт от непривилегированного пользователя, если текущий процесс не такой
sub forceUnPrivelegeMe()
	if arg("privelege_me_forked") then
		Msg "Unpriveleged child process detected"
	else
		launchPad Wscript.ScriptFullName & getArgsStr & " unprivelege_me_forked"
		Msg Wscript.ScriptFullName&" unprivelege_me_forked"
		Halt ( "Parent process exiting due to priveleged state" )
	End if
end sub

'запускает текущий скрипт от привилегированного пользователя, если текущий процесс не такой
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


