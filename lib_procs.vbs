'Библиотечка работы с процессами
Option Explicit
Const procLibVer="1.0"
'ver 1.0
' + утащил из ядра функцию isProcRunning, переименовал в isUserProcRunning

'Проверяет работает ли процесс под пользователем
Function isUserProcRunning(procName,procUser)
	dim strNameOfUser
	dim colProcesses
	dim objProcess
	dim Return
	on error resume next
	Set colProcesses = objWmi.ExecQuery("Select * from Win32_Process where name like """&procName&"""")
	isUserProcRunning=False
	if not isNull(colProcesses) then
		For Each objProcess in colProcesses
			Return = objProcess.GetOwner(strNameOfUser)
			If Return = 0 Then
				msg "Process " & objProcess.Name & " is owned by " & "\" & strNameOfUser & "."
				if strNameOfUser = procUser then
					isUserProcRunning=True
					exit function
				end if
			else 
				msg "Could not get owner info for process " & objProcess.Name & " // Error = " & Return
			end if
		Next
	end if
	on error goto 0
End Function

'Проверяет работает ли процесс под пользователем
Function isProcRunning(procName)
	dim colProcesses
	dim objProcess
	dim Return
	on error resume next
	Set colProcesses = objWmi.ExecQuery("Select * from Win32_Process where name like """&procName&"""")
	isProcRunning=False
	if not isNull(colProcesses) then
		For Each objProcess in colProcesses
			isProcRunning=True
			exit function
		Next
	end if
	on error goto 0
End Function

'Проверяет работает ли процесс с указанным PID
Function isPidProcRunning(procName)
	dim colProcesses
	dim objProcess
	dim Return
	on error resume next
	Set colProcesses = objWmi.ExecQuery("Select * from Win32_Process where ProcessID="""&procName&"""")
	isPidProcRunning=False
	if not isNull(colProcesses) then
		For Each objProcess in colProcesses
			isPidProcRunning=True
			exit function
		Next
	end if
	on error goto 0
End Function

Function killProc(byVal procName, byVal Timeout)
'Authors: Denis St-Pierre and Rob van der Woude
'Purpose: Kills a process and waits until it is truly dead
	Dim boolRunning, colProcesses, objProcess
	boolRunning = False

	Set colProcesses = objWmi.ExecQuery( "Select * From Win32_Process", , 48 )
	For Each objProcess in colProcesses
		If LCase( procName ) = LCase( objProcess.Name ) Then
		' Confirm that the process was actually running
		boolRunning = True
		' Get exact case for the actual process name
		procName  = objProcess.Name
		' Kill all instances of the process
		'иногда тут вылетает ошибка SWbemObjectEx not found
		'так и не понял что это значит. Ниже обсуждение аналогичной ошибки
		'https://social.technet.microsoft.com/Forums/ie/en-US/57d80534-0777-43e4-bbeb-1b858c79ba16/terminate-process-by-owner-on-local-or-remote-computer?forum=ITCG
		on error resume next
			objProcess.Terminate()
		on error goto 0
		Msg_ "Waiting " & Timeout & "s for " & procName & " to stop ..."
        End If
    Next

    Dim StartTime : StartTime=Timer
    Dim ReportTime: ReportTime=10
    If boolRunning Then
        ' Wait and make sure the process is terminated.
        ' Routine written by Denis St-Pierre.
        Do Until (Not boolRunning) or (Timer - StartTime) > Timeout
		Msg__ "."
            Set colProcesses = objWmi.ExecQuery( "Select * From Win32_Process Where Name = '" & procName & "'" )
            WScript.Sleep 100 'Wait for 100 MilliSeconds
            If colProcesses.Count = 0 Then 'If no more processes are running, exit loop
                boolRunning = False
            End If
	    if (Timer - StartTime) > ReportTime then
		Msg ""
		Msg_ (Timer - StartTime) & "sec passed. Stil running ..."
		ReportTime = ReportTime + 10
	    end if
        Loop
        ' Display a message
	if not boolRunning then
		Msg ""
	        Msg procName & " was terminated"
		killProc=true
	end if
	if (Timer - StartTime) > Timeout then
		Msg ""
		Msg "Can not kill " & procName & ". Timeout exceeded"
		killProc=false
	end if
    Else
        Msg "Process """ & procName & """ not found"
	killProc=true
    End If
End function

'Kill -9
Function killProc9(byVal procName, byVal Timeout)
	dim attempt
	for attempt=1 to 4
		if not killProc(procName,round(Timeout/4)) then
			Msg "Kill attempt " & attempt
			safeRun "taskkill /IM " & procName & " /F"
		else
			killProc9=true
			exit function
		end if
	next
	killProc9=killProc(procName,2)
end function

                
Function GetCurrentProcessID()
    With GetObject("winmgmts:root\cimv2:win32_process.Handle='" &_
        CreateObject("WScript.Shell").Exec("rundll32 kernel32,Sleep").ProcessId & "'")
        GetCurrentProcessID = .ParentProcessId
        .Terminate
    End With
End Function

'Записывает в файл PID текущего процесса
Function WritePidFile(ByVal FileName)
	writeFile FileName,GetCurrentProcessID
End Function

'Очищает PID файл
Function ClearPidFile(ByVal FileName)
	if objFSO.fileExists(FileName) then
		objFSO.DeleteFile(FileName)
	end if
End Function

'Проверяет наличие PID файла
'если на месте - проверяет запущенн ли процесс владелец
Function CheckPidFile(ByVal FileName)
	dim pid
	pid=GetIntFile(FileName)
	msg_ "Got pid " & pid & " from " & FileName
	if pid>0 then
		if isPidProcRunning(pid) then
			Msg_n " - It is alive"
			CheckPidFile=pid
			Exit Function
		else
			Msg_n " - It is dead"
		end if
	end if
	CheckPidFile=0	
End Function
