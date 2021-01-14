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
	isUserProcRunning=False
	if not isNull(colProcesses) then
		For Each objProcess in colProcesses
			isUserProcRunning=True
			exit function
		Next
	end if
	on error goto 0
End Function

Function killProc(procName)
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
            objProcess.Terminate()
		Msg "Waiting for " & procName & " to stop ..."
        End If
    Next

    If boolRunning Then
        ' Wait and make sure the process is terminated.
        ' Routine written by Denis St-Pierre.
        Do Until Not boolRunning
            Set colProcesses = objWmi.ExecQuery( "Select * From Win32_Process Where Name = '" & procName & "'" )
            WScript.Sleep 100 'Wait for 100 MilliSeconds
            If colProcesses.Count = 0 Then 'If no more processes are running, exit loop
                boolRunning = False
            End If
        Loop
        ' Display a message
        Msg procName & " was terminated"
    Else
        Msg "Process """ & procName & """ not found"
    End If
End function