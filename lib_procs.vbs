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

'' SIG '' Begin signature block
'' SIG '' MIIH0QYJKoZIhvcNAQcCoIIHwjCCB74CAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' 3VzY8sLSZ4PdxUq3HmGHhzTPOz1ndCLdH5G0F3R+VdKg
'' SIG '' ggUsMIIFKDCCAxCgAwIBAgIBADANBgkqhkiG9w0BAQsF
'' SIG '' ADBtMQswCQYDVQQGEwJSVTENMAsGA1UECAwEVXJhbDEU
'' SIG '' MBIGA1UEBwwLQ2hlbHlhYmluc2sxETAPBgNVBAoMCFJl
'' SIG '' dmlha2luMQswCQYDVQQLDAJJVDEZMBcGA1UEAwwQcmV2
'' SIG '' aWFraW4tcm9vdC1DQTAeFw0yMzA1MjQwNDU2NTdaFw0y
'' SIG '' NDA2MDIwNDU2NTdaMGAxCzAJBgNVBAYTAlJVMQ0wCwYD
'' SIG '' VQQIDARVcmFsMQ0wCwYDVQQHDARDaGVsMREwDwYDVQQK
'' SIG '' DAhSZXZpYWtpbjELMAkGA1UECwwCSVQxEzARBgNVBAMM
'' SIG '' CnNjcmlwdHNpZ24wggEiMA0GCSqGSIb3DQEBAQUAA4IB
'' SIG '' DwAwggEKAoIBAQDBTtnKwGde6qQttu1TOo/JIGTZ2hoa
'' SIG '' HIGDBFKgexeDT8choad2DXRQzxGyu2y9w7djwuthEODY
'' SIG '' KLVf12PcofOKnowAoSIqQ7VW77I8I4VLI7hi0VDGZ9V9
'' SIG '' W4pC/mcJjkaEMSAFj6/CST5tpeczI2KxYpM1f+mEWGiu
'' SIG '' TkB3K3jVhsaDCuWZYZoszAJkUgp3SevPyqA6JuqzwpHD
'' SIG '' aDbNG5ohd1MwcwvRKab6HNwkEprYyTiX6uWZ8rBGyIGE
'' SIG '' 4ZtshlAt6yyn6U/tYREG9+pA9CzoPHfB3gh6taeR0/25
'' SIG '' oeZ5WYHuGMNeHaHYeeIXKS9gfPh3ND/fJGQaTljVSGX5
'' SIG '' e3StAgMBAAGjgd8wgdwwHQYDVR0OBBYEFDc+8unMGviq
'' SIG '' cvfVA1vXi3LqheoJMB8GA1UdIwQYMBaAFKJJoRQ/bOk/
'' SIG '' S1B2wDmCrQ0ZzJbKMA8GA1UdEwEB/wQFMAMBAf8wDgYD
'' SIG '' VR0PAQH/BAQDAgGGMDgGA1UdHwQxMC8wLaAroCmGJ2h0
'' SIG '' dHA6Ly9wa2kucmV2aWFraW4ubmV0L2NybC9yb290LWNh
'' SIG '' LmNybDA/BggrBgEFBQcBAQQzMDEwLwYIKwYBBQUHMAKG
'' SIG '' I2h0dHA6Ly9wa2kucmV2aWFraW4vbmV0L3Jvb3QtY2Eu
'' SIG '' Y3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCyB0c0PKF0ffSX
'' SIG '' RmTBaqNWVOEpokgkdJbUNhVhKL4d7MR2wF1GX6rTeGTD
'' SIG '' hF4p1R3N6wRR0AAFVfp63st3w51XoQbJmGInJ7IFgrB2
'' SIG '' 7G6XzFVkp0llNu/1ygiqHm8v7JZEhdiqCun+JDd0ata4
'' SIG '' HKz2lca85tg2wnDfm0n3N7cdI56UkB+dKAzMLINVNT9X
'' SIG '' GSF70kXtCSPeLPDorVge0oWLxDvUiYAzlLvXk2+MTlrJ
'' SIG '' ka3R/s84X5W6CP9JJptIuzVuSd5ETB+tz/6xid2ELhNK
'' SIG '' ihkETnTViqdKp0CFGS/tRSDnfQ7Kp+Udr/SL7V/cg6Kh
'' SIG '' y8tXMCW+EJQBhrAGhudOvnIcFtTrUmhjupqMUuaLqDVY
'' SIG '' ACSwtmuihx7RAKREee0d8DJ99P3unNqfThtTPfHCzgeU
'' SIG '' Yk+i505Y8Op7G286bAwMv+m6SvnOT8vexSzJ3c77Vuyv
'' SIG '' HEU49MkgZAhpajQjTeOq2Kj3o1m+jxQ3OkWgMD6EMoJ8
'' SIG '' PIQS1XPhXcZ1N81uheeUf9EX13m32CulsDHmOnhQcT56
'' SIG '' jKt/9dn4jqHodqEodaz2Jb/tu7u6uIHmuaB2g6DTRxAO
'' SIG '' v33V/0yI40FG0SPAoNsWNVFySO5UwnewXA6H1hWEFezZ
'' SIG '' UPWWnqWb+F2uNUC8gl7Uguc2q3pJ5RhoJX+TxgBIt3oW
'' SIG '' SrZ8foMC3jGCAf0wggH5AgEBMHIwbTELMAkGA1UEBhMC
'' SIG '' UlUxDTALBgNVBAgMBFVyYWwxFDASBgNVBAcMC0NoZWx5
'' SIG '' YWJpbnNrMREwDwYDVQQKDAhSZXZpYWtpbjELMAkGA1UE
'' SIG '' CwwCSVQxGTAXBgNVBAMMEHJldmlha2luLXJvb3QtQ0EC
'' SIG '' AQAwDQYJYIZIAWUDBAIBBQCgXjAQBgorBgEEAYI3AgEM
'' SIG '' MQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAv
'' SIG '' BgkqhkiG9w0BCQQxIgQgmxTdfqJExIbZ3p7tXmewS+iO
'' SIG '' tlPMEXnZF5WWi58G9b4wDQYJKoZIhvcNAQEBBQAEggEA
'' SIG '' htfpovFUuUs031YVNfL3fDwPOAaIcn+T3aS8/lrHYwi8
'' SIG '' PoMqSN9+Z6oVfOAztQlsNxZOVAStivvkSes8CRyzSr7H
'' SIG '' ruwhsNYTq/kynJYpSxYQmhome61GLuX9Tn4QeB4Xi7sR
'' SIG '' TBAWOcJZZw8hRsYHLW5baNVNqjFETTQGdQ9ypgev4Ev/
'' SIG '' Oqem2N8xfCCpGIcsDZbbOp3A/iZqdBiDD2Vt6BQu3pti
'' SIG '' DLsS8RYXltGjg6RDXuW/Le2Z4akbEBdsPdtakChO2sUD
'' SIG '' F6EKBQg4txsU/IxcFBOY5TDspuJ7SQ8EvGolivAWUxlZ
'' SIG '' htepobwg9NP4n+wCJ6OsaIzBut+lFibIZg==
'' SIG '' End signature block
