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
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' 3VzY8sLSZ4PdxUq3HmGHhzTPOz1ndCLdH5G0F3R+VdKg
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
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIJsU3X6iRMSG
'' SIG '' 2d6e7V5nsEvojrZTzBF52ReVloufBvW+MA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAHEe1vtB5eN5J+5jVgj57O0Xc8UzZbYC
'' SIG '' 2F9EzGWXgBmmDJL3fJEZ333Zz43TgXVrdYp/srP+znxv
'' SIG '' jB4A2B5xtotmABMxS2NOtjYkfdwtOQtFqO6sQ7xlR8C3
'' SIG '' bRe+Oqe/Hrcc6pu1CUOHe88/zmd/R50v7RrVaeEnWfeG
'' SIG '' K7uHT26uOkEOfCAsjHG3p0ByKxpsOyGhjxuoExB15J5o
'' SIG '' L3+dfP5PtGdMFcaC5yYsOqqSTTaLsdfZ91XP4BgcLfLs
'' SIG '' 1wBPl7Y4D5ZGTVnyLCnJ+r/SxLZQYxLxgQuGVcYLXxZA
'' SIG '' OUQzFxoow0Mv7T5yFnH8osLqEH8RFbKyHgCzJfcLVBSPukA=
'' SIG '' End signature block
