'Библиотека с маппингом принтеров и дисков (уже использовалась в домене RTS-DEVELOP), вроде с правками Комиссарова
'Путем нудной отладки выяснилось следующее:
'На Windows XP вызов процедуры objNetwork.AddWindowsPrinterConnection "\\chl-fsrv-open\hpLJm601n"
'вызывает ошибку "The procedure entry point sprintf_s could not be located in the dynamic link library msvcrt.dll"
'или по русски "Точка входа в процедуру sprintf_s не найдена в библиотеке DLL msvcrt.dll"
'путем дебага АПИ вызовов выяснилось что ошибка возникает на самом деле при попытке открыть файлы
'C:\WINDOWS\system32\spool\drivers\w32x86\3\PrintConfig.dll.2.Config
'C:\WINDOWS\system32\spool\drivers\w32x86\3\PrintConfig.dll.2.Mainfest
'которых в системе нет. создание пустых файлов по этом пути решило проблему. вероятно нужно избегать добавления принтеров подобным образом в XP
'
'
'

Dim objNetwork : Set objNetwork = CreateObject("WScript.Network")

Function MapDrive(ByVal strDrive, ByVal strShare, ByVal strName)
    ' Function to map network share to a drive letter.
    ' If the drive letter specified is already in use, the function
    ' attempts to remove the network connection.
    ' objFSO is the File System Object, with global scope.
    ' objNetwork is the Network Object, with global scope.
    ' Returns True if drive mapped, False otherwise.
    ' strDrive - буква диска.
    ' strShare - сетевой путь.
    ' strName - название диска, отображаемое в проводнике Windows.
    ' oShell - команда записи короткого имени, отображаемого в проводнике Windows.
    ' Введена запись ошибок подключения сетевых путей в EventLog.
    ' По умолчанию диски мапятся с параметром "/PERSISTENT:NO" (bUpdateProfile = false).
    Msg "" : Msg " -- Mapping " & strDrive & " <- " & strShare

    Dim objDrive
    Dim oShell

    Set oShell = CreateObject("Shell.Application")

    On Error Resume Next

    CheckDir(strShare) 'создаем сетевой путь, если он отсутствует

    MapDrive = True
    If (objFSO.DriveExists(strDrive) = True) Then
	Msg strDrive & " exists"
	Err.Clear
        Set objDrive = objFSO.GetDrive(strDrive)
        If (Err.Number <> 0) Then
            On Error GoTo 0
	    Msg "Error getting FS Object on " & strDrive
	    Msg "Exit"
            MapDrive = False
            Exit Function
        End If
        If (objDrive.DriveType = 3) Then
	    Msg "Unmapping " & strDrive
            objNetwork.RemoveNetworkDrive strDrive, True, True
        Else
	    Msg "Drive letter is busy with local drive " & strDrive
            MapDrive = False
	    Msg "Exit"
            Exit Function
        End If
    End If
	
    Msg "Mapping " & strShare & "..."
    objNetwork.MapNetworkDrive strDrive, strShare
    oShell.NameSpace(strDrive).Self.Name = strName
    
    If (Err.Number = 0) Then
	Msg " - OK"
    Else
        Err.Clear
	Msg " - Error"
        MapDrive = False
    End If
 
    If (objFSO.DriveExists(strDrive) = True) Then
        Msg "Checking drive " & strDrive & " - OK"
    else
        Msg "Checking drive " & strDrive & " - Err"
	MapDrive = False
    end If

    If (objFSO.FolderExists(strDrive & "\") = True) Then
        Msg "Checking path " & strDrive & "\ - OK"
    else
        Msg "Checking path " & strDrive & "\ - Err"
	MapDrive = False
    end If

    Set oDrives = objNetwork.EnumNetworkDrives

    For i = 0 to oDrives.Count - 1 Step 2
        Msg "Drive " & oDrives.Item(i) & " = " & oDrives.Item(i+1)
    Next 
'    CheckDir(strDrive & "\I_am_test-Remove_me_PLZ")
	
'    On Error GoTo 0
'	отключил дополнительные действия 2016-02-03, поскольку были проблемы с пропаданием дисков уже после подключения 
'    Select Case Err.Number
'        Case 0            ' No error.
'        Case -2147023694
'            objNetwork.RemoveNetworkDrive strDrive, True, True
'            objNetwork.MapNetworkDrive strDrive, strShare
'            oShell.NameSpace(strDrive).Self.Name = strName
'        Case -2147024811
'            objNetwork.RemoveNetworkDrive strDrive, True, True
'            objNetwork.MapNetworkDrive strDrive, strShare
'            oShell.NameSpace(strDrive).Self.Name = strName
'        Case Else
'            Msg "WARNING!!! Mapping network drive error: " & CStr(Err.Number) & " 0x" & Hex(Err.Number)
'            Msg "Error description: " & Err.Description
'            Msg "Domain: " & objNetwork.UserDomain
'            Msg "Computer Name: " & objNetwork.ComputerName
'            Msg "User Name: " & objNetwork.UserName
'            Msg "Drive name: " & strDrive
'            Msg "Map path: " & strShare
            'WshShell.LogEvent 1, Msg, FileServ00
'    End Select
    Msg "Complete"
End Function

function MapPrintersByGrp(Dict)
	Dim i,grpKeys		' переменная для определения "попадания" пользователя хотябы в одну из групп принтеров.
	MapPrintersByGrp = False

	msg ("Attaching printers...")
	grpKeys = Dict.Keys   ' Get the keys.
	For i = 0 To Dict.Count -1 ' Iterate the array.
		If (IsMember(objUser, grpKeys(i)) = True) Then
			msg ("Group "& grpKeys(i) & " found > attaching " & Dict(grpKeys(i)) & "...")
			On Error Resume Next	'нам нужно тут самим обработать возможную ошибку (поскольку вероятность ее велика)
			Err.Clear      ' Clear any possible Error that previous code raised
			'objNetwork.AddWindowsPrinterConnection Dict(grpKeys(i))
			'objNetwork.SetDefaultPrinter Dict(grpKeys(i))
			If Err.Number <> 0 Then
				msg( "Error: " & Err.Number & " /Hex: " & Hex(Err.Number) )
    				msg( "Source: " &  Err.Source )
    				msg( "Description: " &  Err.Description )
    				Err.Clear             ' Clear the Error
			else
				msg ("Prn "& Dict(grpKeys(i)) & " attached. ")
				MapPrintersByGrp = True
			End If
			On Error Goto 0           ' Don't resume on Error			
		End If
	Next
End function

'возвращает флаг того что диск виртуальный, сделанный через комманду subst
function DiskSubstituted(Disk)
	DiskSubstituted=False
	Set objExecObject = WshShell.Exec("cmd /c subst")
	Do While Not objExecObject.StdOut.AtEndOfStream
	    	strText = objExecObject.StdOut.ReadLine()
		'wscript.echo Instr(strText, Disk)
	    	If Instr(strText, Disk) = 1 Then
        		DiskSubstituted=True
			'wscript.echo "GotIt"
	        	Exit Do
	   	End If
	Loop	
end function

'возвращает UUID диска. 
'Использует утилиту mountvol, я проверял, она есть в XP и Server 2012 R2
'рассчитываю что промежуточные версии тоже ее включают
function DiskUUID(Disk)
	DiskUUID=""
	Set objExecObject = WshShell.Exec("mountvol " & Disk & " /L")
	Do While Not objExecObject.StdOut.AtEndOfStream
	    	DiskUUID = LTrim(objExecObject.StdOut.ReadLine())
        	Exit Do
	Loop	
end function
'' SIG '' Begin signature block
'' SIG '' MIIH0QYJKoZIhvcNAQcCoIIHwjCCB74CAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' jJiETa4bz0dlDGYcJLd8JQNWaTai+/EI3gM7/90l30+g
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
'' SIG '' BgkqhkiG9w0BCQQxIgQgUsiNyDX52kXIzTfyVW4ZX6ME
'' SIG '' gBn3cdOYWgsNrWJDJvwwDQYJKoZIhvcNAQEBBQAEggEA
'' SIG '' WCkx6VTRxNpdKXOGLBpCUv9+MrW1f5YWFV07CH8WGprV
'' SIG '' G13dbfcZMLjh4q9pKOiajcEVz4zrs1Ye2aV1aJ1EkRr2
'' SIG '' pZzxsEvpvaW5DI7EbnMJGgi9ufArxRZ3yfOeTDo3AMK3
'' SIG '' y01Y+hXtzqTO3u9P9JruYobAajH5wySfrt4BwNazwoCx
'' SIG '' E5QpdDnlaykMAxjB1EuIGxpH8JmMCKJrjXPYWSr9myr2
'' SIG '' 8ogGK9GuoOzjYK3kjSBZLEebaEneGIDXz6UCeiuDobWP
'' SIG '' VDng1V04JKMzhLzwQ4ntXl0iwMextcmGPdBPZ0DWHX69
'' SIG '' mG9fKym+LPge+9H8ZAxpwFLDGGyx0Aj6wQ==
'' SIG '' End signature block
