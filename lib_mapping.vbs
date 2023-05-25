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
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' jJiETa4bz0dlDGYcJLd8JQNWaTai+/EI3gM7/90l30+g
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
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIFLIjcg1+dpF
'' SIG '' yM038lVuGV+jBIAZ93HTmFoLDa1iQyb8MA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBACZ35kAVLwUIA7HCsnAyJm0R58FCHxX9
'' SIG '' K72AnQ4fYEr6X/WIJwhUH0+cTC2/NYMRBr48oucLZ5nw
'' SIG '' +5I9Yagm6iAh4/Awzwf8PiWDunCQdtuhFI/c4yE9SKvi
'' SIG '' 5ne4aRBQ4vYnMtFT+ESE+fQR+9ixlOpDFSAdfwwlmBIZ
'' SIG '' x8nRbys6jPgN98T8IBalGbXo7BgzaK5bhFYAqqdeI/tF
'' SIG '' iHJ38UIvOuagsowSJxFMD8xWLe8vJghv5oSJK6gCpAPH
'' SIG '' WDhwP+Sh9uXxtWhxqaI6WClWSEr2Am9LETRemEj8gl9z
'' SIG '' BI5H5wAik0Kd89CrVlKBJzdjtixeCr48dS3oP3zX84Rwm2s=
'' SIG '' End signature block
