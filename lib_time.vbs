'Время в UTC
function timeGetUtcNow
	dim colItems, item
	Set colItems = objWmi.ExecQuery("Select * from Win32_UTCTime")
	For Each item In colItems
		If Not IsNull(item) Then
			timeGetUtcNow = item.Day & "/" & item.Month & "/" & item.Year & " " & item.Hour & ":" & item.Minute & ":" & item.Second
			exit function
		End If
	Next
end function

'Timestamp (YYYY-MM-DD HH:MM:SS) в UTC
function timeGetUtcTimstamp
	dim colItems, item
	Set colItems = objWmi.ExecQuery("Select * from Win32_UTCTime")
	For Each item In colItems
		If Not IsNull(item) Then
			timeGetUtcTimstamp = item.Year & "-" & item.Month & "-" & item.Day & " " & item.Hour & ":" & item.Minute & ":" & item.Second
			exit function
		End If
	Next
end function


'время логона в формате WMI 20191026103227.687031+300
function timeLogonWmi
	dim logItems, objItem
	Set logItems = objWMIService.ExecQuery ("Select * from Win32_LogonSession")
	'почемуто вот этот вот запрос отдает мне 2 элемента. по собственно времени логона они практически идентичны
	'отличаются в микросекундах, но все же отличны. потому будем выбирать тот, что раньше
	timeLogonWmi=Null
	For Each objItem in logItems
	'	Msg "AuthenticationPackage: " & objItem.AuthenticationPackage &VBCR _
	'	& "Caption: " & objItem.Caption &VBCR _
	'	& "Description: " & objItem.Description &VBCR _
	'	& "InstallDate: " & objItem.InstallDate &VBCR _
	'	& "LogonId: " & objItem.LogonId &VBCR _
	'	& "Name: " & objItem.Name &VBCR _
	'	& "LogonType: " & objItem.LogonType &VBCR _
	'	& "StartTime: " & objItem.StartTime &VBCR _
	'	& "Status: " & objItem.Status
		if (objItem.LogonType = 2 ) then 'интерактивный вход
			if (isnull(timeLogonWmi)) then
				timeLogonWmi=objItem.startTime
			elseif (objItem.startTime < timeLogonWmi) then
				timeLogonWmi=objItem.startTime
			end if
		else
			Msg "Non-interactive logon " & objItem.startTime & " type " & objItem.LogonType 
		end if
	Next
end function


'из встроенного формата в строку вида YYYY-MM-DD HH:MM:SS
function timeVbsToTimestamp(byVal dTimestamp)
	'собственно тут у нас дата логона в UTC, осталось только сложить ее в журнал
	timeVbsToTimestamp = Year(dTimestamp) & "-" & Month(dTimestamp) & "-" & Day(dTimestamp) &_
	" " &_
	Hour(dTimestamp) & ":" & Minute(dTimestamp) & ":" & Second(dTimestamp)	
end function

' дата возвращается в формате 20191026103227.687031+300
' где цифры до точки - дата в местном часовом поясе, а после плюса (или теоретически минуса - смещение).
' т.е. распарсиваем строку на предмет даты-времени, потом вычитаем смещение в минутах и получаем дату в UTC
function timeWmiToVbs(byVal wmiTime)
	debugMsg "timeWmiToVbs: Parsing " & wmiTime
	'wscript.echo logonTime
	dim strYear,strMon,strDay,strHour,strMin,strSec,strShift
	dim plusPos,minusPos,shiftPos,sLogonDate,dLogonDate,uLogonDate
	strYear=Mid(wmiTime,1,4)
	strMon=Mid(wmiTime,5,2)
	strDay=Mid(wmiTime,7,2)
	strHour=Mid(wmiTime,9,2)
	strMin=Mid(wmiTime,11,2)
	strSec=Mid(wmiTime,13,2)

	sTimestamp = strDay & "/" & strMon & "/" & strYear & " " & strHour & ":" & strMin & ":" & strSec

	'ищем смещение
	plusPos=instr(15,wmiTime,"+")
	minusPos=instr(15,wmiTime,"-")
	shiftPos=max(plusPos,minusPos)
	'msg shiftPos
	'переводим в число и меняем знак, т.к нам надо его компенсировать и перейти в UTC
	shift=-1*CInt(mid(wmiTime,shiftPos,Len(wmiTime)-shiftPos+1))

	'wscript.echo sLogonDate & " " & shift
	'смещаем на нужное количество минут время входа
	timeWmiToVbs=dateAdd("n",shift,sTimestamp)

end function

'время входа в формате VBS
function timeLogonVbs
	timeLogonVbs=timeLogonWmi
	if (isnull(timeLogonVbs)) then
		exit function
	end if
	timeLogonVbs=timeWmiToVbs(timeLogonVbs)	
end function


function timeTzShiftHours
	timeTzShiftHours = DateDiff("h", Now(), timeGetUtcNow())
end function

Function timeGetUnixEpoch
	timeGetUnixEpoch = DateDiff("s", "01/01/1970 00:00:00", DateAdd("h",timeTzShiftHours(),Now()))
End Function

Function timeGetUnixEpochUtc
	timeGetUnixEpochUtc = DateDiff("s", "01/01/1970 00:00:00", Now())
End Function

Function timeGetUnixEpochUt
	timeGetUnixEpochUtc = DateDiff("s", "01/01/1970 00:00:00", Now())
End Function


function timeUnixToVbs (unixTime)
	timeUnixToVbs = DateAdd("s", unixTime, "01/01/1970 00:00:00")
end function

Function timeVbsToUnix (vbsTime)
	timeVbsToUnix = DateDiff("s", "01/01/1970 00:00:00", vbsTime)
End Function


'20.11.2021 20:40:21 -> 20211120204021
function timeVbsToWmi (vbsTime)
	dim tokens0,tokens1,tokens2
	'msg(vbsTime)
	tokens0 = Split(vbsTime," ")
	tokens1 = Split(tokens0(0),".")
	if (Ubound(tokens0)>0) then
		tokens2 = Split(tokens0(1),":")
	else
		tokens2 = Split("00:00:00",":")
	end if
	'msg(tokens0(0) & " " & tokens0(1))
	timeVbsToWmi=_
		stringPrependTo(tokens1(2),"0",4) &_
		stringPrependTo(tokens1(1),"0",2) &_ 
		stringPrependTo(tokens1(0),"0",2) &_
		stringPrependTo(tokens2(0),"0",2) &_ 
		stringPrependTo(tokens2(1),"0",2) &_ 
		stringPrependTo(tokens2(2),"0",2) & ".000000-000"
end function
'' SIG '' Begin signature block
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' lLRIxbCZ0P7wktw/23tCfWBa+jBnFbH3u0uvazSnzymg
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
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIBNHTB2u9E4j
'' SIG '' bQAYmPEX6YOv/efrAFlHE7zbIW97ZepMMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAIfmU8DBEbRfV0+cCp4DR1kpZcGgYclo
'' SIG '' mm2emTLWxaaBFJvRa1D61tJqiYfDhKgzpoyME8R18eiW
'' SIG '' gEMXn/gfsVpbUltg3kW7E1OiwJ6XDOxKuQQ9LDG+WhqC
'' SIG '' lNBcHPQArqVQWCnd5f1aHtdrDSnn4ZbZLGfAgfzpqVaf
'' SIG '' R8H18z5Xv1+KQYr+uNXkKtWa2qR1ks/JNtdLiBKxss8K
'' SIG '' XRT+A5TJDP+5d/tkqhY+TzAPMCX6veKkbx7thJ9o74xt
'' SIG '' JpCpynrWEg/t8X9jbMut8ljkycz+nkFk1jBFjyTlOcJ3
'' SIG '' wjZf8ay2pcFRpsVb/1wYToJEK+7IxUV7V9R8PYa9ftf2IKg=
'' SIG '' End signature block
