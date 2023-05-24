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
'' SIG '' MIIH0QYJKoZIhvcNAQcCoIIHwjCCB74CAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' lLRIxbCZ0P7wktw/23tCfWBa+jBnFbH3u0uvazSnzymg
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
'' SIG '' BgkqhkiG9w0BCQQxIgQgE0dMHa70TiNtABiY8Rfpg6/9
'' SIG '' 5+sAWUcTvNshb3tl6kwwDQYJKoZIhvcNAQEBBQAEggEA
'' SIG '' UKxBIrhtFe4ugQnR3qgkfJf77DT/qtu1IlDzuyOjkxno
'' SIG '' hEsRgtBTyH+KYIjhrBwycUFZ3oezlUzMmUyq4xGScoIR
'' SIG '' I9q4j/h8gEV2Y4CaMlcR0w7haPGDItABALpaZcg6/3vV
'' SIG '' QVC87YxdTclgJiqqjwgSwkLGx8N56aKbi6AOOIHJY1kl
'' SIG '' BC+mCY2yeVfV7a8ajPeIGYpuntgtiLKErGbc4Dec23w2
'' SIG '' DclPGvSGuL8ZCi7DEdJhpbhN05LLyqEX8GkywEvMVE85
'' SIG '' pUso2q2NnOAtj+e7xv/kcSgnnzJwCEObvVT+bRE3ywv8
'' SIG '' kwGfZ3P337zW08u1rXgb+Fw3CYi9ij4P4Q==
'' SIG '' End signature block
