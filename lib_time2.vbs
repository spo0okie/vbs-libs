'сдвиг в минутах часового пояса
'-300  урал
'-180  МСК
Function timeZoneOffsetMinutes()
    Dim offsetMinutes
    
    ' Получаем смещение в минутах из реестра
    timeZoneOffsetMinutes = wshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
End Function


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
	shift=-1*CLng(mid(wmiTime,shiftPos,Len(wmiTime)-shiftPos+1))

	'wscript.echo sLogonDate & " " & shift
	'смещаем на нужное количество минут время входа
	timeWmiToVbs=dateAdd("n",shift,sTimestamp)

end function

'текущее время в формате VBS но в UTC а не в текущем часовом поясе
'нужно для сра
function timeVbsNowUtc()
	timeVbsNowUtc=DateAdd("s", timeZoneOffsetMinutes()*60, Now())
end function

'конвертирует дату VBS в UNIXTIME
Function timeVbsToUnix (vbsTime)
	timeVbsToUnix = DateDiff("s", "01/01/1970 00:00:00", vbsTime)
End Function

'UNIXTIME в дату VBS
function timeUnixToVbs (unixTime)
	timeUnixToVbs = DateAdd("s", unixTime, "01/01/1970 00:00:00")
end function

'текущее время в UNIXTIME
Function timeGetUnixEpoch
	timeGetUnixEpoch = timeVbsToUnix(timeVbsNowUtc())
End Function




'20.11.2021 20:40:21 -> 20211120204021
'используется для формирования запросов в WMI с указанием времени
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

function uptimeSeconds()
	dim colOperatingSystems : colOperatingSystems = getWmiQueryArrayCrit("SELECT LastBootUpTime FROM Win32_OperatingSystem")
	dim objOS, wmiBootTime, bootTime
	For Each objOS In colOperatingSystems
		wmiBootTime = objOS.LastBootUpTime
		Exit For
	Next
	bootTime = timeWmiToVbs(wmiBootTime)
	uptimeSeconds = DateDiff("s",bootTime,timeVbsNowUtc())
end function

function uptimeString()
	Dim days, hours, minutes, seconds, uptime
	uptime=uptimeSeconds()
	' Разбиваем на дни, часы, минуты, секунды
	days = Int(uptime / 86400)
	uptime = uptime Mod 86400
	hours = Int(uptime / 3600)
	uptime = uptime Mod 3600
	minutes = Int(uptime / 60)
	seconds = uptime Mod 60
	uptimeString = days & " days " & hours & ":" & minutes & ":" & seconds
end function