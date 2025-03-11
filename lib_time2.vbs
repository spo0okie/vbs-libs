'Время в UTC
function timeGetUtcNow
	dim colItems, item
	Set colItems = getWmiQueryCrit("Select * from Win32_UTCTime")
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
	Set colItems = getWmiQueryCrit("Select * from Win32_UTCTime")
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
	Set logItems = getWmiQueryCrit ("Select * from Win32_LogonSession")
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
	shift=-1*CLng(mid(wmiTime,shiftPos,Len(wmiTime)-shiftPos+1))

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
