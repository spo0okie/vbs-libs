function timeGetUtcNow
	dim colItems, item
	Set colItems = objWmi.ExecQuery("Select * from Win32_UTCTime")
	For Each item In colItems
		If Not IsNull(item) Then
			timeGetUtcNow = item.Month & "/" & item.Day & "/" & item.Year & " " & item.Hour & ":" & item.Minute & ":" & item.Second
			exit function
		End If
	Next
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

function timeUnixToVbs (unixTime)
	timeUnixToVbs = DateAdd("s", unixTime, "01/01/1970 00:00:00")
end function

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