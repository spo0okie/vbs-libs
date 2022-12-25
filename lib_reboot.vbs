Option Explicit

function isRebootPending() 

	isRebootPending=true

	dim testValue
	testValue=regRead("HKLM\SOFTWARE\Microsoft\Updates\UpdateExeVolatile")
	if ((not testValue = false) and (not testValue = 0)) then exit function


	dim checkExistance, item
	checkExistance=array(_
		"HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations",_
		"HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations2",_
		"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired\",_
		"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting\",_
		"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\DVDRebootSignal",_
		"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",_
		"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress",_
		"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending",_
		"HKLM\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttempts",_
		"HKLM\SYSTEM\CurrentControlSet\Services\Netlogon\JoinDomain",_
		"HKLM\SYSTEM\CurrentControlSet\Services\Netlogon\AvoidSpnSet"_
	)
	for each item in checkExistance
		if (regExists(item)) then exit function
	next
	
	dim arrSubKeys, subkey
	arrSubkeys=RegEnumKeys("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\Pending")
  	If Not IsNull(arrSubKeys) Then
   		For Each subkey In arrSubKeys
			exit function
   		Next
  	End If

	isRebootPending=false
end function

