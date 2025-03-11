'Библиотечка с функциями работы с реестром
'без этого ну никак не обойтись
Option Explicit

'Ошибки работы с 
Dim RegErr

on error resume next
Dim objReg	: Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
on error goto 0

Dim WorkDir : WorkDir =	WshShell.ExpandEnvironmentStrings("%TEMP%") & "\"
Dim WindowsDir : WindowsDir = objFSO.GetSpecialFolder(0)


' Constants (taken from WinReg.h)
'
Const HKEY_CLASSES_ROOT   = &H80000000
Const HKEY_CURRENT_USER   = &H80000001
Const HKEY_LOCAL_MACHINE  = &H80000002
Const HKEY_USERS          = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005

Const REG_SZ        = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY    = 3
Const REG_DWORD     = 4
Const REG_MULTI_SZ  = 7


Dim SessionName: SessionName = wshShell.ExpandEnvironmentStrings( "%SESSIONNAME%" )
if ( SessionName = "%SESSIONNAME%" ) then
	Dim arrSubkeys
	Dim counter
	on error resume next
	'Ошибка: Сбой загрузки поставщика
	'Код: 80041013
	'Источник: SWbemObjectEx
	objReg.EnumKey HKEY_CURRENT_USER, "Volatile Environment", arrSubKeys
	on error goto 0
	If IsArray(arrSubKeys) then
		if Ubound(arrSubKeys)>0 Then
			counter=arrSubKeys(0)
			objReg.GetStringValue HKEY_CURRENT_USER, "Volatile Environment\" & counter, "SESSIONNAME", SessionName
			SessionName=SessionName & " "
		End If
	End If
End if






'REGISTRY ROUTINE ------------------------------------------------------
'Вытащить куст (HIVE) из строки пути
Function getRegHive(RegKey)
	dim strHive
	strHive=left(RegKey,instr(RegKey,"\"))
	if strHive="HKCR\" or strHive="HKR\" or strHive="HKEY_CLASSES_ROOT\" then getRegHive=HKEY_CLASSES_ROOT
	if strHive="HKCU\" or strHive="HKEY_CURRENT_USER\" then getRegHive=HKEY_CURRENT_USER
	if strHive="HKCC\" or strHive="HKEY_CURRENT_CONFIG\" then getRegHive=HKEY_CURRENT_CONFIG
	if strHive="HKLM\" or strHive="HKEY_LOCAL_MACHINE\" then getRegHive=HKEY_LOCAL_MACHINE
	if strHive="HKU\"  or strHive="HKEY_USERS\" then getRegHive=HKEY_USERS
End Function


Function getRegKeyPath(Regkey)
	getRegKeyPath = right(RegKey,len(RegKey)-instr(RegKey,"\"))
End Function

'Получить список подключей
Function RegEnumKeys(RegKey)
	dim hive, strKeyPath, arrSubKeys
	hive=getRegHive(RegKey)
	strKeyPath=getRegKeyPath(RegKey)

	on error resume next
	objReg.EnumKey Hive, strKeyPath, arrSubKeys
	on error goto 0

	RegEnumKeys=arrSubKeys
End Function


'читает реестр
Function regRead(ByVal Path)
	debugMsg "Reading " & Path & " ... "
	on error resume next
	RegRead = WshShell.RegRead (Path)
	if err.number<>0 then
		debugMsg "Error while Reading " & Path
		regRead = false
	end if
	on error goto 0
End Function

'пишет в реестр
sub regWrite (ByVal Path, ByVal varType, ByVal varVal)
	debugMsg "Writing " & Path & "=" & varVal & "(" & varType & ") ... "
	on error resume next
	WshShell.RegWrite Path, varVal, varType
	if err.number<>0 then
		debugMsg "Error while Writing " & Path
		regWrite = false
	end if
	on error goto 0
End Sub

'удаляет путь в реестре
sub regDelete (ByVal Path)
	debugMsg "Deleting " & Path & " ... "
	on error resume next
	WshShell.RegDelete Path
	if err.number<>0 then
		debugMsg "Error while Deleting " & Path & vbCrLf &_
			"Err code: " & Err.Number & vbCrLf &_ 
			"Description: " & Err.Description & vbCrLf &_ 
			"Source: " & Err.Source 
		regDelete = false
	end if
	on error goto 0
End Sub

'сверяет содержимое с переданными параметрами и корректирует реестр
sub regCheck (ByVal Path, ByVal varType, ByVal varVal)
	Dim tmp : tmp=regRead(Path)
	if tmp = varVal Then
		Msg tmp & " already set"
	else
		Msg "Got " &tmp& " instead of " & varVal
		if (varVal<>False) then
			regWrite Path, varType, varVal
		else
			regDelete Path
		end if
	end if
End Sub

'проверяет ветку на наличие значения
function regExists (ByVal strKey)
	dim ssig: ssig="Unable to open registry key"
	debugMsg "Searchin "&strKey
	on error resume next
    	err.clear
	dim present: present = WshShell.RegRead(strKey)
	on error goto 0
	if err.number<>0 then
		debugMsg "Got some error on "&strKey
	    	if right(strKey,1)="\" then    'strKey is a registry key
        		if instr(1,err.description,ssig,1)<>0 then
				regExists=true
        		else
            			regExists=false
        		end if
    		else    'strKey is a registry valuename
        		regExists=false
    		end if
	    	err.clear
	else
    		regExists=true
	end if
end function


Sub regCleanFolder(hive, path)
	Msg "Cleaning reg folder " & hive & "," & path & "..."
	dim  arrSubKeys, subkey
	on error resume next
    	err.clear
	objReg.EnumKey hive, path, arrSubKeys
	if err.number<>0 then
		debugMsg "Got some error on getting subkeys on " & hive & "," & path
  	If Not IsNull(arrSubKeys) Then
    		For Each subkey In arrSubKeys
			Msg "Deleting reg folder " & path & "\" & subkey & "..."
      			objReg.DeleteKey hive, path & "\" & subkey
    		Next
	else
		Msg "Empty"
  	End If
End Sub


Sub regDeleteRecursive(RegPath)
	'удаляем принципиально папки
	if not (right(RegPath,1) = "\") then
		regPath=regPath & "\"
	end if

	if (not regExists (RegPath)) then
		Msg "Folder " & RegPath & " not exist (no need to delete)"
		exit sub
	end if
	
	Msg "Deleting reg folder " & RegPath & "..."
	dim arrSubKeys, subkey
	arrSubkeys=RegEnumKeys(RegPath)
  	If Not IsNull(arrSubKeys) Then
    		For Each subkey In arrSubKeys
      			call regDeleteRecursive(RegPath & subkey & "\")
    		Next
	else
		Msg "No subfolders"
  	End If
	regDelete RegPath
End Sub


