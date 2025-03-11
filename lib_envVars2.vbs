'Библиотечка работы с переменными окружения
'без этого ну никак не обойтись
Option Explicit
const unset_me=		"#UNSET_me#" 'это значение ставить в переменные окружения которые надо убрать

'Dim objUserEnv	   : Set objUserEnv   = wshShell.Environment("USER")
'Dim objSystemEnv  : Set objSystemEnv = wshShell.Environment("SYSTEM")
'Dim objProcessEnv : Set objSystemEnv = wshShell.Environment("PROCESS")
'Dim objVolatileEnv: Set objSystemEnv = wshShell.Environment("VOLATILE")

'-----------------------------------------------------------------------------------
'ENVIRONMENT VARIABLES


function EnvironmentVariableName (ByVal setString)
	dim eqPos
	eqPos=instr(1,setString,"=",vbTextCompare)
	if (eqPos = 0) then
		EnvironmentVariableName=setString
	else
		EnvironmentVariableName=Left(setString,eqPos-1)
	end if
end function

'проверяет установлена ли переменная в нужном окружении и устанавливает если нужно (или удаляет) 
'System		– системные переменные_среды, 
'User		– переменные_среды пользователя
'Volatile	– временные_переменные (туда пишутся всякая архитектура процессора и прочая шняга)
'Process	– переменные_среды текущего процесса
function EnvironmentVariableCorrect (ByVal Environment, ByVal varName, ByVal varVal)
	Dim objEnvironment,index
	Set objEnvironment = wshShell.Environment(Environment)
	EnvironmentVariableCorrect=false
	if (varVal<>unset_me) then
		Msg "Checking if " & Environment & " variable " & varName & " is set to " & varVal
		if (objEnvironment(varName) = varVal) then
			Msg " - yes"
		else
			objEnvironment(varName) = varVal
			EnvironmentVariableCorrect=true
			Msg " - No. Fixed"
		end if
	else
		Msg "Checking if " & Environment & " variable " & varName & " is unset "
		varName=UCase(varName)
		For Each index In objEnvironment
			'DebugMsg UCase(EnvironmentVariableName(index)) &" vs "& varName
			if UCase(EnvironmentVariableName(index)) = varName then
				objEnvironment.Remove(varName)
				EnvironmentVariableCorrect=true
				Msg " - No. Fixing"
				exit For
			end if
		Next 
		if (EnvironmentVariableCorrect = false) then
			Msg " - Yes"
		end if
	end if
end Function

sub EnvironmentVariableSet (ByVal Environment, ByVal varName, ByVal varVal)
	Dim objEnvironment
	Set objEnvironment = wshShell.Environment(Environment)
	objEnvironment(varName)=varVal
	unset(objEnvironment)
End Sub

function EnvironmentVariableGet (ByVal Environment, ByVal varName)
	Dim objEnvironment
	Set objEnvironment = wshShell.Environment(Environment)
	EnvironmentVariableGe = objEnvironment(varName)
	unset(objEnvironment)
end function

'удаляет определение переменной в каком-либо окружении
function EnvironmentVariableUnset (ByVal Environment, ByVal varName)
	Dim objEnvironment,index
	Set objEnvironment = wshShell.Environment(Environment)
	EnvironmentVariableUnset=false
	varName=Ucase(varName)
	For Each index In objEnvironment
		if UCase(EnvironmentVariableName(index)) = varName then
			objEnvironment.Remove(varName)
			EnvironmentVariableUnset=true
			exit For
		end if
	Next 
	unset(objEnvironment)
End function


sub EnvVarCorrectNow (ByVal varName, ByVal varVal)
	call EnvVarCorrect(varName, varVal)
end sub

sub EnvVarCorrect (ByVal varName, ByVal varVal)
	if (EnvironmentVariableCorrect ("SYSTEM",varName, varVal)) then
		call EnvironmentVariableCorrect ("PROCESS",varName, varVal)
		call EnvironmentVariableCorrect ("VOLATILE",varName, varVal)
	end if
end sub

sub EnvUsrVarCorrect (ByVal varName, ByVal varVal)
	if (EnvironmentVariableCorrect ("USER",varName, varVal)) then
		call EnvironmentVariableCorrect ("PROCESS",varName, varVal)
		call EnvironmentVariableCorrect ("VOLATILE",varName, varVal)
	end if
end sub


function EnvVarCheck(ByVal varName, ByVal varVal)
	Msg "Checking environment variable " & varName & " ... "
	on error resume next
	dim current
	current = WshShell.ExpandEnvironmentStrings(varName)
	if err.number <> 0 then
		Msg "Error expanding variable " & varName
		if (varVal=unset_me) then
			EnvVarCheck=true
		else
			EnvVarCheck=false
		end if
	else
		if (LCase(current)<>LCase(varVal)) then
			EnvVarCheck=false
			Msg(varName & " set to """ & current & """ instead of """ & varVal & """")
		else
			EnvVarCheck=true
		End if
	end if
end function


function EnvPathCorrect(ByVal testPath)
'проверяет наличие переданного пути в переменной PATH, добавляет если нет
	dim dirs,found,i

	Msg "Checking path variable for " & testPath & " presence ... "
	EnvPathCorrect=false

	testPath=unquotePath(trim(testPath))
	dirs=split(EnvironmentVariableGet ("SYSTEM", "PATH"),";")
	found=false

	for i=0 to ubound(dirs)
		if UCase(trim(dirs(i)))=UCase(testPath) then
			found=true
		end if
	next

	if found then
		msg " - found"
	else
		msg " - not found. Adding"
		ReDim Preserve dirs(UBound(dirs) + 1)
		dirs(UBound(dirs)) = testPath
	end if

	if (not found) then
		msg " - saving changes..."
		call EnvironmentVariableSet ("SYSTEM","PATH", join(dirs,";"))
		call EnvironmentVariableSet ("PROCESS","PATH", join(dirs,";"))
		call EnvironmentVariableSet ("VOLATILE","PATH", join(dirs,";"))
		EnvPathCorrect=true
		msg " - done"
	end if

end function
