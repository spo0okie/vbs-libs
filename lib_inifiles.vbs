option Explicit
'-----------------------------------------------------------------------
'INI FILES LIBRARY           -------------------------------------------
'-----------------------------------------------------------------------
'Мутная библиотека по правке разных текстовых файлов (INIподобных, но не ХМЛ)
'Переменные принятые для обозначения формата файла
'eq		: сивол равенства  					= : => := ...
'secL, secR	: Ограничители объявления секции 	[, ]
'ComDelims	: с чего начинаются каменты 		# rem :: ' // ...
'Cst, Cend	: сиволы начала и конца блока каментов /* */ <!-- --!> НЕ ОБРАБАТЫВАЮТСЯ (оставлено на будущее)
'CaSense	: чувствительность переменных и секций к регистру true\false
'defCrLf	: Перенос строки по умолчанию (используем его, если не поймем какой в файле уже используется)
'Section	: C какой секцией хотим работать
'KeyName	: С какой переменной
'Value		: Значение (значение по умолчанию при чтении, установить при записи)
'req		: read\write

'реальные определения типов файлов для библиотеки находятся в конце файла


'возвращает начало камента в строке (ComDelims - массив символов начала камента)
'или конец строки+1 если нет камента
function CommentAt(byVal INIString, byVal ComDelims, byVal Cst, byVal Cend, byVal CaSense)
	Dim startPos, testPos, Delim
	if not CaSense then
		INIString=UCase(INIString)
	end if
	startPos=Len(INIString)+1	'по умолчанию камента нет и он за концом строки
	for each Delim in ComDelims
		if not CaSense then
			Delim=UCase(Delim)
		end if
		testPos = InStr(1, INIString, Delim, vbTextCompare)
		if testPos>0 then
			startPos=min(startPos,testPos)
		end if
	next
	CommentAt = startPos
end function

'возвращает полезную часть строки, без каментов и пустот по краям
function GetUsefulPart(INIString, ComDelims, Cst, Cend, CaSense)
	GetUsefulPart=TrimWithTabs(Left(INIString,CommentAt(INIString, ComDelims, Cst, Cend, CaSense)-1))
end function

'если строка является объявлением секции, то возвращает ее имя, иначе ""
function GetSectionName(INIString,ldelim,rdelim)
	If len(ldelim)>0 and len(rdelim)>0 and Left(INIString, len(ldelim)) = ldelim and Right(INIString, len(rdelim)) = rdelim then
		GetSectionName=TrimWithTabs(Mid(INIString,len(ldelim)+1,Len(INIString)-len(ldelim)-len(rdelim)))
	else
		GetSectionName=""
	end if
end function

'возвращает строку-объявение секции, если она отличается от текущей
function declareSectionName(CurSection,Section,ldelim,rdelim)
	if Len(Section)>0 and Section<>CurSection and Len(ldelim)>0 and Len(rdelim)>0 then
		declareSectionName=ldelim&Section&rdelim	'объявление секции
	else
		declareSectionName=""	'ничего не надо объявлять
	end if
end function

'возвращает строку-объявение переменной
function declareVariable(varName,eq,varValue)
	declareVariable=varName&eq&varValue
end function

'ищет в строке имя переменной
function GetVariableName(INIString,eq)
	dim eqPos
	eqPos = InStr(1, INIString, eq, vbTextCompare)
	If eqPos>0 Then
		GetVariableName=TrimWithTabs(Left(INIString,eqPos-len(eq)))
	else
		'wscript.echo eq & " not found in " & INIString
		GetVariableName=""
	end if
end function

'возвращает конец строки, если строка не пустая
function optionalCrLf(somedata,CrLf)
	if Len(somedata)>0 then
		optionalCrLf=CrLf
	else
		optionalCrLf=""
	end if
end function

'возвращает строку-комментарий
function commentLine(Comment,ComDelims,CrLf)
	commentLine=""
	'wscript.echo "comline: " & UBound(ComDelims) & " " & Len(Comment)
	if IsArrayDimmed(ComDelims) and Len(Comment)>0 then
		commentLine = ComDelims(0) & Comment & CrLf
	end if
end function

'ищет в строке значение переменной
function GetVariableVal(INIString,eq)
	dim eqPos
	eqPos = InStr(1, INIString, eq, vbTextCompare)
	If eqPos>0 Then
		GetVariableVal=TrimWithTabs(Right(INIString,Len(INIString)-eqPos-len(eq)+1))
	else
		GetVariableVal=""
	end if
end function

'ищет в файле набор символов обозначающих перенос строки
function GetFileCrLf(INIContents,default)
	'если не передали какой перенос по умолчанию, то вот этот пусть будет
	if Len(default)=0 then
		default=vbCrLf
	end if

	'смотрим какой перенос у нас где встречается в файле
	dim crPos,lfPos
	crPos = InStr(1, INIContents, vbCr, vbTextCompare)
	lfPos = InStr(1, INIContents, vbLf, vbTextCompare)
	If crPos>0 or lfPos>0 Then
		if (crPos>0) and ((crPos<lfPos) or (lfPos=0)) then 'подозрение на CrLf
			if lfPos=crPos+1 then
				'wscript.echo "GetFileCrLf: dected Windows CRLF at " & crPos
				GetFileCrLf=vbCr&vbLf 	'Windows
			else
				'wscript.echo "GetFileCrLf: dected UNIX CR at " & crPos
				GetFileCrLf=vbCr		'UNIX
			end if
		elseif (lfPos>0) and ((lfPos<crPos) or (crPos=0)) then 'подозрение на LfCr
			if lfPos=crPos-1 then
				'wscript.echo "GetFileCrLf: dected LFCR WTF!!!??? " & lfPos
				GetFileCrLf=vbLf&vbCr'файл может и не текстовый вовсе???
			else
				'wscript.echo "GetFileCrLf: dected Mac LF at " & lfPos
				GetFileCrLf=vbLf		'Mac
			end if
		else '???WTF???
			GetFileCrLf=default 'по умолчанию
		end if
	else ' не нашли ни cr ни lf
		wscript.echo "GetFileCrLf: Newline symbol not detected, using default: "&default
		GetFileCrLf=default 'по умолчанию
	end if
end function


'далее идет пара громоздких, не очень удобных, функций которые используя все что написано выше работают непосредственно с файлами


'универсальный парсер конфиг файлов (INIподобных, но не ХМЛ) - читает или пишет секция->ключ->значение
'Читает файл со следующим форматированием:
'eq			: сивол равенства  					= : => := ...
'secL, secR	: Ограничители объявления секции 	[, ]
'ComDelims	: с чего начинаются каменты 		# rem :: ' // ...
'Cst, Cend	: сиволы начала и конца блока каментов /* */ <!-- --!> НЕ ОБРАБАТЫВАЮТСЯ (оставлено на будущее)
'CaSense	: чувствительность переменныхи секций к регистру true\false
'defCrLf	: Перенос строки по умолчанию (используем такой, если не поймем какой в файле уже используется)
'Section	: C какой секцией хотим работать
'KeyName	: С какой переменной
'Value		: Значение (значение по умолчанию при чтении, установить при записи)
'req		: read\write
Function parseINIString(FileName, eq, secL, secR, ComDelims, Cst, Cend, CaSense, defCrLf, Section, KeyName, Value, req, Comment)
	Dim i, INIContents, INIStrings, INIString, UsefulPart, CurSection, CrLf
	Dim testSec, testVar, testVal, jobSec, jobVar, jobDone

	'Get contents of the INI file As a string
	DebugMsg "reading file : "&FileName
	INIContents = GetFile(FileName)
	CrLf=GetFileCrLf(INIContents,defCrLf)
	INIStrings = Split (INIContents, CrLf)
	DebugMsg "found lines : "&ubound(INIStrings)&CrLf

	CurSection=""		'на начало файла секции нет (может и не будет до конца файла)
	jobDone=false		'работа не сделана
	jobSec = Section	'ищем такую секцию
	jobVar = KeyName	'и такую переменную
	if not CaSense then	'если синтаксис не чувствителен к регистру, то приводим все символы к одному регистру
		jobSec=UCase(jobSec)
		jobVar=UCase(jobVar)
	end if

	for i = 0 to UBound(INIStrings)
		INIString=INIStrings(i)
		UsefulPart=GetUsefulPart(INIString, ComDelims, Cst, Cend, CaSense) 'выделяем полезную часть строки
		testSec=GetSectionName(UsefulPart,secL,secR)	'смотрим есть ли в ней новая секция
		testVar=GetVariableName(UsefulPart,eq)	'или переменная
		if not CaSense then	'если синтаксис не чувствителен к регистру, то приводим все символы к одному регистру
			testSec=UCase(testSec)
			testVar=UCase(testVar)
		end if
		if Len(testSec)>0 then
			DebugMsg "Section detected: " & testSec & " vs " & jobSec
			if CurSection=jobSec then	'кончилась секция в которой мы искали переменную
				if req="write" and not jobDone then 'в той секции надо было записать переменную, а ее не было
					'допишем в конец предыдущей секции
					DebugMsg "INSERTING BEFORE SECTION :" &i
					INIStrings(i)=	commentLine(Comment,ComDelims,CrLf)& _
									declareVariable(KeyName,eq,Value)&_
									CrLf & INIStrings(i)
					jobDone=true
				end if
			end if
			CurSection=testSec
		elseif len(testVar)>0 then
			testVal=GetVariableVal(UsefulPart,eq)
			DebugMsg "Variable detected: " & testVar & " => " & testVal
			if CurSection=jobSec and testVar=jobVar then
			'мы нашли свою переменную
				if req="write" then 'запись
					'рисуем камент об изменении файла 'комментируем текущую линию 'пишем свою
					if testVal<>Value then
						DebugMsg "CHANGING CURRENT :" &i
						INIStrings(i)=	commentLine(INIStrings(i),ComDelims,CrLf)& _
										commentLine(Comment,ComDelims,CrLf)& _
										declareVariable(KeyName,eq,Value)
					end if
					jobDone=true
				elseif req="delete" then 'удаление значения 'комментируем текущую линию 
						DebugMsg "CHANGING CURRENT :" &i
						INIStrings(i)=	commentLine(Comment,ComDelims,CrLf)& _
										commentLine(INIStrings(i),ComDelims,"")
				
					jobDone=true
				else 'чтение
					parseINIString=testVal
					jobDone=true
				end if
			end if
		end if
	next
	if not jobDone then	'переменную не нашли
		if req="write" then 'запись
			'добавляем еще одну сроку
			redim preserve INIStrings(UBound(INIStrings)+1)
			if (UBound(INIStrings) > i+1) then
				dim j:	for j=UBound(INIStrings) to i+1
					DebugMsg(j)
					INIStrings(j)=INIStrings(j-1)
				next
			end if
			'i=i-1
			DebugMsg "Inserting in string " &i
			INIStrings(i)=	declareSectionName(CurSection, Section, secL, secR)& _
							optionalCrLf(declareSectionName(CurSection, Section, secL, secR),CrLf)&_
							commentLine(Comment,ComDelims,CrLf)& _
							declareVariable(KeyName,eq,Value) '&CrLf
			DebugMsg INIStrings(i)
			jobDone=true
		else 'чтение
			parseINIString=Value
		end if
	end if
	if req="write" or req="delete" then 'запись
		if jobDone then WriteFile FileName, Join(INIStrings,CrLf)
		parseINIString=jobDone
	end if

End Function


'универсальный парсер текстовых файлов, ищет строки, вставляет нужную до или после ключевой
'Читает файл со следующим форматированием:
'eq			: сивол равенства  					= : => := ...
'secL, secR	: Ограничители объявления секции 	[, ]
'ComDelims	: с чего начинаются каменты 		# rem :: ' // ...
'Cst, Cend	: сиволы начала и конца блока каментов /* */ <!-- --!> НЕ ОБРАБАТЫВАЮТСЯ (оставлено на будущее)
'CaSense	: чувствительность переменныхи секций к регистру true\false
'defCrLf	: Перенос строки по умолчанию (используем такой, если не поймем какой в файле уже используется)
'Section	: C какой секцией хотим работать
'Keystring	: Какая строка опорная
'Addition	: Что добавить
'Position	: before\after
'Many		: Единожды или все вхождения
Function parseTXTString(FileName, eq, secL, secR, ComDelims, Cst, Cend, CaSense, defCrLf, Section, Keystring, Addition, Position, Many)
	Dim i, INIContents, INIStrings, INIString, UsefulPart, CurSection, CrLf
	Dim testSec, jobSec, jobDone

	'Get contents of the INI file As a string
	'wscript.echo "reading file : "&FileName
	INIContents = GetFile(FileName)
	CrLf=GetFileCrLf(INIContents,defCrLf)
	INIStrings = Split (INIContents, CrLf)
	'wscript.echo "parseTXTString: found " & ubound(INIStrings) & " lines with delim of " & Len(CrLf) &" bytes"

	CurSection=""		'на начало файла секции нет (может и не будет до конца файла)
	jobDone=false		'работа не сделана
	jobSec = Section	'ищем такую секцию
	if not CaSense then	'если синтаксис не чувствителен к регистру, то приводим все символы к одному регистру
		jobSec=UCase(jobSec)
		'msg ("parseTXTString: case insensetive mode")
	end if
	for i = 0 to UBound(INIStrings)
		INIString=INIStrings(i)
		'msg("at first it is "&INIString	)
		testSec=GetSectionName(GetUsefulPart(INIString, ComDelims, Cst, Cend, CaSense),secL,secR)	'смотрим есть ли новая секция
		if not CaSense then	'если синтаксис не чувствителен к регистру, то приводим все символы к одному регистру
			testSec=UCase(testSec)
		end if
		if Len(testSec)>0 then
			CurSection=testSec
		elseif INIString=Keystring then
			'msg ("parseTXTString: found position for insert!")
			if Ucase(Position)="BEFORE" then
				INIStrings(i) =	Addition & CrLf & INIStrings(i)
				msg ("parseTXTString: found position for insert!")
			else
				INIStrings(i) =	INIStrings(i) & CrLf & Addition
			end if
			jobDone=true
			if not Many then
				exit for
			end if
		else
			'msg (INIString & " != " & Keystring)
		end if
	next
	if jobDone then WriteFile FileName, Join(INIStrings,CrLf)
	parseTXTString=jobDone
End Function


'далее идут уже функции, которые облегчают работу с этими 2мя что выше


'проверяет что переданная переменная - действительно структура описывающая тип файла
Function CheckFileTypeDescr(ByVal FType)
	'eq			: сивол равенства  					= : => := ...
	'secL, secR	: Ограничители объявления секции 	[, ]
	'ComDelims	: с чего начинаются каменты 		# rem :: ' // ...
	'Cst, Cend	: сиволы начала и конца блока каментов /* */ <!-- --!> НЕ ОБРАБАТЫВАЮТСЯ (оставлено на будущее)
	'CaSense	: чувствительность переменныхи секций к регистру true\false
	'defCrLf	: Перенос строки по умолчанию (используем такой, если не поймем какой в файле уже используется)
	'Msg typeName (FType)
	CheckFileTypeDescr=true
	dim flds,fld
	flds=array("eq","secL","secR","ComDelims","Cst","Cend","CaSense","defCrLf")
	for each fld in flds
		if not FType.exists(fld) then
			msg "CheckFileTypeDescr err: " & fld & " not set"
			CheckFileTypeDescr=false
		end if
	next
End Function


'вот собственно ради этих трех и писалась вся библиотечка

'получить значение из INI файла
Function conffile_get(ByVal FPath, ByVal FType, ByVal Section, ByVal Key, ByVal Default)
	if CheckFileTypeDescr(FType) then
		conffile_get=parseINIString(FPath, FType("eq"), FType("secL"), FType("secR"), FType("ComDelims"), FType("Cst"), FType("Cend"), FType("CaSense"), FType("defCrLf"), Section, Key, Default, "read", "")
	else
		conffile_get=Default
	end if
end Function


'записать значение в INI файл
Function conffile_set(ByVal FPath, ByVal FType, ByVal Section, ByVal Key, ByVal Value, ByVal Comment)
	dim mode
	if Value=unset_me then
		mode="delete"
	else
		mode="write"
	end if
	if CheckFileTypeDescr(FType) then
		conffile_set=parseINIString(FPath, FType("eq"), FType("secL"), FType("secR"), FType("ComDelims"), FType("Cst"), FType("Cend"), FType("CaSense"), FType("defCrLf"), Section, Key, Value, mode, Comment)
	else
		conffile_set=false
	end if
end Function


'исправить значение в INI файл
'возвращает признак произведенных в файле изменений
Function conffile_fix(ByVal FPath, ByVal FType, ByVal Section, ByVal Key, ByVal Value, ByVal Comment)
	if CheckFileTypeDescr(FType) then
		if (Value=conffile_get(FPath,FType,Section,Key,"value not found marker")) then
			conffile_fix=false
		else
			conffile_set FPath,FType,Section,Key,Value,Comment
			conffile_fix=true
		end if
	else
		conffile_fix=false
	end if
end Function




'вставляет переменную после нужной строки
Function textfile_set_after(ByVal FPath, ByVal FType, ByVal Section, ByVal Key, ByVal Value, ByVal Comment, ByVal after)
	if CheckFileTypeDescr(FType) then
		if not conffile_get(FPath, FType, Section, Key, "un1Que_deFFault") = "un1Que_deFFault" then 'если уже есть какое-то значение - меняем
			textfile_set_after=parseINIString(FPath, FType("eq"), FType("secL"), FType("secR"), FType("ComDelims"), FType("Cst"), FType("Cend"), FType("CaSense"), FType("defCrLf"), Section, Key, Value, "write", Comment)
		else
			dim CrLf, Addition
			CrLf=GetFileCrLf(GetFile(FPath),FType("defCrLf"))
			Addition= commentLine(Comment,FType("ComDelims"),CrLf) & declareVariable(Key,FType("eq"),Value)
			'msg("trying to add " & Addition & " after " & after)
			textfile_set_after=parseTXTString(FPath, FType("eq"), FType("secL"), FType("secR"), FType("ComDelims"), FType("Cst"), FType("Cend"), FType("CaSense"), FType("defCrLf"), Section, after, Addition, "after", true)
		end if
	else
		conffile_set=false
	end if
end Function




'определения типов файлов для библиотеки
'файл .bat/.cmd
dim ftype_bat : set ftype_bat = CreateObject("Scripting.Dictionary")
	ftype_bat.add "eq",		"="
	ftype_bat.add "secL",	""
	ftype_bat.add "secR",	""
	ftype_bat.add "ComDelims",array("rem ","::")
	ftype_bat.add "Cst",	""
	ftype_bat.add "Cend",	""
	ftype_bat.add "CaSense",false
	ftype_bat.add "defCrLf",vbCrLf

'файл .properties от ТС
dim ftype_tc_conf : set ftype_tc_conf = CreateObject("Scripting.Dictionary")
	ftype_tc_conf.add "eq",		"="
	ftype_tc_conf.add "secL",	""
	ftype_tc_conf.add "secR",	""
	ftype_tc_conf.add "ComDelims",array("#")
	ftype_tc_conf.add "Cst",	""
	ftype_tc_conf.add "Cend",	""
	ftype_tc_conf.add "CaSense",true
	ftype_tc_conf.add "defCrLf",vbCrLf

'файл etc\hosts
dim ftype_etc_hosts : set ftype_etc_hosts = CreateObject("Scripting.Dictionary")
	ftype_etc_hosts.add "eq",		" "
	ftype_etc_hosts.add "secL",	""
	ftype_etc_hosts.add "secR",	""
	ftype_etc_hosts.add "ComDelims",array("#")
	ftype_etc_hosts.add "Cst",	""
	ftype_etc_hosts.add "Cend",	""
	ftype_etc_hosts.add "CaSense",false
	ftype_etc_hosts.add "defCrLf",vbCrLf

'файл .ini
dim ftype_tc_ini : set ftype_tc_ini = CreateObject("Scripting.Dictionary")
	ftype_tc_ini.add "eq",		"="
	ftype_tc_ini.add "secL",	"["
	ftype_tc_ini.add "secR",	"]"
	ftype_tc_ini.add "ComDelims",array(";")
	ftype_tc_ini.add "Cst",	""
	ftype_tc_ini.add "Cend",	""
	ftype_tc_ini.add "CaSense",true
	ftype_tc_ini.add "defCrLf",vbCrLf

'файл .ini
dim ftype_ini : set ftype_ini = CreateObject("Scripting.Dictionary")
	ftype_ini.add "eq",		"="
	ftype_ini.add "secL",	"["
	ftype_ini.add "secR",	"]"
	ftype_ini.add "ComDelims",array("#")
	ftype_ini.add "Cst",	""
	ftype_ini.add "Cend",	""
	ftype_ini.add "CaSense",true
	ftype_ini.add "defCrLf",vbCrLf

'' SIG '' Begin signature block
'' SIG '' MIIH0QYJKoZIhvcNAQcCoIIHwjCCB74CAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' RUoK2rWWTbwKmcV+4WXW7mD7H1j6qN3tv/r3iQoAa6eg
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
'' SIG '' BgkqhkiG9w0BCQQxIgQgD6Px3+NW3zJ7r3JSKfRPh9tV
'' SIG '' hh0yL8mcj5oyrfuojyUwDQYJKoZIhvcNAQEBBQAEggEA
'' SIG '' Q7tMChXKneWPTDY6dAC2Hp3RmRH5Rm0HX8z5DTwIzJMc
'' SIG '' RYJ7PzLVD4q6nimcmihGkJU0VGMBIGpSdx6rw9vWRL0O
'' SIG '' ugVN59z9JRmwj/bkNlwePzSTvQ6+nGvRm7m3xMo5KWwT
'' SIG '' Xpyjrpyl2fOhmGno4wd51k6ZNE6IjMptNKq/dPBQth6P
'' SIG '' CABvwJaPe5HOyM+42WWyfqAmB/d8MSW37FMGLy9G6wsK
'' SIG '' NmbEgi1ck6SAPQ9Sh/ZOhoL6opH0Iuo8oYPU01UglTTe
'' SIG '' epLTGO3O/Q+EN5D1SQ9dQ+Gn13wbu1B7LDb6AmqyKsKN
'' SIG '' huAY/YGeCuqNQySZ+ZLPTttGW6FuS1UBKA==
'' SIG '' End signature block
