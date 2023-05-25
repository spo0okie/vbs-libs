'хз какие тут версии. Сразу не завел ща уже какой смысл
'2023-04-20
'	dirMasterSlaveSync + игнор директорий через файл .sync-skip-local-files
'			   ! фикс с копированием пустых директорий
'FILE ROUTINE ----------------------------------------------------------
'option explicit

'Копирование файла
Function xCopyFile(ByVal fromf, ByVal tof)
	'/C	Игнорирует ошибки.
	'/f	Отображает имена исходных и целевых файлов при копировании.
	'/i	Если Source является каталогом или содержит подстановочные знаки, а назначение не существует, команда xcopy предполагает, 
	'	что в поле назначение указано имя каталога и создается новый каталог. Затем команда xcopy копирует все указанные файлы в новый каталог. 
	'	По умолчанию команда xcopy предложит указать, является ли назначение файлом или каталогом.
	'/r	Копирует файлы, которые доступны только для чтения.
	'/h	Копирует файлы с атрибутами скрытых и системных файлов. По умолчанию команда xcopy не копирует скрытые или системные файлы.
	'/y	Подавляет запрос на подтверждение перезаписи существующего целевого файла.
	'/z	Выполняет копирование по сети в перезапускаемом режиме.
	dim command: command="%windir%\system32\XCOPY.exe /Y /C /F /R /H /I /Z """ & fromf & """ """ & tof & """"
	debugMsg_ "Running " & command
	xCopyFile=wshShell.run(command,0,true)
	debugMsg_n " - complete"
End Function

'Копирование файла
Sub safeCopy(ByVal fromf, ByVal tof)
	msg_ "Copying " & fromf & " to " & tof 
	dim ret : ret=xCopyFile(fromf,tof)
	msg_n " - return code: " & ret
	'safeCopy=ret
End sub

'Удаление файла, если он есть
Sub safeDelete(ByVal FName)
	msg_ "Deleting " & Fname & " ..."
	if objFSO.fileExists(Fname) then
		objFSO.deleteFile(Fname)
		msg_n " complete"
	elseif  objFSO.folderExists(Fname) then
		objFSO.deleteFolder Fname, true
		msg_n " complete"
	else
		msg_n " not exists"
	end if 
End Sub


'возвращает путь к файлу без имени файла
Function GetFilePath(ByVal FileName)
	GetFilePath=Left(FileName,InStrRev(FileName,"\"))
End Function

'возвращает некорень пути, который можно пытаться создать через мкдир
Function GetPathDirs(ByVal FileName)
	If InStr(FileName, ":\") = 2 then
		GetPathDirs=Mid(FileName,4) 'c:\
	elseif Left (FileName,2) = "\\" Then	'\\servername\...
		dim nextSlash : nextSlash = InStr(3,FileName,"\")
		nextSlash = InStr(nextSlash+1,FileName,"\")
		if (nextSlash=0) then	'\\servername
			GetPathDirs=""
		else 					'\\servername\...
			GetPathDirs=Mid(FileName,nextSlash)
		end if
	else						'????
		GetPathDirs = FileName
	End If
End Function

'возвращает корень пути, который не надо пытаться создать через мкдир
Function GetPathRoot(ByVal FileName)
	If InStr(FileName, ":\") = 2 then	'c:\
		GetPathRoot=left(FileName,3)
	elseif Left (FileName,2)="\\" Then	'\\servername\...
		dim nextSlash : nextSlash = InStr(3,FileName,"\")
		nextSlash = InStr(nextSlash+1,FileName,"\")
		if (nextSlash=0) then	'\\servername
			GetPathRoot=FileName
		else 					'\\servername\...
			GetPathRoot=Left(FileName,nextSlash)
		end if
	else						'????
		GetPathRoot = ""
	End If
End Function

'возвращает некорень пути, который можно пытаться создать через мкдир
Function GetPathDirs(ByVal FileName)
	If InStr(FileName, ":\") = 2 then
		GetPathDirs=Mid(FileName,4) 'c:\
	elseif Left (FileName,2) = "\\" Then	'\\servername\...
		dim nextSlash : nextSlash = InStr(3,FileName,"\")
		nextSlash = InStr(nextSlash+1,FileName,"\")
		if (nextSlash=0) then	'\\servername
			GetPathDirs=""
		else 					'\\servername\...
			GetPathDirs=Mid(FileName,nextSlash)
		end if
	else						'????
		GetPathDirs = FileName
	End If
End Function

'проверка наличия директории и создание в случае отсутствия
Sub CheckDir(ByVal ckPath)
	Msg "Checking dir " & ckPath & " ... "
	If (objFSO.FolderExists(ckPath) = False) Then
		dim tmp, dirs, i
		dirs = split(GetPathDirs(ckPath),"\")
		tmp = GetPathRoot(ckPath)
		'msg ("Got " & tmp)
		for i = 0 to Ubound(dirs)
			if (len(dirs(i))>0) then
				tmp = tmp & dirs(i) & "\"
			end if
			If not objFSO.FolderExists(tmp) Then
				Msg "Creating " & tmp & " ... "
				objFSO.CreateFolder(tmp)
			end if
		next
		'Msg " - Done "
	End If
End Sub


'Ищет позицию иголки в файле
function FindInFile(ByVal FileName, ByVal Needle)
	FindInFile = InStr(1, GetFile(FileName) , Needle, vbTextCompare)
end function

'То же что и сверху, только ищет конец строки иголки
'По сути за иголкой ищем CR или LF и ищем где последовательность из
'CR LF заканчивается
function FindStrEndInFile(ByVal FileName, ByVal Needle)
	msg "Searching " & needle
	dim contents : contents = GetFile(FileName)
	dim strpos : strpos = InStr(1, contents , Needle, vbTextCompare)
	dim crpos, lfpos
	if strpos>0 then
		msg "Found at " & strpos
		strpos=strpos+len(needle)-1
		'ищем какой конец строки будет первым
		crpos=InStr(strpos, contents , vbCr, vbTextCompare)
		lfpos=InStr(strpos, contents , vbLf, vbTextCompare)
		'проверяем, вдруг какого-то из них не было
		if (crpos=0) then
			crpos=len(contents)
		end if
		if (lfpos=0) then
			lfpos=len(contents)
		end if
		strpos=min(crpos,lfpos)
		msg "String ends at " & strpos
		do while strpos<len(contents) and ((mid(contents, strpos, 1)=vbCr) or (mid(contents, strpos, 1)=vbLf))
			if (mid(contents, strpos, 1)=vbCr) then
				msg "Found additional CR at " & strpos
				strpos = strpos+1
			end if
			if (mid(contents, strpos, 1)=vbLf) then
				msg "Found additional LF at " & strpos
				strpos = strpos+1
			end if
		loop
	else
		msg "not found"
	end if
	FindStrEndInFile=strpos
end function

'вставляет данные в указанную позицию в файле
'вставка в позицию 6 означает, что данные будут вставлены начиная с
'6го байта, т.е. исходный файл попилится на части 1-5 байт и 6-до конца
'между этими частями вставляем новые данные
function InsertInFile(ByVal FileName, ByVal InsertPos, ByVal InsertData)
	Dim content,before,after
	content=GetFile(FileName)
	before=left(content,InsertPos-1)
	after=mid(content,InsertPos)
	WriteFile FileName, before & InsertData & after
end function

'заменяет в файле все вхождения needle на Replace
function ReplaceInFile(ByVal FileName, ByVal Needle, ByVal Replacement)
	WriteFile FileName, replace(GetFile(FileName),Needle,Replacement)
end function



'Процедура множественного поиска в файле
'Передаются параметры:
'FileName - итак понятно
'Dict - Словарь с фразами для поиска
'Prefix - префикс для индексов искомых слов
'NeedFound - как нужно искать? findOne/findAll/missOne/missAll
'если префикс например "findme", то будет искать Dict(findme0),Dict(findme1)...
'пока в словаре будут находиться слова с нужными индексами
function MultiFindInFile(ByVal FileName, ByVal Dict, ByVal Prefix, ByVal needFound)

'защита режима поиска от исключений
	needFound=lCase(needFound)

	Select Case needFound
	Case "findone"
		msg "Searching for any occurence"
	Case "findall"
		msg "Searching for all key-phrases"
	Case "missone"
		msg "Trying to miss any key-phrase"
	Case "missall"
		msg "Trying to miss all key-phrases"
	case else
		msg "Searching for one occurence"
		needFound="findone"
	End Select


	dim index : index=0
	dim searchin : searchin=getDict(Prefix & index,Dict,false)
	do while (searchin<>false)
		msg "Searching """ & searchin & """ ..."
		if (FindInFile(FileName, searchin)>0) then
			msg " - Found"
			Select Case needFound
			Case "findone"	'нужно было найти один и мы его нашли
				MultiFindInFile=true
				exit function
			Case "missall"	'нужно было не найти ни ододного а мы его нашли
				MultiFindInFile=false
				exit function
			End Select
		else
			Select Case needFound
			Case "findall"	'нужно было найти найти все, а одного уже не хватает
				MultiFindInFile=false
				exit function
			Case "missone"	'нужно было не найти хотя бы один, и вот как раз такой случай
				MultiFindInFile=true
				exit function
			End Select
			msg " - Miss"
		end if
		index=index+1
		searchin=getDict(Prefix & index,Dict,false)
	loop
	Select Case needFound
	Case "findone"	'нужно было найти один а мы дошли до конца - значит ничего не было
		MultiFindInFile=false
	Case "missone"	'нужно было не найти хотя бы один, а мы тут, неудача
		MultiFindInFile=false
	Case else 'раз дошли до сюда значит или нашли все, или пропустили все, как было задумано
		MultiFindInFile=true
	End Select
end function

'проверяет наличие всех файлов из мастер директории в слейв директории (файлы должны быть одинаковые по размеру)
function masterDirCheck(ByVal MasterDir, ByVal SlaveDir)
Dim slaveFName,Subfolder,File
	masterDirCheck=false
	if Not objFSO.FolderExists(MasterDir) then
		Msg ("Folder "&MasterDir&" not found!")
		Exit Function
	end If
	if Not objFSO.FolderExists(SlaveDir) then
		Msg ("Folder "&SlaveDir&" not found!")
		Exit Function
	end If

	For Each File in objFSO.GetFolder(MasterDir).Files
		slaveFName = SlaveDir&"\"&File.Name
		if Not objFSO.FileExists(slaveFName) then
			Msg ("File "&slaveFName&" not found!")
			Exit Function
		end If
		if File.Size <> objFSO.getFile(slaveFName).Size then
			Msg ("File "&slaveFName&" size mismatches original")
			Exit Function
		end If
		wscript.echo slaveFName&" - OK"
	Next
	For Each Subfolder in objFSO.GetFolder(MasterDir).SubFolders
		if Not masterDirCheck(MasterDir&"\"&Subfolder.Name, SlaveDir&"\"&Subfolder.Name) then
			Exit Function
		end If
	Next
	masterDirCheck=true
end function


'создает ярлычок
function createLnk(byVal lnkPath, byVal targetFile, byVal args, byval workDir, byVal icon, byVal descr)

	if icon="" Then
		icon=targetFile
	End if

	if workDir="" Then
		workDir = GetFilePath (targetFile)
	End if
	
	Msg ("Creating link " & lnkPath & " -> " & targetFile & "...")
	Set lnk = WshShell.CreateShortcut(lnkPath)
	lnk.TargetPath = targetFile
	lnk.Arguments = args
	lnk.Description = descr
	'lnk.HotKey = "ALT+CTRL+F"
	lnk.IconLocation = icon
	lnk.WindowStyle = "1"
	lnk.WorkingDirectory = workDir
	on error resume next
	lnk.Save
	if err.number<>0 then
		Msg "Error while saving Shortcut " & lnkPath
		createLnk = false
	end if
	on error goto 0

	Set lnk = Nothing	
end function

'создает ярлычок на рабочем столе
function createDesktopLnk(byVal lnkName, byVal targetFile, byVal args, byval workDir, byVal icon, byVal descr)
	lnkPath=WshShell.SpecialFolders("Desktop") & "\" & lnkName & ".lnk"
	createLnk lnkPath, targetFile, args, workDir, icon, descr
end function

'создает ярлычок на рабочем столе
function createAllUsersDesktopLnk(byVal lnkName, byVal targetFile, byVal args, byval workDir, byVal icon, byVal descr)
	lnkPath=WshShell.SpecialFolders("AllUsersDesktop") & "\" & lnkName & ".lnk"
	createLnk lnkPath, targetFile, args, workDir, icon, descr
end function

'контролирует наличие ярлычка (если add==true, то добавляет или корректирует, иначе проверяет чтоб такого ярлычка не было)
function ctrlDesktopLnk(byVal lnkName, byVal targetFile, byVal args, byval workDir, byVal icon, byVal descr, byVal add)
	lnkPath=WshShell.SpecialFolders("Desktop") & "\" & lnkName & ".lnk"
	if add then
		createLnk lnkPath, targetFile, args, workDir, icon, descr
	else
		if objFSO.fileExists(lnkPath) then
			Msg ("Removing link " & lnkPath & "...")
			objFSO.deleteFile(lnkPath)
		else 
			'Msg ("Ignoring link " & lnkPath & "...")
		end if
	end if
end function


'сравнивает два файла по размеру и дате
function compareFilesDateSize(byVal strFile1, byVal strFile2)
	compareFilesDateSize=false
        if (not objFSO.fileExists(strFile1)) then
		Msg ("Compare fail - file not exist: " & strFile1)
		exit function
	end if

        if (not objFSO.fileExists(strFile2)) then
		Msg ("Compare fail - file not exist: " & strFile2)
		exit function
	end if

	dim objFile1,objFile2
	set objFile1=objFSO.getFile(strFile1)
	set objFile2=objFSO.getFile(strFile2)
	if (not((objFile1.size = objFile2.size) and (objFile1.dateLastModified = objFile2.dateLastModified))) then
		Msg ("Comparison of " & strFile1 & " vs " & strFile2 & " DIFF: date/size mistmatch!")
		exit function
	end if

	compareFilesDateSize=true
end function



'сравнивает мастердиректорию со слейв и проверяет что весь мастер есть в слейв (сравнение файлов по размеру)
function dirMasterSlaveCompare(byVal master, byVal slave)
	Set objMaster = objFSO.GetFolder(master)
	Set objSlave = objFSO.GetFolder(slave)
	dirMasterSlaveCompare=false
	For Each strFile In objMaster.Files
	    	secFileName = slave & "\" & strFile.Name
	        if (objFSO.fileExists(secFileName)) then
			set secFile=objFSO.getFile(secFileName)
			if (not secFile.size = strFile.size) then
				Msg ("Comparison of " & strFile.Name & " in " & master & " -> " & slave & " failed: size mistmatch!")
				exit function
			end if
		else
			Msg ("Comparison of " & strFile.Name & " in " & master & " -> " & slave & " failed: file not found in slave dir.")
			exit function
		end if
	Next	
	For Each strDir In objMaster.SubFolders
	    	secDirName = slave & "\" & strDir.Name
	        if (objFSO.folderExists(secDirName)) then
			if (not dirMasterSlaveCompare(strDir.path, secDirName)) then
				exit function
			end if
		else
			Msg ("Comparison of dir " & strDir.Name & " in " & master & " -> " & slave & " failed: dir not found in slave dir.")
			exit function
		end if
	Next
	dirMasterSlaveCompare=true
end function

'сравнивает мастердиректорию со слейв и проверяет что весь мастер есть в слейв (сравнение файлов по размеру)
function dirMasterSlaveCopy(byVal master, byVal slave)
	If (not objFSO.FolderExists(master)) then
		dirMasterSlaveCopy=false
		exit function
	end if

	Set objMaster = objFSO.GetFolder(master)
	CheckDir slave
	dirMasterSlaveCopy=false
	Set objSlave = objFSO.GetFolder(slave)
	For Each strFile In objMaster.Files
	    	secFileName = slave & "\" & strFile.Name
	        if (objFSO.fileExists(secFileName)) then
			set secFile=objFSO.getFile(secFileName)
			if (not((secFile.size = strFile.size) and (secFile.dateLastModified = strFile.dateLastModified))) then
				Msg ("Comparison of " & strFile.Name & " in " & master & " -> " & slave & " failed: size/date mistmatch!")
				safeCopy master & "\" & strFile.Name, slave
				dirMasterSlaveCopy=true
			end if
		else
			safeCopy master & "\" & strFile.Name, slave
			dirMasterSlaveCopy=true
		end if
	Next	
	For Each strDir In objMaster.SubFolders
		if (dirMasterSlaveCopy(strDir.path, slave & "\" & strDir.Name)) then
			dirMasterSlaveCopy=true
		end if
	Next
end function


'все что есть в мастере - переносит в слейв
function dirMasterSlaveMove(byVal master, byVal slave)
	Msg "Moving " & master & " -> " & slave
	If (not objFSO.FolderExists(master)) then
		dirMasterSlaveMove=false
		exit function
	end if
	dim objMaster, objSlave, objFile, objDir

	Set objMaster = objFSO.GetFolder(master)
	CheckDir slave

	dirMasterSlaveMove=true
	'Msg ("Moving files in " & master)
	For Each objFile In objMaster.Files
		if (xCopyFile(master & "\" & objFile.Name, slave) = 0) then
			'если копирование прошло успешно - удаляем исходный файл
			objFSO.DeleteFile(master & "\" & objFile.Name)
		else
			msg("Error moving " & master & "\" & objFile.Name & " -> " & slave )
			dirMasterSlaveMove=false
		end if
	Next	

	'Msg ("Moving folders in " & master)
	For Each objDir In objMaster.SubFolders
		'msg("trying to move " & objDir.Path)
		if (dirMasterSlaveMove(objDir.Path , slave & "\" & objDir.Name) = false) then
			'тут я словил удивительнейший глюк, на который потратил несколько часов но так и не победил
			'обращаться к переменной objDir после рекурсивного вызова функции нельзя - vbScript вылетает с ошибкой "путь не найден"
			'подозреваю что под путем имеется в виду поля объекта через точку, но это не точно. Для дальнейшего расследования
			'надо бы написать var_dump и посмотреть что там внутри objDir после выхода.
			'в общем на текущий момент следующее обращение только после повторной инициализации (следующая итерация цикла)
			'msg("wow! " & objDir.Name & " moved")
			'msg("Error moving " & master & "\" & objDir.Name & "->" & slave & "\" & objDir.Name)
			dirMasterSlaveMove=false
		end if
	Next
	
	'если в процессе были ошибки - выходим
	if (not dirMasterSlaveMove) then
		exit function
	end if

	'проверяем что все удалено из директории
	set objMaster = objFSO.GetFolder(master)
	For Each objFile In objMaster.Files
		Msg("Err: file " & objFile.Name & " found in " & master & " after move ")
		dirMasterSlaveMove=false
		exit function
	Next
	For Each objDir In objMaster.SubFolders
		Msg("Err: file " & objDir.Name & " found in " & master & " after move ")
		dirMasterSlaveMove=false
		exit function
	Next

	'Если дошли до сюда не вылетев - значит в папке пусто. удаляем ее
	objFSO.DeleteFolder(master)
	'Msg "Moving " & master & " -> " & slave & " done."
	
end function



'сравнивает мастердиректорию со слейв и проверяет что весь мастер есть в слейв 
'(сравнение файлов по размеру и дате)
'все лишнее что есть на слейве но нет на мастере - удаляет
'возвращает признак того, что были сделаны изменения
function dirMasterSlaveSync(byVal master, byVal slave)
	dim objMaster,_
		objSlave,_
		ignoreLocalFiles,_
		keepLocalFiles,_
		objFile,_
		objDir,_
		secFile,_
		secFileName,_
		secDirname,_
		strDir
	Set objMaster = objFSO.GetFolder(master)
	dirMasterSlaveSync=false
	'удалять файлы на слейве, которых нет на мастере
	keepLocalFiles=false
	'файлы-флаги, которые не нужные на слейве
	Set ignoreLocalFiles = CreateObject("Scripting.Dictionary")
	ignoreLocalFiles.Add ".sync-skip-dir",true
	ignoreLocalFiles.Add ".sync-skip-local-files",true
	ignoreLocalFiles.Add ".sync-keep-slave-files",true

	'проверяем необходимость полностью пропустить синхронизацию папки
	DebugMsg "Comparing " & master & " vs " & slave & " ..."
	if objFSO.fileExists(master & "\.sync-skip-dir") then
		Msg "Skipping " & slave & " because of flag file in it"
		unset(ignoreLocalFiles)
		unset(objMaster)
		exit function
	end if

	'проверяем наличие файла-списка с исключениями вложенных в нее файлов
	if objFSO.fileExists(master & "\.sync-skip-local-files") then
		dim file
		Set file = objFSO.OpenTextFile (master & "\.sync-skip-local-files", 1)
		Do Until file.AtEndOfStream
			line = file.Readline
			ignoreLocalFiles.Add line,true
		Loop
		unset(file)
		Msg "Got exclusions for " & slave & " because of exclusions list file in it"
		if DEBUGMODE then
			dim i,arrItems
			arrItems=ignoreLocalFiles.Keys
			For i = 0 To ignoreLocalFiles.count - 1
			    debugMsg arrItems(i)
			Next
		end if
	end if

	'признак того, что файлы в слейве не будут удаляться если их нет на мастере
	keepLocalFiles=objFSO.fileExists(master & "\.sync-keep-slave-files") 


	'----- Конец инициализации - работаем

	'проверяем папку
	CheckDir slave
	Set objSlave = objFSO.GetFolder(slave)

	DebugMsg "dirMasterSlaveSync forward files passage (" & master & ")"
	'прямая проходка (копирование с мастера на слейв)
	For Each objFile In objMaster.Files
		if (not ignoreLocalFiles.Exists(objFile.Name)) then
		    	secFileName = slave & "\" & objFile.Name
		        if (objFSO.fileExists(secFileName)) then
				set secFile=objFSO.getFile(secFileName)
				if (not((secFile.size = objFile.size) and (secFile.dateLastModified = objFile.dateLastModified))) then
					Msg ("Comparison of " & objFile.Name & " in " & master & " -> " & slave & " failed: date/size mistmatch!")
					call safeCopy(master & "\" & objFile.Name, slave)
					DebugMsg "dirMasterSlaveSync safeCopy.complete"
					dirMasterSlaveSync=true
				end if
				unset(secFile)
			else
				call safeCopy (master & "\" & objFile.Name, slave)
				dirMasterSlaveSync=true
			end if
		end if
	Next	
	
	DebugMsg "dirMasterSlaveSync backward files passage (" & master & ")"
	'обратная проходка (удаление на слейве того чего нет на мастере)
	if not keepLocalFiles then
		For Each objFile In objSlave.Files
		    	secFileName = master & "\" & objFile.Name
			if (not objFSO.fileExists(secFileName)) and not ignoreLocalFiles.Exists(objFile.Name) then
				Msg ("Comparison of " & objFile.Name & " in " & master & " -> " & slave & " failed: removed from master repo")
				objFSO.DeleteFile(slave & "\" & objFile.Name)
				dirMasterSlaveSync=true
			end if
		Next
	end if

	DebugMsg "dirMasterSlaveSync forward dirs passage (" & master & ")"
	'прямая проходка по подпапкам
	For Each objDir In objMaster.SubFolders
		DebugMsg "Testing " & objDir.Name
		if (not ignoreLocalFiles.Exists(objDir.Name)) then
			DebugMsg "Testing passed " & objDir.Name
			if (dirMasterSlaveSync(objDir.path, slave & "\" & objDir.Name)) then dirMasterSlaveSync=true
		else
			DebugMsg "Testing failed " & objDir.Name
		end if
	Next

	DebugMsg "dirMasterSlaveSync backward dirs passage (" & master & ")"
	'обратная проходка по подпапкам
	if not keepLocalFiles then
		For Each objDir In objSlave.SubFolders
			secDirName = master & "\" & objDir.Name
			if (not objFSO.folderExists(secDirName)) then
				Msg ("Comparison of " & objDir.Name & " in " & master & " -> " & slave & " failed: removed from master repo")
				objFSO.DeleteFolder slave & "\" & objDir.Name ,true
				dirMasterSlaveSync=true
			end if
		Next
	end if
	DebugMsg "dirMasterSlaveSync complete (" & master & ")"
	unset(ignoreLocalFiles)
	unset(objMaster)
	unset(objSlave)
	unset(objDir)
	unset(objFile)
end function
'' SIG '' Begin signature block
'' SIG '' MIIIXwYJKoZIhvcNAQcCoIIIUDCCCEwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' +BC/CK4A8aEtOCHzo/TqNjZHnvDsWIUcqGqJ6R42CI2g
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
'' SIG '' AQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIFks8ybDPMhE
'' SIG '' dL7PHJSCEkga10kaVVMnoPXy+89aalFXMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAJit4SphGtoacyizazzoXrKD60uJ9YdC
'' SIG '' iYy9h74dHyLERn6RI76Yf3kskbWaim2qnMnS5eJ/92pA
'' SIG '' XvOLrpWE9P66Vj7Lnt4fPeFyY+3Q8ZehLqmNFpApsMRe
'' SIG '' VM6bchPm9V0J9wxvwzg0Eyq/piKczsDE6TggcDLnuDkp
'' SIG '' 7DRkpxfx+CmFXs8ky5YdwU7lLsRMpDM5pUDSpeFBIQcs
'' SIG '' tQx3da/TzOjOBJVwUY9A76g0hxluN/nenWfikzZHSc1Q
'' SIG '' 34Hp1U4p0LXkrvqFlGmo63cWSz8rsy0QYqBXYutQD+Ix
'' SIG '' PPQMqmDiYS1kc3Pmh68WydCz5cRlXvk7/zgCb8tsrDj9mtI=
'' SIG '' End signature block
