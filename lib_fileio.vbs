'FILE ROUTINE ----------------------------------------------------------

'Копирование файла
Sub safeCopy(ByVal fromf, ByVal tof)
	on error resume next
	dim ret : ret=wshShell.run("%windir%\system32\XCOPY.exe /Y /C /F /R /H /I /Z """ & fromf & """ """ & tof & """",0,true)
	msg "Copying " & fromf & " to " & tof & " - return code: " & ret
	on error goto 0
End sub


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
		msg ("Got " & tmp)
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
	lnk.Save
	Set lnk = Nothing	
end function

'создает ярлычок на рабочем столе
function createDesktopLnk(byVal lnkName, byVal targetFile, byVal args, byval workDir, byVal icon, byVal descr)
	lnkPath=WshShell.SpecialFolders("Desktop") & "\" & lnkName & ".lnk"
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
	Set objMaster = objFSO.GetFolder(master)
	CheckDir slave
	dirMasterSlaveCopy=false
	Set objSlave = objFSO.GetFolder(slave)
	For Each strFile In objMaster.Files
	    	secFileName = slave & "\" & strFile.Name
	        if (objFSO.fileExists(secFileName)) then
			set secFile=objFSO.getFile(secFileName)
			if (not secFile.size = strFile.size) then
				Msg ("Comparison of " & strFile.Name & " in " & master & " -> " & slave & " failed: size mistmatch!")
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