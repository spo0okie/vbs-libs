'FILE ROUTINE ----------------------------------------------------------
'option explicit

'����������� �����
Function xCopyFile(ByVal fromf, ByVal tof)
	'/C	���������� ������.
	'/f	���������� ����� �������� � ������� ������ ��� �����������.
	'/i	���� Source �������� ��������� ��� �������� �������������� �����, � ���������� �� ����������, ������� xcopy ������������, 
	'	��� � ���� ���������� ������� ��� �������� � ��������� ����� �������. ����� ������� xcopy �������� ��� ��������� ����� � ����� �������. 
	'	�� ��������� ������� xcopy ��������� �������, �������� �� ���������� ������ ��� ���������.
	'/r	�������� �����, ������� �������� ������ ��� ������.
	'/h	�������� ����� � ���������� ������� � ��������� ������. �� ��������� ������� xcopy �� �������� ������� ��� ��������� �����.
	'/y	��������� ������ �� ������������� ���������� ������������� �������� �����.
	'/z	��������� ����������� �� ���� � ��������������� ������.
	dim command: command="%windir%\system32\XCOPY.exe /Y /C /F /R /H /I /Z /E """ & fromf & """ """ & tof & """"
	debugMsg "Running " & command
	'on error resume next
		'Err.Clear
		xCopyFile=wshShell.run(command,0,true)
		'Err.Clear
	'on error goto 0
	debugMsg "Complete"
End Function

'����������� �����
Sub safeCopy(ByVal fromf, ByVal tof)
	msg_ "Copying " & fromf & " to " & tof 
	dim ret : ret=xCopyFile(fromf,tof)
	msg__ " - return code: " & ret & vbCrLf
	'safeCopy=ret
End sub

'�������� �����, ���� �� ����
Sub safeDelete(ByVal FName)
	msg_ "Deleting " & Fname & " ..."
	if objFSO.fileExists(Fname) then
		objFSO.deleteFile(Fname)
		msg__ " complete"
	else
		msg__ " not exists"
	end if 
End Sub


'���������� ���� � ����� ��� ����� �����
Function GetFilePath(ByVal FileName)
	GetFilePath=Left(FileName,InStrRev(FileName,"\"))
End Function

'���������� �������� ����, ������� ����� �������� ������� ����� �����
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

'���������� ������ ����, ������� �� ���� �������� ������� ����� �����
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

'���������� �������� ����, ������� ����� �������� ������� ����� �����
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

'�������� ������� ���������� � �������� � ������ ����������
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


'���� ������� ������ � �����
function FindInFile(ByVal FileName, ByVal Needle)
	FindInFile = InStr(1, GetFile(FileName) , Needle, vbTextCompare)
end function

'�� �� ��� � ������, ������ ���� ����� ������ ������
'�� ���� �� ������� ���� CR ��� LF � ���� ��� ������������������ ��
'CR LF �������������
function FindStrEndInFile(ByVal FileName, ByVal Needle)
	msg "Searching " & needle
	dim contents : contents = GetFile(FileName)
	dim strpos : strpos = InStr(1, contents , Needle, vbTextCompare)
	dim crpos, lfpos
	if strpos>0 then
		msg "Found at " & strpos
		strpos=strpos+len(needle)-1
		'���� ����� ����� ������ ����� ������
		crpos=InStr(strpos, contents , vbCr, vbTextCompare)
		lfpos=InStr(strpos, contents , vbLf, vbTextCompare)
		'���������, ����� ������-�� �� ��� �� ����
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

'��������� ������ � ��������� ������� � �����
'������� � ������� 6 ��������, ��� ������ ����� ��������� ������� �
'6�� �����, �.�. �������� ���� ��������� �� ����� 1-5 ���� � 6-�� �����
'����� ����� ������� ��������� ����� ������
function InsertInFile(ByVal FileName, ByVal InsertPos, ByVal InsertData)
	Dim content,before,after
	content=GetFile(FileName)
	before=left(content,InsertPos-1)
	after=mid(content,InsertPos)
	WriteFile FileName, before & InsertData & after
end function

'�������� � ����� ��� ��������� needle �� Replace
function ReplaceInFile(ByVal FileName, ByVal Needle, ByVal Replacement)
	WriteFile FileName, replace(GetFile(FileName),Needle,Replacement)
end function



'��������� �������������� ������ � �����
'���������� ���������:
'FileName - ���� �������
'Dict - ������� � ������� ��� ������
'Prefix - ������� ��� �������� ������� ����
'NeedFound - ��� ����� ������? findOne/findAll/missOne/missAll
'���� ������� �������� "findme", �� ����� ������ Dict(findme0),Dict(findme1)...
'���� � ������� ����� ���������� ����� � ������� ���������
function MultiFindInFile(ByVal FileName, ByVal Dict, ByVal Prefix, ByVal needFound)

'������ ������ ������ �� ����������
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
			Case "findone"	'����� ���� ����� ���� � �� ��� �����
				MultiFindInFile=true
				exit function
			Case "missall"	'����� ���� �� ����� �� �������� � �� ��� �����
				MultiFindInFile=false
				exit function
			End Select
		else
			Select Case needFound
			Case "findall"	'����� ���� ����� ����� ���, � ������ ��� �� �������
				MultiFindInFile=false
				exit function
			Case "missone"	'����� ���� �� ����� ���� �� ����, � ��� ��� ��� ����� ������
				MultiFindInFile=true
				exit function
			End Select
			msg " - Miss"
		end if
		index=index+1
		searchin=getDict(Prefix & index,Dict,false)
	loop
	Select Case needFound
	Case "findone"	'����� ���� ����� ���� � �� ����� �� ����� - ������ ������ �� ����
		MultiFindInFile=false
	Case "missone"	'����� ���� �� ����� ���� �� ����, � �� ���, �������
		MultiFindInFile=false
	Case else '��� ����� �� ���� ������ ��� ����� ���, ��� ���������� ���, ��� ���� ��������
		MultiFindInFile=true
	End Select
end function

'��������� ������� ���� ������ �� ������ ���������� � ����� ���������� (����� ������ ���� ���������� �� �������)
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


'������� �������
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

'������� ������� �� ������� �����
function createDesktopLnk(byVal lnkName, byVal targetFile, byVal args, byval workDir, byVal icon, byVal descr)
	lnkPath=WshShell.SpecialFolders("Desktop") & "\" & lnkName & ".lnk"
	createLnk lnkPath, targetFile, args, workDir, icon, descr
end function

'������� ������� �� ������� �����
function createAllUsersDesktopLnk(byVal lnkName, byVal targetFile, byVal args, byval workDir, byVal icon, byVal descr)
	lnkPath=WshShell.SpecialFolders("AllUsersDesktop") & "\" & lnkName & ".lnk"
	createLnk lnkPath, targetFile, args, workDir, icon, descr
end function

'������������ ������� ������� (���� add==true, �� ��������� ��� ������������, ����� ��������� ���� ������ ������� �� ����)
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


'���������� ��� ����� �� ������� � ����
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



'���������� ���������������� �� ����� � ��������� ��� ���� ������ ���� � ����� (��������� ������ �� �������)
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

'���������� ���������������� �� ����� � ��������� ��� ���� ������ ���� � ����� (��������� ������ �� �������)
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


'��� ��� ���� � ������� - ��������� � �����
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
			'���� ����������� ������ ������� - ������� �������� ����
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
			'��� � ������ ��������������� ����, �� ������� �������� ��������� ����� �� ��� � �� �������
			'���������� � ���������� objDir ����� ������������ ������ ������� ������ - vbScript �������� � ������� "���� �� ������"
			'���������� ��� ��� ����� ������� � ���� ���� ������� ����� �����, �� ��� �� �����. ��� ����������� �������������
			'���� �� �������� var_dump � ���������� ��� ��� ������ objDir ����� ������.
			'� ����� �� ������� ������ ��������� ��������� ������ ����� ��������� ������������� (��������� �������� �����)
			'msg("wow! " & objDir.Name & " moved")
			'msg("Error moving " & master & "\" & objDir.Name & "->" & slave & "\" & objDir.Name)
			dirMasterSlaveMove=false
		end if
	Next
	
	'���� � �������� ���� ������ - �������
	if (not dirMasterSlaveMove) then
		exit function
	end if

	'��������� ��� ��� ������� �� ����������
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

	'���� ����� �� ���� �� ������� - ������ � ����� �����. ������� ��
	objFSO.DeleteFolder(master)
	'Msg "Moving " & master & " -> " & slave & " done."
	
end function



'���������� ���������������� �� ����� � ��������� ��� ���� ������ ���� � ����� 
'(��������� ������ �� ������� � ����)
'��� ������ ��� ���� �� ������ �� ��� �� ������� - �������
'���������� ������� ����, ��� ���� ������� ���������
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
	'������� ����� �� ������, ������� ��� �� �������
	keepLocalFiles=false
	'�����-�����, ������� �� ������ �� ������
	Set ignoreLocalFiles = CreateObject("Scripting.Dictionary")
	ignoreLocalFiles.Add ".sync-skip-dir",0
	ignoreLocalFiles.Add ".sync-skip-local-files",0
	ignoreLocalFiles.Add ".sync-keep-slave-files",0

	DebugMsg "Comparing " & master & " vs " & slave & " ..."
	if objFSO.fileExists(master & "\.sync-skip-dir") then
		Msg "Skipping " & slave & " because of flag file in it"
		unset(ignoreLocalFiles)
		unset(objMaster)
		exit function
	end if

	if objFSO.fileExists(master & "\.sync-skip-local-files") then
		dim file
		Set file = objFSO.OpenTextFile (master & "\.sync-skip-local-files", 1)
		Do Until file.AtEndOfStream
			line = file.Readline
			ignoreLocalFiles.Add line,0
		Loop
		unset(file)
		Msg "Skipping " & slave & " because of flag file in it"
	end if

	keepLocalFiles=objFSO.fileExists(master & "\.sync-keep-slave-files") 

	CheckDir slave
	Set objSlave = objFSO.GetFolder(slave)

	DebugMsg "dirMasterSlaveSync forward files passage"
	'������ �������� (����������� � ������� �� �����)
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
				safeCopy master & "\" & objFile.Name, slave
				dirMasterSlaveSync=true
			end if
		end if
	Next	
	
	DebugMsg "dirMasterSlaveSync backward files passage"
	'�������� �������� (�������� �� ������ ���� ���� ��� �� �������)
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

	DebugMsg "dirMasterSlaveSync forward dirs passage"
	'������ �������� �� ���������
	For Each strDir In objMaster.SubFolders
		if (dirMasterSlaveSync(strDir.path, slave & "\" & strDir.Name)) then
			dirMasterSlaveSync=true
		end if
	Next

	DebugMsg "dirMasterSlaveSync backward files passage"
	'�������� �������� �� ���������
	if not keepLocalFiles then
		For Each strDir In objSlave.SubFolders
		    	secDirName = master & "\" & strDir.Name
	    	    if (not objFSO.folderExists(secDirName)) then
				Msg ("Comparison of " & strDir.Name & " in " & master & " -> " & slave & " failed: removed from master repo")
				objFSO.DeleteFolder slave & "\" & strDir.Name ,true
				dirMasterSlaveSync=true
			end if
		Next
	end if
	'DebugMsg "dirMasterSlaveSync reuturning " & 
	unset(ignoreLocalFiles)
	unset(objMaster)
	unset(objSlave)
	unset(objDir)
	unset(objFile)
end function