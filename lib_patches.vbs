option Explicit
'-----------------------------------------------------------------------
'PATCHES ROUTINE             -------------------------------------------
'-----------------------------------------------------------------------
'���������� ��������� ������ ���� ������, � �������� ����������� �������� ���:
'patch{Method}(patch_var) 
'������� - ��� ����� ������� ���� �����������
'� ���������� �������� ������� ������������ ���������
'����������� ��� ������ ����������� ���������� ����� ����� ����� �������� ���� ���� ����� ���������

' �������:
'patchCheckFileVariables patch_start_nxmanager,ftype_bat
'patchCopyDir(patch_NX_fonts)
'patchReplaceInFile(patch_Teamcenter_rcs)


'���������� ������ �� �������
function getDict(idx,dict,def)
	idx=LCase(idx)
	if dict.exists(idx) then
		getDict=dict(idx)
	else
		getDict=def
	end if
end function

'��������� ������� ���� � ���������� ����������� ����
function patchStructCk(ByVal Patch, ByVal ckField)

	if (getDict(ckField,Patch,false) = false) then
		Msg "ERROR: patchStructCk: incorrect patch struct passed (no " & ckField & ")!"
		patchStructCk = false
	else
		patchStructCk = true
	end if
end function


Function patchAppliance(ByVal Patch)
'�������� ������������ �����
'block_file,			������ ������������� (���� ������) ��� �������� �����
'block_file_size,		���� ������ �� ����������� ���� ����������� ������ ��� ���������� �������
'in_blockfile_search0,1,...	������ ������������� � ���
'in_blockfile_search_type 	����� ��������� ��� ������ �������� ����� (�� MultiFindInFile)
'presence_check,		������ �������������� (���� ������) � �������� �����
'in_presence_search0,1,...	������� ����������� � ���
'in_presence_search_type 	����� ��������� ��� ������ �������� ����� (�� MultiFindInFile)
	dim blockmasterdir : blockmasterdir = getDict("block_master_dir",patch,false)
	If (not (blockmasterdir=false)) then
		dim slavedir : slavedir=getDict("block_slave_dir",patch,false)
		If (not (slavedir=false)) then '���� � ��� ����� ������
			msg "Checking all files from" & blockmasterdir & " existance in " & slavedir & " ..."
			if masterDirCheck(blockmasterdir, slavedir) then
				msg (" - all match. Patch blocked! ")
				patchAppliance=false
				Exit function
			else
				msg (" - difference found!")
			end if
		end if
	end if	

	dim addchecks : addchecks=false
	dim blockfile : blockfile = getDict("block_file",patch,false)
	If (not (blockfile=false)) then
		msg "Checking " & blockfile & " existance ..."
		if (objFSO.FileExists(blockfile)) Then
			msg (" - Found")

			If (not (getDict("block_file_size",patch,false)=false)) then '���� � ��� ����� ������
				addchecks=true
				msg ("Checking block file size ... ")
				dim f : set f=objFSO.GetFile(blockfile)
				if (f.size=Patch("block_file_size")) then
					msg (" - size match. Patch blocked! ")
					patchAppliance=false
					Exit function
				else
					msg (" - size mismatch")
				end If
			End If

			If (not (getDict("in_blockfile_search0",patch,false)=false)) then '���� � ��� ���� ������ ������ ������
				addchecks=true
				msg ("Checking file contents ... ")
				if MultiFindInFile(blockfile,patch,"in_blockfile_search", getDict("in_blockfile_search_type",patch,"findone")) then
				'� ����� ���� ���������� ����� �������� �������� (�� ��������� - ����� ���� ���� ����������)
					msg "SKIP: Block phrase Found. No need to patch."
					patchAppliance=false
					Exit function
				else
					msg "Not blocking phrases found. Need to patch!"
				end if
			end if

			if not addchecks then '������ ������� ������������� �������� �� ���� - ���������� ��� �������
				Msg "SKIP: " & blockfile & " exists. Nothing to do"
				patchAppliance=false
				Exit function
			end if

		else
			msg "File " & blockfile & " not found."
		end if
	end if

	dim checkfile : checkfile = getDict("presence_check",patch,false)
	If (not (checkfile=false)) then
		msg "Checking " & checkfile & " existance ..."
		if (not objFSO.FileExists(checkfile)) Then
			Msg "SKIP: " & checkfile & " not found. Nothing to do."
			patchAppliance=false
			Exit function
		else
			msg "File " & checkfile & " found."
			If (not (getDict("in_presence_search0",patch,false)=false)) then
				msg ("Found. Checking file contents ... ")
				if MultiFindInFile(checkfile,patch,"in_presence_search", getDict("in_presence_search_type",patch,"findall")) then
					'� ����� ���� ���������� ����� �������� �������� (�� ��������� - ����� ��� ����������)
					msg "Found! Need to patch!"
				else
					msg "SKIP: Not found. No need to patch."
					patchAppliance=false
					Exit function
				end if
			end if
		end if
	end if
	patchAppliance=true
End Function


sub patchReplaceInFile(ByVal Patch)
'������ ��������� ������
	if ((not patchStructCk(Patch,"file_to_patch"))_
	or 	(not patchStructCk(Patch,"replace"))_
	or 	(not patchStructCk(Patch,"with"))_
	) then
		Msg "SKIP: patch incorrect"
		exit sub
	end if

	dim file : file = getDict("file_to_patch",patch,false)
	msg "Checking patch on " & file
	If (not objFSO.FileExists(file)) Then
		Msg "SKIP: file " & file & " not found. Nothing to patch."
		exit sub
	else
		msg "File " & file & " found. Searching ... "
	end if

	if (FindInFile(file, patch("replace"))=0) then
		msg "SKIP: File " & file & " already patched"
		exit sub
	else
		msg "Patching ... "
		ReplaceInFile file, patch("replace"), patch("with")
		if (FindInFile(file, patch("replace"))=0) then
			msg " - Success"
		else
			msg " - No luck"
		end if
	end if
End Sub



sub patchTextFile(ByVal Patch)
'������ ��������� ������
	if ((not patchStructCk(Patch,"file_to_patch"))_
	or 	(not patchStructCk(Patch,"insert_string"))_
	or 	(not patchStructCk(Patch,"insert_after"))_
	) then
		Msg "SKIP: patch incorrect"
		exit sub
	end if

	dim file : file = getDict("file_to_patch",patch,false)
	msg "Checking patch on " & file
	If (not objFSO.FileExists(file)) Then
		Msg "SKIP: file " & file & " not found. Nothing to patch."
		exit sub
	else
		msg "File " & file & " found. Searching ... "
	end if

	dim insert_string : insert_string = getDict("insert_string",patch,false)
	if (FindInFile(file, insert_string)>0) then
		msg "SKIP: File " & file & " already patched"
		exit sub
	else
		msg "Need to patch. Searching pos ... "
	end if

	dim insert_before : insert_before = getDict("insert_after",patch,false) '

	dim insertPos : insertPos = FindStrEndInFile(file, insert_before)
	if (insertPos>0) then
		Msg	"Inserting at " & insertPos
		InsertInFile file,InsertPos, insert_string & vbCrLf
	End if
End Sub



sub patchCheckFileVariables(ByVal Patch,ByVal Ftype)
'������ ������ ����������
	if ((not patchStructCk(Patch,"file_to_patch"))_
	or 	(not patchStructCk(Patch,"var0"))_
	) then
		Msg "SKIP: patch incorrect"
		exit sub
	end if

	dim file : file = getDict("file_to_patch",patch,false)
	msg "Checking patch on " & file
	If (not objFSO.FileExists(file)) Then
		Msg "SKIP: file " & file & " not found. Nothing to patch."
		exit sub
	else
		msg "File " & file & " found."
	end if

	If (not CheckFileTypeDescr(Ftype)) Then
		Msg "SKIP: given filetype description is incorrect."
		exit sub
	end if

	'��� ���������� ���������� �� ����� ������� � var0
	dim index : index=0
	dim searchin : searchin=getDict("var" & index,Patch,false)
	dim testVar,testSec,testVal,secPos,current,placeAfter
	do while (searchin<>false)
		'wscript.echo "patchCheckFileVariables: parsing " & searchin
		testVar=GetVariableName(searchin,Ftype("eq"))
		if Len(testVar)>0 then
			testVal=GetVariableVal(searchin,Ftype("eq"))
			'���� � ����������� ���������� ���� �� ���, ���� �� ����, �� ��� ����� ��������� ������ - ������. ����� ������������ �������� � REG ������
			secPos=instrRev(testVar,"\",-1,vbTextCompare)
			if secPos>0 then
				testSec=Left(testVar,secPos-1)
				testVar=Right(testVar,Len(testVar)-secPos)
			else
				testSec=""
			end if
			msg_ "Searching if [" & testSec & "]\""" & testVar & """ is set to " & testVal & " ... "
			
			current=conffile_get(file, FType, testSec, testVar, unset_me)
			if current = testVal then
				msg_n "- Yes"
			else
				'����� ������� "var" & index & "place_after_str" - ������ ����� ������� �������� ���������� - �� ������ �������� ������, � �� INI
				placeAfter=getDict("var" & index & "_place_after",Patch,"")
				if placeAfter="" then
					msg__ "- No. Changing ... "
					if conffile_set(file, FType, testSec, testVar, testVal, "changed by "&scrName&" ver "&scrVer&" at "&Date&" "&time) then
						msg_n "- Success"
					else
						msg_n "- No luck"
					end If
				else
					msg__ "- No. Changing (ins)... "
					if textfile_set_after(file, FType, testSec, testVar, testVal, "added by "&scrName&" ver "&scrVer&" at "&Date&" "&time,placeAfter) then
						msg_n "- Success "
					else
						msg_n "- No luck"
					end If
				end if
			end if
		else
			msg("patchCheckFileVariables: parsing " & searchin & " error! can not find variable name!")
		end if
		index=index+1
		searchin=getDict("var" & index, Patch, false)
	loop

End Sub


sub patchRemoveApp(ByVal Patch)
'������ - �����������
'patch_old_vis.add "presence_check",	"C:\Siemens\Teamcenter10.1\Visualization"
'patch_old_vis.add "remove_app",	"Teamcenter Visualization 10.1 64-bit"
	if (not patchStructCk(Patch,"remove_app")) then
		Msg "SKIP: patch incorrect"
		exit sub
	end if

	if not patchAppliance(Patch) then
		exit sub
	end if

	safeRun "wmic product where name=""" & patch("remove_app") & """ call uninstall"
End Sub


sub patchInstallMsi(ByVal Patch)
'������ - MSI ���������
'patch_otw_vis.add "block_file",	"C:\Siemens\Visualization\etc\copyright.txt"
'patch_otw_vis.add "msi_file",	vis10_1_10_msi
'patch_otw_vis.add "msi_params",	vis10_1_10_params
	if (not patchStructCk(Patch,"msi_file")) then
		Msg "SKIP: patch incorrect"
		exit sub
	end if

	if not patchAppliance(Patch) then
		exit sub
	end if

	if (not objFSO.FileExists(patch("msi_file"))) Then
		Msg "SKIP: " & patch("msi_file") & " not found. Nothing to do"
		exit sub
	else
		msg "File " & patch("msi_file") & " found."
	end if

	safeRun "msiexec /i " & patch("msi_file") & " " & getDict("msi_params",patch,"") & " /qn"
End Sub

sub patchInstallExe(ByVal Patch)
'������ - EXE ���������
'patch_otw_vis.add "block_file",	"C:\Siemens\Visualization\etc\copyright.txt"
'patch_otw_vis.add "msi_file",		vis10_1_10_msi
'patch_otw_vis.add "msi_params",	vis10_1_10_params
	if (not patchStructCk(Patch,"exe_file")) then
		Msg "SKIP: patch incorrect"
		exit sub
	end if

	if not patchAppliance(Patch) then
		exit sub
	end if

	if (not objFSO.FileExists(getDict("exec_dir",patch,"")&patch("exe_file"))) Then
		Msg "SKIP: " &getDict("exec_dir",patch,"")&patch("exe_file") & " not found. Nothing to do"
		exit sub
	else
		msg "File " &getDict("exec_dir",patch,"")&patch("exe_file") & " found."
	end if

	if (getDict("exec_dir",patch,false)<>false) then
		if (not objFSO.FolderExists(patch("exec_dir"))) Then
			Msg "SKIP: " & patch("exec_dir") & " not found. Nothing to do"
			exit sub
		else
			safeExec patch("exe_file"), getDict("exe_params",patch,""), getDict("exec_dir",patch,"")
		end if
	else
		safeRun patch("exe_file") & " " & getDict("exe_params",patch,"")
	end if

End Sub

sub patchCopyDir(ByVal Patch)
'������ - �������� ���������� ���� �� ����� �������������, ���� ��� ������ (�� �������)
'patch_vis_view_jar.add "copy_dir",			"\\RTS-DEVELOP\dfs\install\_Scripts\TC\azimutclient_template\"
'patch_vis_view_jar.add "copy_to",			"c:\Siemens\Teamcenter\OTW10\rac"

	if (not patchStructCk(Patch,"copy_dir") or not patchStructCk(Patch,"copy_to")) then
		Msg "SKIP: patch incorrect"
		exit sub
	end if

	if not patchAppliance(Patch) then
		exit sub
	end if

	if (not objFSO.FolderExists(patch("copy_dir"))) Then
		Msg "SKIP: " & patch("copy_dir") & " not found. Nothing to do"
		exit sub
	else
		if (getDict("block_master_slave",Patch,false)) then
			changes=dirMasterSlaveCopy(patch("copy_dir"), patch("copy_to"))
			if(not(getDict("run_after_patch",Patch,false)=false) and changes=true) then
				msg "Changes made to " & patch("copy_to") & ". Running patch_after script "
				safeRun getDict("run_after_patch",Patch,false)
			end if
		else
			msg "File " & patch("copy_dir") & " found. Patching ... "
			safeRun "%windir%\system32\XCOPY.exe /Y /C /F /R /H /E /I """ & patch("copy_dir") & """ """ & patch("copy_to") & """"' >> " & logFPath
			msg "done"
		end if
	end if

End Sub

sub patchSyncDir(ByVal Patch)
'������ - �������������� ����� ���������� ��� ����� ��� ��������������� ������ ����������
'��������� �����/���������� �����, ������� ������
'patch_vis_view_jar.add "copy_dir",			"\\RTS-DEVELOP\dfs\install\_Scripts\TC\azimutclient_template\"
'patch_vis_view_jar.add "copy_to",			"c:\Siemens\Teamcenter\OTW10\rac"
	dim changes
	if (not patchStructCk(Patch,"copy_dir") or not patchStructCk(Patch,"copy_to")) then
		Msg "SKIP: patch incorrect"
		exit sub
	end if

	if not patchAppliance(Patch) then
		exit sub
	end if

	if (not objFSO.FolderExists(patch("copy_dir"))) Then
		Msg "SKIP: " & patch("copy_dir") & " not found. Nothing to do"
		exit sub
	else
		changes=dirMasterSlaveSync(patch("copy_dir"), patch("copy_to"))
		if(not(getDict("run_after_patch",Patch,false)=false) and changes=true) then
			msg "Changes made to " & patch("copy_to") & ". Running patch_after script "
			safeRun getDict("run_after_patch",Patch,false)
		end if
	end if

End Sub

sub patchCopyFile(ByVal Patch)
'������ - �������� ����������
'patch_otw_vis.add "replace_file",	"c:\Siemens\Teamcenter\OTW10\rac\plugins\SingleEmbeddedViewer.jar"
'patch_otw_vis.add "with_file",	"c:\Siemens\Visualization\Program\SingleEmbeddedViewer.jar "

	if (not patchStructCk(Patch,"replace_file") or not patchStructCk(Patch,"with_file")) then
		Msg "SKIP: patch incorrect"
		exit sub
	end if

	if (not objFSO.FileExists(patch("with_file"))) Then
		Msg "SKIP: " & patch("with_file") & " not found. Nothing to do"
		exit sub
	else
		msg "File " & patch("with_file") & " found."
		if (not objFSO.FileExists(patch("replace_file"))) Then
			'����� ��� - ������ ��������
			msg "File " & patch("replace_file") & " not found. Patching ... "
			'safeRun "%comspec% /C COPY /Y """ & patch("with_file") & """ """ & patch("replace_file") & """"
			safeCopy patch("with_file") , patch("replace_file")
			msg "done"
		else
			'���� ����, ���� ��������
			dim f1 : set f1=objFSO.GetFile(patch("with_file"))
			dim f2 : set f2=objFSO.GetFile(patch("replace_file"))
			if (f1.size<>f2.size) then
				msg "File " & patch("replace_file") & " has different size. Patching ... "
				'safeRun "%comspec% /C COPY /Y """ & patch("with_file") & """ """ & patch("replace_file") & """"
				safeCopy patch("with_file") , patch("replace_file")

				msg "done"
			else
				msg "File " & patch("replace_file") & " same size. Nothing to do."
				exit sub
			end if
		end if
	end if

End Sub

sub patchFontInstall(ByVal Patch)
'������ - �������� ������ �� ����� from_dir � ����� C:\windows\fonts � ������������ ��

	if (not patchStructCk(Patch,"from_dir")) then
		Msg "SKIP: patch incorrect"
		exit sub
	end if

	if not patchAppliance(Patch) then
		exit sub
	end if
	'���� �������� ����� ����
	if (not objFSO.FolderExists(patch("from_dir"))) Then
		Msg "SKIP: " & patch("from_dir") & " not found. Nothing to do"
		exit sub
	else
		dim objMaster,objSlave,objShellApp,objFont
		strMaster=Patch("from_dir")
		Set objShellApp = CreateObject("Shell.Application")
		Set objSlave=objShellApp.Namespace( &H14 )
		Set objMaster = objFSO.GetFolder(strMaster)
		strSlave=objSlave.Self.Path
		'Set objSlave = objFSO.GetFolder(strSlave)
		'���������� ��� �����
		For Each objFile In objMaster.Files
			'Msg "Checking "&objFile.Name&"..."
			'���� ����� ������ 4� � ��� ����� ���������� � � ����� ��������� ������������ ����� ������
			if ( _
				(len(objFile.Name)>4) _
				AND _
				(_
					lcase(right(objFile.Name,4))=".ttf"_
					OR _
					lcase(right(objFile.Name,4))=".fon"_
					OR _
					lcase(right(objFile.Name,4))=".otf"_
				)_
			) then
				if (Patch.Exists(objFile.Name)) then
					'���� ����������� ����� ��� � ������� �����
				        if (not compareFilesDateSize(strMaster & "\" &objFile.Name , strSlave & "\" & objFile.Name)) then
						safeCopy strMaster & "\" & objFile.Name, strSlave
					end if
					regCheck "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts\" & patch(objFile.Name), "REG_SZ", objFile.Name
				else
					msg "Not font name for "&objFile.Name
				end if
			end if
		next
	end if
End Sub

'
