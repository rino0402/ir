<%@LANGUAGE=VBScript%>
<%Option Explicit%>
<% Response.Buffer = True %>
<% Response.Expires = -1 %>
<%
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly	= 0
Const adOpenKeyset	= 1
Const adOpenDynamic	= 2
Const adOpenStatic	= 3

'---- LockTypeEnum Values ----
Const adLockReadOnly		= 1
Const adLockPessimistic 	= 2
Const adLockOptimistic		= 3
Const adLockBatchOptimistic	= 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- CommandTypeEnum Values ----
Const adCmdUnspecified	= -1	' Unspecified type of command 
Const adCmdText			= 1	' Evaluates CommandText as a textual definition of a command or stored procedure call 
Const adCmdTable		= 2	' Evaluates CommandText as a table name whose columns are returned by an SQL query 
Const adCmdStoredProc	= 4	' Evaluates CommandText as a stored procedure name 
Const adCmdUnknown		= 8	' Default. Unknown type of command 
Const adCmdFile			= 256	' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only. 
Const adCmdTableDirect	= 512	' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect. Cannot be combined with the ExecuteOptionEnum value adAsyncExecute.  
Const	ForReading = 1, ForWriting = 2, ForAppending = 8

	dim	strFileName
	dim	strSubmit
	dim	strTemp
	dim	sqlStr
	dim	lngRow
	dim	lngCnt
	dim	blnCheck
	dim	objExcelApp
	dim	objExcelSheet
	dim	objRow
	dim	objCell
	dim	objFS
	dim	objTF
	dim	objBasp
	dim	db
	dim	rsList
	dim	dbName
	dim	strPn
	dim	strQty
	dim	strSyushi
	dim	strRECORD
	dim	strArr
	dim	rsZaiko
	dim	i
	dim	qstrYM
	dim	qstrCenterCd
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function formMsg(msg) {
		fileForm.submit.value = "msg";
	}
-->
</SCRIPT>
<%
Sub OnTransactionAbort()
	msgbox("OnTransactionAbort()")
	objExcelApp = nothing
end Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML LANG="ja">
<head>
<meta http-equiv="Pragma" content="no-cache">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
<LINK REL=STYLESHEET TYPE="text/css" HREF="ir.css" TITLE="CSS">
<title>��Ǝ��ԃf�[�^�C���|�[�g</title>
<!-- jdMenu head�p include �J�n -->
<link href="jquery.jdMenu.css" rel="stylesheet" type="text/css" />
<script src="jquery.js" type="text/javascript"></script>
<script src="jquery.dimensions.js" type="text/javascript"></script>
<script src="jquery.positionBy.js" type="text/javascript"></script>
<script src="jquery.bgiframe.js" type="text/javascript"></script>
<script src="jquery.jdMenu.js" type="text/javascript"></script>
<!-- jdMenu head�p include �I�� -->
</HEAD>
<body>
<!-- jdMenu body�p include �J�n -->
<!--#include file="jdmenu-sdc-ir.asp" -->
<!-- jdMenu body�p include �I�� -->
<%
	dim	strReq
	dim	objBasp21
	dim	nTotalBytes
	dim	bBinaryRead
	dim	strYM
	dim	strCenterCD
	dim	strFileSize
	dim	retCode
	dim	strLog
	dim	strMsg
	dim	strStdout

	Response.Flush
	strMsg = ""
	strLog = ""
	nTotalBytes	= Request.TotalBytes
	Response.Write "Request.TotalBytes=" & nTotalBytes & "<br>"
	Response.Flush
	if nTotalBytes then
		bBinaryRead	= Request.BinaryRead(nTotalBytes)
		set objBasp21	= Server.Createobject("basp21")
'		strYM	= objBasp21.Form(bBinaryRead,"YM")
'		Response.Write "�o�^�N��=" & strYM & "<br>"
		strCenterCD	= UCASE(objBasp21.Form(bBinaryRead,"CenterCD"))
		Response.Write "�Z���^�[=" & strCenterCD & "<br>"
		strFileName	= objBasp21.FormFileName(bBinaryRead,"fName")
		Response.Write "�t�@�C����=" & strFileName & "<br>"
		strFileSize	= objBasp21.FormFileSize(bBinaryRead,"fName")
		Response.Write "FileSize=" & strFileSize & "<br>"
		Response.Flush
		'���������`�F�b�N
		strMsg = CheckOption(strCenterCD,strFileName)
		if strMsg = "" then
			dim	strSaveFileName
			' �o�c�����t�@�C��(Excel)���A�b�v���[�h
'			strFileSize = objBasp21.FormSaveAs(bBinaryRead,"fName",Server.MapPath(strFileName))
			strSaveFileName = "�o�c����_" & strCenterCD & ".xls"
			strSaveFileName = Server.MapPath(strSaveFileName)
			strFileSize = objBasp21.FormSaveAs(bBinaryRead,"fName",strSaveFileName)
'			Response.Write "FileSize=" & strFileSize & "(" & Server.MapPath(strFileName) & ")<br>"
			Response.Write "FileSize=" & strFileSize & "(" & strSaveFileName & ")<br>"
			' �o�c�����t�@�C��(Excel)��Ǎ����o�^
			retCode = LoadIrDataXls(strCenterCD,strSaveFileName,objBasp21,strLog)
			if retCode = 0 then
				strMsg = "�o�^���܂����B(" & retCode & ")"
				strLog = LoadTextFile("loadir01_" & strCenterCD & ".log")
			else
				strMsg = "�o�^�G���[�B(" & retCode & ")"
			end if
		end if
		set objBasp21 = Nothing
	end if
%>
	<%=strMsg%>
	<FORM>
	<INPUT type="button" value="�߂�" onClick="history.back()">
	</FORM>
	<hr><div id="info"><Pre>
	<%=strLog%>
	</Pre></div><hr>
</body>
</html>
<%
'--------------------------------------------------------------------
' �I�v�V�����`�F�b�N
'--------------------------------------------------------------------
Function CheckOption(byVal strCenterCD,byVal strFileName)
	dim	strMsg

	strMsg = ""
	if strCenterCD = "" then
		strMsg = "<font color=red>�Z���^�[</font>���w�肵�ĉ�����"
	elseif strFileName = "" then
		strMsg = "<font color=red>�t�@�C��</font>���w�肵�ĉ�����"
	end if
	CheckOption = strMsg
End Function

'--------------------------------------------------------------------
' �e�L�X�g�t�@�C���Ǎ�
'--------------------------------------------------------------------
Function LoadTextFile(byVal strFileName)
	dim	objFSO
	dim	objFile
	dim	strBuff

	set objFSO = CreateObject("Scripting.FileSystemObject")
	strFileName = Server.MapPath(strFileName)
	Set objFile = objFSO.OpenTextFile(strFilename, ForReading, False)
	strBuff = objFile.ReadAll()
	objFile.Close
	Set objFile	= Nothing
	Set objFSO	= Nothing
	LoadTextFile = strBuff
End Function
'---------------------------------------------------------------------
'Wscript.Shell Run���\�b�h�̈���
'---------------------------------------------------------------------
'WshShell.Run ��Q����
Const vbHide = 0             '�E�B���h�E���\��
Const vbNormalFocus = 1      '�ʏ�̃E�B���h�E�A���őO�ʂ̃E�B���h�E
Const vbMinimizedFocus = 2   '�ŏ����A���őO�ʂ̃E�B���h�E
Const vbMaximizedFocus = 3   '�ő剻�A���őO�ʂ̃E�B���h�E
Const vbNormalNoFocus = 4    '�ʏ�̃E�B���h�E�A�������A�őO�ʂɂ͂Ȃ�Ȃ�
Const vbMinimizedNoFocus = 6 '�ŏ����A�������A�őO�ʂɂ͂Ȃ�Ȃ�
'WshShell.Run ��R����
'True	���s�����v���O�������I������܂ŁA�X�N���v�g�̏�����ҋ@
'False	�X�N���v�g�̏����𑱍s
'--------------------------------------------------------------------
' �o�c�����t�@�C��(Excel)��Ǎ����o�^
'--------------------------------------------------------------------
Function LoadIrDataXls(byVal strCenterCD,byVal strFileName,bobj,strStdout)
	Dim	WshShell
	dim	strCommand
	dim	retCode

	strCommand = "cscript.exe "
'	strCommand = strCommand & Server.MapPath("loadir.vbs") & " "
	strCommand = strCommand & Server.MapPath("loadir01.vbs") & " "
'	strCommand = strCommand & """" & Server.MapPath(strFileName) & """"
	strCommand = strCommand & """" & strFileName & """"
	strCommand = strCommand & " " & strCenterCD

	Server.ScriptTimeout = 900
	Response.Write "Command=" & strCommand & "<br>"
	Response.Write "�������ł��B���΂炭���҂����������B<br>"
	Response.Flush

	Set WshShell = CreateObject("Wscript.Shell")
	retCode = WshShell.Run(strCommand, vbHide, True)
	Set WshShell = Nothing

'           0  - �v���O�������N�����ďI����҂����ɖ߂�܂��B
'           1  - �v���O�������N�����ďI������܂ő҂��܂��B
'                �W���o�͂𕶎���Ŏ���܂��B
'           2  - �v���O�������N�����ďI������܂ő҂��܂��B
'                �W���o�͂��o�C�i���Ŏ���܂��B
'           3�ȏ�̒l  - �v���O�����̏I�����w�肵�����ԁimsec)����
'                �҂��܂��B�w�肵�����ԓ��ɏI���ΕW���o�͂�����܂��B
'                ��̏ꍇ�́A�W���o�͂𕶎���ŁA�����̏ꍇ��
'                �W���o�͂��o�C�i���Ŏ���܂��B
'                �^�C���A�E�g���w�肵�����Ƃ��͂�������g���܂��B
'	retCode = bobj.Execute(strCommand,1,strStdout)
'	strStdout = bobj.Execute2(strCommand,1)
	LoadIrDataXls = retCode
End Function

'--------------------------------------------------------------------
'http://sites.google.com/site/kikineria/scripting/fileupload-asp
'--------------------------------------------------------------------
Function UploadFile()
    Dim ObjStream
    Dim nByte,P_Data,Target,B_Data
    Dim i
    Dim S_LOC,E_LOC,FN_SEARCH,SAVE_FN
    dim strSaveFName

    strSaveFName = ""
    If Request.TotalBytes then
        Set ObjStream = Server.CreateObject("ADODB.Stream")
            ObjStream.Open
			'***�X�g���[���^�C�v���o�C�i���֐ݒ�
            ObjStream.Type = 1
			'***�|�X�g�f�[�^�̑��o�C�g��
            nByte = Request.TotalBytes
			'***�|�X�g�f�[�^�擾
            P_Data = Request.BinaryRead(nByte)
			'***�t�@�C������������̐ݒ�
            FN_SEARCH = "filename="
			'***�t�@�C�����̊J�n�ʒu���擾
            S_LOC = InStrB(P_Data, STR2BIN(FN_SEARCH)) + Len(FN_SEARCH)
			'***�t�@�C�����̏I���ʒu���擾
            E_LOC = InStrB(S_LOC,P_Data, STR2BIN(Chr(13))) - 1

            For i = S_LOC + 1 TO E_LOC - 1
                If (&h81 <= AscB(MidB(P_Data,i,1)) And _
                    AscB(MidB(P_Data,i,1)) <= &h9F) Or _
                    (&hE0 <= AscB(MidB(P_Data,i,1)) And _
                    AscB(MidB(P_Data,i,1)) <= &hEF) Then
                    SAVE_FN = SAVE_FN & Chr(AscB(MidB(P_Data, i, 1)) * 256 + _
                    AscB(MidB(P_Data, i + 1, 1)))
                    i = i + 1
                Else
                    SAVE_FN = SAVE_FN & Chr(AscB(MidB(P_Data,i,1)))
                End If
            Next

			'***�t�@�C�����̎擾
            SAVE_FN = Right(SAVE_FN,Len(SAVE_FN) - InstrRev(SAVE_FN,"\"))
			'***�f�[�^�J�n�ʒu�̎擾
            S_LOC = InStrB(P_Data, STR2BIN(Chr(13)))
			'***�f�[�^������̎擾
            Target = LeftB(P_Data, S_LOC)

            For i = 0 TO 2
                S_LOC = InStrB(S_LOC+1, P_Data, STR2BIN(Chr(13)))
            Next

			'***�f�[�^�I���ʒu�̎擾
            E_LOC = InStrB(S_LOC+1, P_Data, Target)
			'***�o�C�i���f�[�^���X�g���[���֏�������
            ObjStream.Write P_Data
			'***�J�n�ʒu���w��
            ObjStream.Position = S_LOC + 1
			'***�X�g���[���̓ǂݎ��
            B_Data = ObjStream.Read(E_LOC - (S_LOC + 2) - 2)
			'***���݂̈ʒu��������
            ObjStream.Position = 0
			'***���݂̈ʒu���X�g���[���̏I�[�ɐݒ�
            ObjStream.SetEOS
			'***�o�C�i���f�[�^���X�g���[���֏�������
            ObjStream.Write B_Data
			'***�X�g���[�����t�@�C���֕ۑ�(�㏑��)
'			strSaveFName = "zaikoglics.time"
			strSaveFName = SAVE_FN
            ObjStream.SaveToFile Server.MapPath(strSaveFName),2
'		    Response.Write "Upload file:" & Server.MapPath(strSaveFName)

            ObjStream.Close
	        Set ObjStream = Nothing
    End If
    UploadFile = strSaveFName
End Function

'***��������o�C�i���ϊ�����֐�
Function STR2BIN(strData)
    Dim i
    For i = 1 To Len(strData)
        STR2BIN = STR2BIN & ChrB(AscB(Mid(strData, i, 1)))
    Next
End Function

'--------------------------------------------------------------------
'POST�f�[�^�̎��
'--------------------------------------------------------------------
Function GetRequest(byVal strName)
	dim	strReq
	dim	objBasp21
	dim	nTotalBytes
	dim	bBinaryRead

	strReq = ""
	nTotalBytes	= Request.TotalBytes
	if nTotalBytes then
		bBinaryRead	= Request.BinaryRead(nTotalBytes)
		set objBasp21	= Server.Createobject("basp21")
		strReq	= objBasp21.Form(bBinaryRead,strName)
		set objBasp21 = Nothing
	end if
	GetRequest = strReq
End Function
'--------------------------------------------------------------------
%>
