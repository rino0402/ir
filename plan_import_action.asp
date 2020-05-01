<%@LANGUAGE=VBScript%>
<%Option Explicit%>
<% Response.Buffer = False %>
<% Response.Expires = -1 %>
<%
Const	ForReading = 1, ForWriting = 2, ForAppending = 8
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function formMsg(msg) {
		fileForm.submit.value = "msg";
	}
function getBrowserWidth ( ) {  
    if ( window.innerWidth ) { return window.innerWidth; }  
    else if ( document.documentElement && document.documentElement.clientWidth != 0 ) { return document.documentElement.clientWidth; }  
    else if ( document.body ) { return document.body.clientWidth; }  
    return 0;  
}
-->
</SCRIPT>
<%
Sub OnTransactionAbort()
	msgbox("OnTransactionAbort()")
	set objExcelApp = nothing
end Sub
Function GetActionType()
	dim	strActionType
	strActionType = Request.QueryString("btnSubmit")
	GetActionType = strActionType
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML LANG="ja">
<HEAD>
	<meta http-equiv="Pragma" content="no-cache">
	<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
	<LINK REL=STYLESHEET TYPE="text/css" HREF="ir.css" TITLE="CSS">
	<title>�v���Ǝ��ԃt�@�C��(Excel)�o�^(�A�b�v���[�h) ������</title>
	<STYLE type="text/css">
	<!--
	.log {
		overflow: scroll;   /* �X�N���[���\�� */
/*		width: 500px;	*/
		height: 300px;
		font-family	: monospace;
		FONT-SIZE: x-small;
	}
	-->
	</STYLE>
</HEAD>
<BODY>
	<section>
		<FORM>
			<INPUT type="button" value="����" onClick="window.close();" disable>
		</FORM>
		<table>
		<tr>
		<td align="right">�����^�C�v�F</td><td><%=GetActionType()%></td>
		</tr>
		<tr>
		<td align="right">�t�@�C���A�b�v���[�h�F</td><td><progress id="pFile" value="0" max="100"><span id="pFileMsg"></span></progress></td>
		</tr>
		<tr>
		<td align="right">�v��p��Ǝ��ԓo�^�F</td><td><progress id="pLoad" value="0" max="13"><span id="pLoadMsg"></span></progress></td>
		</tr>
		<tr>
		<td align="right">�z�������F</td><td><progress id="pHaibun" value="0" max="13"><span id="pHaibunMsg"></span></progress></td>
		</tr>
		<!--tr>
		<td align="right">�����󋵁F</td><td><progress id="pAction" value=null	max="100"><span id="pActionMsg"></span></progress></td>
		</tr-->
		</table>
	</section>
	<hr>
	<pre><code><div class="log" id="log"></div></code></pre>
	<script>
		document.getElementById("log").innerWidth = getBrowserWidth() - 10;
//		alert(document.getElementById("log").innerWidth);
	</script>
	<%Call Main()%>
</BODY>
</HTML>
<%
'--------------------------------------------------------------------
' �v���O���X�o�[�X�V
'--------------------------------------------------------------------
Function SetProgress(pId,pValue,pMsg)
%>
	<script>
		document.getElementById("<%=pId%>").value			= "<%=pValue%>";
		document.getElementById("<%=pId%>Msg").innerHTML	= "<%=pMsg%>";
	</script>
<%
End Function
%>
<%
'--------------------------------------------------------------------
' ���O�Ƀ��b�Z�[�W�\��
'--------------------------------------------------------------------
Function Log(byVal pMsg)
	pMsg = Replace(pMsg,"\","\\")
%>
	<script>
		document.getElementById("log").innerHTML	= document.getElementById("log").innerHTML + '<%=pMsg%>';
		document.getElementById("log").scrollTop	= document.getElementById("log").scrollHeight;
	</script>
<%
End Function
%>
<%
'--------------------------------------------------------------------
' �N����Ԃ�
'--------------------------------------------------------------------
Function GetYYYYMM(byVal strYYYY,byVal i)
	dim	strYm
	if i < 10 then
		strYm = strYYYY & right("0" & i + 3,2)
	else
		strYm = CStr(CInt(strYYYY)+1) & right("0" & i - 9,2)
	end if
	GetYYYYMM = strYm
End Function
'--------------------------------------------------------------------
' ���C������
'--------------------------------------------------------------------
Function Main()
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

'	Server.ScriptTimeout = 60 * 30	' 30��
'	Server.ScriptTimeout = 0		'�^�C���A�E�g�Ȃ�
	Server.ScriptTimeout = 60 * 60	' 60��
'	Response.Flush
	strMsg = ""
	strLog = ""
	nTotalBytes	= Request.TotalBytes
	Log "Request.TotalBytes=" & nTotalBytes & "<br>"
'	Response.Flush
	if nTotalBytes then
		'----------------------------------------------------
		'�t�@�C���A�b�v���[�h
		'----------------------------------------------------
		bBinaryRead	= Request.BinaryRead(nTotalBytes)
		set objBasp21	= Server.Createobject("basp21")
		' �N�x
		dim	strYYYY
		strYYYY	= objBasp21.Form(bBinaryRead,"YYYY")
		Log "�o�^�N�x=" & strYYYY & "<br>"
		' �Z���^�[
		strCenterCD	= UCASE(objBasp21.Form(bBinaryRead,"CenterCD"))
		Log "�Z���^�[=" & strCenterCD & "<br>"
		' �t�@�C����
		dim	strFileName
		strFileName	= objBasp21.FormFileName(bBinaryRead,"fName")
		Log "�t�@�C����=" & strFileName & "<br>"
		strFileSize	= objBasp21.FormFileSize(bBinaryRead,"fName")
		Log "FileSize=" & strFileSize & "<br>"
'		Response.Flush
		'���������`�F�b�N
		strMsg = CheckOption(strCenterCD,strYYYY,strFileName)
		Log strMsg
		if strMsg = "" then
			dim	strSaveFileName
			' �o�c�����t�@�C��(Excel)���A�b�v���[�h
'			strFileSize = objBasp21.FormSaveAs(bBinaryRead,"fName",Server.MapPath(strFileName))
			strSaveFileName = "�v��p��Ǝ���_" & strCenterCD & "_" & strYYYY & ".xls"
			strSaveFileName = Server.MapPath(strSaveFileName)
			if CheckAction("pFile") then
				Call SetProgress("pFile","null","�A�b�v���[�h��...")
				Log "�A�b�v���[�h�J�n<br>"
				strFileSize = objBasp21.FormSaveAs(bBinaryRead,"fName",strSaveFileName)
				Log "FileSize=" & strFileSize & "(" & strSaveFileName & ")<br>"
				Call SetProgress("pFile",100,"����")
				Log "�A�b�v���[�h�I��<br>"
			end if
			'----------------------------------------------------------------
			' �����v���Ǝ��ԃt�@�C��(Excel)��Ǎ����o�^
			'----------------------------------------------------------------
			if CheckAction("pLoad") then
				Call SetProgress("pLoad","null","�v��p��Ǝ��Ԃ�o�^��...")
				Log "�o�^�����J�n<br>"
				dim	i
				for i = 1 to 12
					Call SetProgress("pLoad",i,"�v��p��Ǝ��Ԃ�o�^��..." & i & "/12")
					retCode = LoadIrDataXls(strCenterCD,GetYYYYMM(strYYYY,i),strSaveFileName,objBasp21,strLog)
					if retCode <> 0 then
						exit for
					end if
				next
				if retCode = 0 then
					strMsg = "����I��(" & retCode & ")"
					strLog = LoadTextFile("loadplan03_" & strCenterCD & ".log")
				else
					strMsg = "�o�^�G���[(" & retCode & ")"
				end if
				Log strMsg & "<br>"
				Call SetProgress("pLoad",13,strMsg)
				Log "�o�^�����I��<br><hr>"
			end if
			'--------------------------------------------------------------------
			' �z������
			'--------------------------------------------------------------------
			if CheckAction("pHaibun") _
			or CheckAction("pLoad") then
				Call SetProgress("pHaibun","null","�z������ ���s��...")
				Log "�z�������J�n<br>"
				for i = 1 to 12
					Call SetProgress("pHaibun",i,"�z������ ���s��..." & i & "/12")
					retCode = CallHaibun(strCenterCD,GetYYYYMM(strYYYY,i),objBasp21,strLog)
					if retCode <> 0 then
						exit for
					end if
				next
				if retCode = 0 then
					strMsg = "����I��(" & retCode & ")"
				else
					strMsg = "�o�^�G���[(" & retCode & ")"
				end if
				Log strMsg & "<br>"
				Call SetProgress("pHaibun",13,strMsg)
				Log "�z�������I��:" & retCode & "<br>"
	'			Response.Flush
				strLog = strLog & vbCrlf & "<hr>"
				strLog = strLog & vbCrlf & "-- attendance.log --"
				strLog = strLog & vbCrlf & LoadTextFile("attendance.log")
			end if
		end if
		set objBasp21 = Nothing
	end if
End Function
'--------------------------------------------------------------------
' �����`�F�b�N
'--------------------------------------------------------------------
Function CheckAction(byVal strAction)
	dim	bRet
	bRet = False
	if GetActionType() = "pAll" then
		bRet = True
	elseif GetActionType() = strAction then
		bRet = True
	end if
	CheckAction = bRet
End Function
'--------------------------------------------------------------------
' �I�v�V�����`�F�b�N
'--------------------------------------------------------------------
Function CheckOption(byVal strCenterCD,byVal strYYYY,byVal strFileName)
	dim	strMsg

	strMsg = ""
	if strCenterCD = "" then
		strMsg = "<font color=red>�Z���^�[</font>���w�肵�ĉ�����"
	elseif strYYYY = "" then
		strMsg = "<font color=red>�N�x</font>���w�肵�ĉ�����"
	elseif strFileName = "" then
		select case GetActionType()
		case "pAll","pFile"
			strMsg = "<font color=red>�t�@�C��</font>���w�肵�ĉ�����"
		end select
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
	dim	objTextFile

	strBuff = ""
	set objFSO = CreateObject("Scripting.FileSystemObject")
	strFileName = Server.MapPath(strFileName)
	Set objFile = objFso.GetFile(strFilename)
	if objFile.Size > 0 then
		Set objTextFile = objFSO.OpenTextFile(strFilename, ForReading, False)
		strBuff = objTextFile.ReadAll()
		objTextFile.Close
		set objTextFile = Nothing
	end if
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
' �����v���Ǝ��ԃt�@�C��(Excel)��Ǎ����o�^
'--------------------------------------------------------------------
Function CallHaibun(byVal strCenterCD,byVal strYYYYMM,bobj,strStdout)
	Dim	WshShell
	dim	strCommand
	dim	retCode
	dim	strLog

	strLog = "att_" & strYYYYMM & "_" & strCenterCD & ".log"
	strLog = Server.MapPath(strLog)
	strCommand = "cmd /S /C cscript //nologo "
'	strCommand = strCommand & Server.MapPath("attendance.vbs")
	strCommand = strCommand & Server.MapPath("att.vbs")
	strCommand = strCommand & " " & strCenterCD
	strCommand = strCommand & " " & strYYYYMM
	strCommand = strCommand & " > " & strLog

	Log strCommand & "<br>"

	Set WshShell = CreateObject("Wscript.Shell")
	retCode = WshShell.Run(strCommand, vbHide, True)
	Set WshShell = Nothing
	CallHaibun = retCode
End Function
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
'--------------------------------------------------------------------
' �����v���Ǝ��ԃt�@�C��(Excel)��Ǎ����o�^
'--------------------------------------------------------------------
Function LoadIrDataXls(byVal strCenterCD,byVal strYYYY,byVal strFileName,bobj,strStdout)
	Dim	WshShell
	dim	strCommand
	dim	retCode
	dim	strLog

	strLog = "plan_" & strYYYY & "_" & strCenterCD & ".log"
	strLog = Server.MapPath(strLog)

	strCommand = "cmd /S /C cscript //nologo "
	select case strCenterCD
	case "G","D"
		strCommand = strCommand & Server.MapPath("loadplan030.vbs")
	case else
		strCommand = strCommand & Server.MapPath("loadplan03.vbs")
	end select
	strCommand = strCommand & " " & strCenterCD
	strCommand = strCommand & " " & strYYYY
	strCommand = strCommand & " """ & strFileName & """"
	strCommand = strCommand & " > " & strLog

	Log strCommand & "<br>"

	Set WshShell = CreateObject("Wscript.Shell")
	retCode = WshShell.Run(strCommand, vbHide, True)
	Set WshShell = Nothing

	LoadIrDataXls = retCode
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
