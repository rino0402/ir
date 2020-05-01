<%@LANGUAGE=VBScript%>
<%Option Explicit%>
<% Response.Buffer = True %>
<% Response.Expires = -1 %>
<%
'--------------------------------------------------------------------
'�o�[�W�������
'--------------------------------------------------------------------
Function GetVersion()
	GetVersion = "2012.05.19 �V�K�쐬"
	GetVersion = "2012.12.14 2013�N�x�Ή�...��ƒ�..."
	GetVersion = "2013.05.11 �N�x�̏����l�Ɍ��݂̔N���Z�b�g����悤�ɕύX"
	GetVersion = "2017.04.28 �l����(�p�[�g) ��B�񂪒ǉ����ꂽ�`���ɑΉ�"
End Function
'--------------------------------------------------------------------
%>
<%
'--------------------------------------------------------------------
'POST�f�[�^�̎��
'--------------------------------------------------------------------
Function GetRequest(byVal strName)
	dim	strV
	strV = ucase(Request.QueryString(strName))
	if strV = "" then
		strV =  Request.Cookies("SDC_IR")(strName)
	end if
	if strV = "" then
		if strName = "YYYY" then
			strV = Year(Now())
		end if
	end if
	GetRequest = strV
End Function
'--------------------------------------------------------------------
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML LANG="ja">
<head>
<meta http-equiv="Pragma" content="no-cache">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
<LINK REL=STYLESHEET TYPE="text/css" HREF="ir.css" TITLE="CSS">
<title>�v��p��Ǝ��� �o�^(�A�b�v���[�h)</title>
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
<form name="fileForm" METHOD="post" enctype="multipart/form-data" ACTION="plan_import_call.asp">
	<table>
	<caption style="text-align:left;">�v��p��Ǝ��ԃt�@�C��(Excel)�o�^(�A�b�v���[�h)</caption>
	<tr>
		<td align="right">�N�x</td>
		<td align="left">
			<INPUT TYPE="text" NAME="YYYY" id="YYYY" VALUE="<%=GetRequest("YYYY")%>" size="10" class="input" style="text-align:center;" required pattern="^[0-9]+$">
			<span>�o�^���鎖�ƌv��̔N�x�����</span>
		</td>
	</tr>
	<tr>
		<td align="right">�Z���^�[</td>
		<td align="left">
			<div><INPUT class="input" TYPE="text" NAME="CenterCD" id="CenterCD" VALUE="<%=GetRequest("CenterCD")%>" size="2" style="text-align:center;" required>
			B:����PC/D:���ꕨ��/E:����PC/F:�ޗ�/G:���PC/H:�܈�PC/I:�L��</div>
		</td>
	</tr>
	<tr>
		<td align="right">�t�@�C����</td>
		<td><input class="input" type="file" name="fName" size="100" required><br>
		�v��p��Ǝ��ԃt�@�C��(Excel)���w�肵�ĉ������B��F�o�c�T�������ƌv�掞�ԕ��z.xls
		</td>
	</tr>
	<tr>
		<td align="right">����</td>
		<td>
			<!--input class="cssbutton" type="submit" value="���s"		id="btnSubmit"	onclick="return submitNew();"-->
			<input class="cssbutton" type="submit" value="�S�Ď��s"				id="pAll"		onclick="return submitAction(this);">
			<input class="" 		 type="submit" value="1.�A�b�v���[�h�̂�"	id="pFile"		onclick="return submitAction(this);">
			<input class="" 		 type="submit" value="2.��Ǝ��ԓo�^�̂�"	id="pLoad"		onclick="return submitAction(this);">
			<input class="" 		 type="submit" value="3.�z�������̂�"		id="pHaibun"	onclick="return submitAction(this);">
			<INPUT class="cssbutton" TYPE="reset"  value="���Z�b�g"	id="btnReset"	onClick="location.href='<%=Request.ServerVariables("URL")%>';">
		</td>
	</tr>
	</table>
	<div class="info"><%=getVersion()%></div>
</form>
<h3>�A�b�v���[�h����t�@�C���̃T���v���C���[�W</h3>
*.xlsx �͎g�p�ł��܂���B*.xls �ŕۑ������t�@�C�����g�p���ĉ������B
<h4>�l����(�Ј�)</h4>
<div><img src="jinkenhi.jpg" alt="�l����(�Ј�)"><div>
<h4 class="info_new">�l����(�p�[�g)2017�N�x B��ǉ�</h4>
<div><img src="pert2017.jpg" alt="�l����(�p�[�g)B��ǉ�"><div>
<h4>�l����(�p�[�g)</h4>
<div><img src="pert.jpg" alt="�l����(�p�[�g)"><div>
<h4>�o��z���\</h4>
<div><img src="haibun.jpg" alt="�o��z���\"><div>
</body>
</html>
<script type="text/javascript">
    <!--
        function submitNew() {
			var w = 600;
			var h = 480;
			var x = (screen.width - w) / 2;
			var y = (screen.height - h) / 2;
            //�@ blank�E�B���h�E���J���B
            var win = window.open("","myNewWnd","screenX="+x+",screenY="+y+",left="+x+",top="+y+",width="+w+",height="+h+",status=1,scrollbar=yes");
            
            //�A myform��action���������o���ۑ�
            var action = document.fileForm.action;
            //�B �������ꂽ�{�^����btnSubmit�{�^���ł��邱�Ƃ��T�[�o�ɒʒm
            document.fileForm.action = action + "?btnSubmit=dummy";
            //�C ���X�|���X��myNewWnd��ʂ��󂯎��悤�ɐݒ�
            document.fileForm.target = "myNewWnd";
            //�D �t�H�[���̃T�u�~�b�g
            document.fileForm.submit();
            //�E myNewWnd��ʂ��t�H�[�J�X�����Ă�
            win.focus();
            //�F ���̃{�^���ɉe�������Ȃ��悤��target�y��action�̒l�����ɖ߂�
            document.fileForm.target = "_self";
            document.fileForm.action = action;
            return false;
        }
        function submitAction(t) {
			var w = 600;
			var h = 480;
			var x = (screen.width - w) / 2;
			var y = (screen.height - h) / 2;
            //�@ blank�E�B���h�E���J���B
            var win = window.open("","myNewWnd","screenX="+x+",screenY="+y+",left="+x+",top="+y+",width="+w+",height="+h+",status=1,scrollbar=yes");
            
            //�A myform��action���������o���ۑ�
			var action = document.fileForm.action;
//			var action = "plan_import_action.asp";
            //�B �������ꂽ�{�^����btnSubmit�{�^���ł��邱�Ƃ��T�[�o�ɒʒm
            document.fileForm.action = "plan_import_action.asp?btnSubmit="+ t.id;
            //�C ���X�|���X��myNewWnd��ʂ��󂯎��悤�ɐݒ�
            document.fileForm.target = "myNewWnd";
            //�D �t�H�[���̃T�u�~�b�g
            document.fileForm.submit();
            //�E myNewWnd��ʂ��t�H�[�J�X�����Ă�
            win.focus();
            //�F ���̃{�^���ɉe�������Ȃ��悤��target�y��action�̒l�����ɖ߂�
            document.fileForm.target = "_self";
            document.fileForm.action = action;
            return false;
        }
    //-->
</script>

