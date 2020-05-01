<%@LANGUAGE=VBScript%>
<%Option Explicit%>
<% Response.Buffer = True %>
<% Response.Expires = -1 %>
<%
	dim	strVersion
	strVersion = "2012.05.07 IEで実行すると「無効なPathの文字です」になる不具合に対応"
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
	dim	strDenDate
	dim	chkInsert
	dim	strWsName
	dim	wsNameParts
	dim	wsNameChoha
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

	dbName = "newsdc"

	qstrCenterCd	= GetRequest("CenterCd")
	if len(qstrCenterCd) = 0 then
	end if

	strSubmit = GetRequest("submit")

	if len(strSubmit) > 0 then
		Response.Cookies("SDC_IR")("CenterCd") = qstrCenterCd
	else
		if len(qstrCenterCd) = 0 then
			qstrCenterCd = Request.Cookies("SDC_IR")("CenterCd")
		end if
	end if
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

	Sub DispMsg(byval strMsg)
		Response.Write("<SCRIPT LANGUAGE='JavaScript'><!-- " & now() & vbcrlf)
		if strMsg = "" then
			Response.Write("fileForm.disabled = false;" & vbcrlf)
			Response.Write("fileForm.submit.value ='OK';" & vbcrlf)
		else
			Response.Write("fileForm.disabled = true;" & vbcrlf)
			Response.Write("fileForm.submit.value ='" & strMsg & "';" & vbcrlf)
		end if
		Response.Write(" --></SCRIPT>" & vbcrlf)
		Response.Flush()
	End sub
	Function getStrDate(byval dt)
		dim m
		dim d
		m = month(dt)
		m = right("00",2 - len(m)) & m
		d = day(dt)
		d = right("00",2 - len(d)) & d
		getStrDate = year(dt) & m & d
	End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML LANG="ja">
<head>
<meta http-equiv="Pragma" content="no-cache">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
<LINK REL=STYLESHEET TYPE="text/css" HREF="ir.css" TITLE="CSS">
<title>経営資料 登録(アップロード)</title>
<!-- jdMenu head用 include 開始 -->
<link href="jquery.jdMenu.css" rel="stylesheet" type="text/css" />
<script src="jquery.js" type="text/javascript"></script>
<script src="jquery.dimensions.js" type="text/javascript"></script>
<script src="jquery.positionBy.js" type="text/javascript"></script>
<script src="jquery.bgiframe.js" type="text/javascript"></script>
<script src="jquery.jdMenu.js" type="text/javascript"></script>
<!-- jdMenu head用 include 終了 -->
</HEAD>
<body>
<!-- jdMenu body用 include 開始 -->
<!--#include file="jdmenu-sdc-ir.asp" -->
<!-- jdMenu body用 include 終了 -->

<form name="fileForm" METHOD="post" enctype="multipart/form-data" ACTION="irdata_import_call.asp">
	<table>
	<caption style="text-align:left;">経営資料(Excel)登録(アップロード)</caption>
	<tr>
		<td align="right">センター</td>
		<td align="left">
			<div><INPUT TYPE="text" NAME="CenterCD" id="CenterCD" VALUE="<%=qstrCenterCD%>" size="2" style="text-align:center;">
			B:小野PC/D:滋賀物流/E:滋賀PC/F:奈良/G:大阪PC/H:袋井PC/I:広島</div>
		</td>
	</tr>
	<tr>
		<td align="right">ファイル名</td>
		<td><input type="file" name="fName" size="100"><br>
		経営資料(Excel)ファイルを指定して下さい。例：System経営資料.xls
		</td>
	</tr>
	<tr>
		<td></td>
		<td>
			<input type="submit"  name="submit" value="実行">
			<INPUT TYPE="reset" value="リセット" id=reset1 name=reset1 onClick="location.href='<%=Request.ServerVariables("URL")%>';">
		</td>
	</tr>
	</table>
	<div class="info"><%=strVersion%></div>
</form>
<body>
</body>
</html>
<%
'--------------------------------------------------------------------
'POSTデータの受取
'--------------------------------------------------------------------
Function GetRequest(byVal strName)
	GetRequest = ucase(Request.QueryString(strName))
End Function
'--------------------------------------------------------------------
%>
