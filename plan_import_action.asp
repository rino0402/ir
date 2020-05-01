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
	<title>計画作業時間ファイル(Excel)登録(アップロード) 処理中</title>
	<STYLE type="text/css">
	<!--
	.log {
		overflow: scroll;   /* スクロール表示 */
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
			<INPUT type="button" value="閉じる" onClick="window.close();" disable>
		</FORM>
		<table>
		<tr>
		<td align="right">処理タイプ：</td><td><%=GetActionType()%></td>
		</tr>
		<tr>
		<td align="right">ファイルアップロード：</td><td><progress id="pFile" value="0" max="100"><span id="pFileMsg"></span></progress></td>
		</tr>
		<tr>
		<td align="right">計画用作業時間登録：</td><td><progress id="pLoad" value="0" max="13"><span id="pLoadMsg"></span></progress></td>
		</tr>
		<tr>
		<td align="right">配分処理：</td><td><progress id="pHaibun" value="0" max="13"><span id="pHaibunMsg"></span></progress></td>
		</tr>
		<!--tr>
		<td align="right">処理状況：</td><td><progress id="pAction" value=null	max="100"><span id="pActionMsg"></span></progress></td>
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
' プログレスバー更新
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
' ログにメッセージ表示
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
' 年月を返す
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
' メイン処理
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

'	Server.ScriptTimeout = 60 * 30	' 30分
'	Server.ScriptTimeout = 0		'タイムアウトなし
	Server.ScriptTimeout = 60 * 60	' 60分
'	Response.Flush
	strMsg = ""
	strLog = ""
	nTotalBytes	= Request.TotalBytes
	Log "Request.TotalBytes=" & nTotalBytes & "<br>"
'	Response.Flush
	if nTotalBytes then
		'----------------------------------------------------
		'ファイルアップロード
		'----------------------------------------------------
		bBinaryRead	= Request.BinaryRead(nTotalBytes)
		set objBasp21	= Server.Createobject("basp21")
		' 年度
		dim	strYYYY
		strYYYY	= objBasp21.Form(bBinaryRead,"YYYY")
		Log "登録年度=" & strYYYY & "<br>"
		' センター
		strCenterCD	= UCASE(objBasp21.Form(bBinaryRead,"CenterCD"))
		Log "センター=" & strCenterCD & "<br>"
		' ファイル名
		dim	strFileName
		strFileName	= objBasp21.FormFileName(bBinaryRead,"fName")
		Log "ファイル名=" & strFileName & "<br>"
		strFileSize	= objBasp21.FormFileSize(bBinaryRead,"fName")
		Log "FileSize=" & strFileSize & "<br>"
'		Response.Flush
		'処理条件チェック
		strMsg = CheckOption(strCenterCD,strYYYY,strFileName)
		Log strMsg
		if strMsg = "" then
			dim	strSaveFileName
			' 経営資料ファイル(Excel)をアップロード
'			strFileSize = objBasp21.FormSaveAs(bBinaryRead,"fName",Server.MapPath(strFileName))
			strSaveFileName = "計画用作業時間_" & strCenterCD & "_" & strYYYY & ".xls"
			strSaveFileName = Server.MapPath(strSaveFileName)
			if CheckAction("pFile") then
				Call SetProgress("pFile","null","アップロード中...")
				Log "アップロード開始<br>"
				strFileSize = objBasp21.FormSaveAs(bBinaryRead,"fName",strSaveFileName)
				Log "FileSize=" & strFileSize & "(" & strSaveFileName & ")<br>"
				Call SetProgress("pFile",100,"完了")
				Log "アップロード終了<br>"
			end if
			'----------------------------------------------------------------
			' 今期計画作業時間ファイル(Excel)を読込＆登録
			'----------------------------------------------------------------
			if CheckAction("pLoad") then
				Call SetProgress("pLoad","null","計画用作業時間を登録中...")
				Log "登録処理開始<br>"
				dim	i
				for i = 1 to 12
					Call SetProgress("pLoad",i,"計画用作業時間を登録中..." & i & "/12")
					retCode = LoadIrDataXls(strCenterCD,GetYYYYMM(strYYYY,i),strSaveFileName,objBasp21,strLog)
					if retCode <> 0 then
						exit for
					end if
				next
				if retCode = 0 then
					strMsg = "正常終了(" & retCode & ")"
					strLog = LoadTextFile("loadplan03_" & strCenterCD & ".log")
				else
					strMsg = "登録エラー(" & retCode & ")"
				end if
				Log strMsg & "<br>"
				Call SetProgress("pLoad",13,strMsg)
				Log "登録処理終了<br><hr>"
			end if
			'--------------------------------------------------------------------
			' 配分処理
			'--------------------------------------------------------------------
			if CheckAction("pHaibun") _
			or CheckAction("pLoad") then
				Call SetProgress("pHaibun","null","配分処理 実行中...")
				Log "配分処理開始<br>"
				for i = 1 to 12
					Call SetProgress("pHaibun",i,"配分処理 実行中..." & i & "/12")
					retCode = CallHaibun(strCenterCD,GetYYYYMM(strYYYY,i),objBasp21,strLog)
					if retCode <> 0 then
						exit for
					end if
				next
				if retCode = 0 then
					strMsg = "正常終了(" & retCode & ")"
				else
					strMsg = "登録エラー(" & retCode & ")"
				end if
				Log strMsg & "<br>"
				Call SetProgress("pHaibun",13,strMsg)
				Log "配分処理終了:" & retCode & "<br>"
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
' 処理チェック
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
' オプションチェック
'--------------------------------------------------------------------
Function CheckOption(byVal strCenterCD,byVal strYYYY,byVal strFileName)
	dim	strMsg

	strMsg = ""
	if strCenterCD = "" then
		strMsg = "<font color=red>センター</font>を指定して下さい"
	elseif strYYYY = "" then
		strMsg = "<font color=red>年度</font>を指定して下さい"
	elseif strFileName = "" then
		select case GetActionType()
		case "pAll","pFile"
			strMsg = "<font color=red>ファイル</font>を指定して下さい"
		end select
	end if
	CheckOption = strMsg
End Function

'--------------------------------------------------------------------
' テキストファイル読込
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
'Wscript.Shell Runメソッドの引数
'---------------------------------------------------------------------
'WshShell.Run 第２引数
Const vbHide = 0             'ウィンドウを非表示
Const vbNormalFocus = 1      '通常のウィンドウ、かつ最前面のウィンドウ
Const vbMinimizedFocus = 2   '最小化、かつ最前面のウィンドウ
Const vbMaximizedFocus = 3   '最大化、かつ最前面のウィンドウ
Const vbNormalNoFocus = 4    '通常のウィンドウ、ただし、最前面にはならない
Const vbMinimizedNoFocus = 6 '最小化、ただし、最前面にはならない
'WshShell.Run 第３引数
'True	実行したプログラムが終了するまで、スクリプトの処理を待機
'False	スクリプトの処理を続行
'--------------------------------------------------------------------
' 今期計画作業時間ファイル(Excel)を読込＆登録
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
'           0  - プログラムを起動して終了を待たずに戻ります。
'           1  - プログラムを起動して終了するまで待ちます。
'                標準出力を文字列で受取ります。
'           2  - プログラムを起動して終了するまで待ちます。
'                標準出力をバイナリで受取ります。
'           3以上の値  - プログラムの終了を指定した時間（msec)だけ
'                待ちます。指定した時間内に終われば標準出力を受取ります。
'                奇数の場合は、標準出力を文字列で、偶数の場合は
'                標準出力をバイナリで受取ります。
'                タイムアウトを指定したいときはこちらを使います。
'	retCode = bobj.Execute(strCommand,1,strStdout)
'	strStdout = bobj.Execute2(strCommand,1)
'--------------------------------------------------------------------
' 今期計画作業時間ファイル(Excel)を読込＆登録
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

'***文字列をバイナリ変換する関数
Function STR2BIN(strData)
    Dim i
    For i = 1 To Len(strData)
        STR2BIN = STR2BIN & ChrB(AscB(Mid(strData, i, 1)))
    Next
End Function

'--------------------------------------------------------------------
'POSTデータの受取
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
