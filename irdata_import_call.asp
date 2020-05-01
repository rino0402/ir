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
<title>作業時間データインポート</title>
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
'		Response.Write "登録年月=" & strYM & "<br>"
		strCenterCD	= UCASE(objBasp21.Form(bBinaryRead,"CenterCD"))
		Response.Write "センター=" & strCenterCD & "<br>"
		strFileName	= objBasp21.FormFileName(bBinaryRead,"fName")
		Response.Write "ファイル名=" & strFileName & "<br>"
		strFileSize	= objBasp21.FormFileSize(bBinaryRead,"fName")
		Response.Write "FileSize=" & strFileSize & "<br>"
		Response.Flush
		'処理条件チェック
		strMsg = CheckOption(strCenterCD,strFileName)
		if strMsg = "" then
			dim	strSaveFileName
			' 経営資料ファイル(Excel)をアップロード
'			strFileSize = objBasp21.FormSaveAs(bBinaryRead,"fName",Server.MapPath(strFileName))
			strSaveFileName = "経営資料_" & strCenterCD & ".xls"
			strSaveFileName = Server.MapPath(strSaveFileName)
			strFileSize = objBasp21.FormSaveAs(bBinaryRead,"fName",strSaveFileName)
'			Response.Write "FileSize=" & strFileSize & "(" & Server.MapPath(strFileName) & ")<br>"
			Response.Write "FileSize=" & strFileSize & "(" & strSaveFileName & ")<br>"
			' 経営資料ファイル(Excel)を読込＆登録
			retCode = LoadIrDataXls(strCenterCD,strSaveFileName,objBasp21,strLog)
			if retCode = 0 then
				strMsg = "登録しました。(" & retCode & ")"
				strLog = LoadTextFile("loadir01_" & strCenterCD & ".log")
			else
				strMsg = "登録エラー。(" & retCode & ")"
			end if
		end if
		set objBasp21 = Nothing
	end if
%>
	<%=strMsg%>
	<FORM>
	<INPUT type="button" value="戻る" onClick="history.back()">
	</FORM>
	<hr><div id="info"><Pre>
	<%=strLog%>
	</Pre></div><hr>
</body>
</html>
<%
'--------------------------------------------------------------------
' オプションチェック
'--------------------------------------------------------------------
Function CheckOption(byVal strCenterCD,byVal strFileName)
	dim	strMsg

	strMsg = ""
	if strCenterCD = "" then
		strMsg = "<font color=red>センター</font>を指定して下さい"
	elseif strFileName = "" then
		strMsg = "<font color=red>ファイル</font>を指定して下さい"
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
' 経営資料ファイル(Excel)を読込＆登録
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
	Response.Write "処理中です。しばらくお待ちください。<br>"
	Response.Flush

	Set WshShell = CreateObject("Wscript.Shell")
	retCode = WshShell.Run(strCommand, vbHide, True)
	Set WshShell = Nothing

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
			'***ストリームタイプをバイナリへ設定
            ObjStream.Type = 1
			'***ポストデータの総バイト数
            nByte = Request.TotalBytes
			'***ポストデータ取得
            P_Data = Request.BinaryRead(nByte)
			'***ファイ名検索文字列の設定
            FN_SEARCH = "filename="
			'***ファイル名の開始位置を取得
            S_LOC = InStrB(P_Data, STR2BIN(FN_SEARCH)) + Len(FN_SEARCH)
			'***ファイル名の終了位置を取得
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

			'***ファイル名の取得
            SAVE_FN = Right(SAVE_FN,Len(SAVE_FN) - InstrRev(SAVE_FN,"\"))
			'***データ開始位置の取得
            S_LOC = InStrB(P_Data, STR2BIN(Chr(13)))
			'***データ文字列の取得
            Target = LeftB(P_Data, S_LOC)

            For i = 0 TO 2
                S_LOC = InStrB(S_LOC+1, P_Data, STR2BIN(Chr(13)))
            Next

			'***データ終了位置の取得
            E_LOC = InStrB(S_LOC+1, P_Data, Target)
			'***バイナリデータをストリームへ書き込み
            ObjStream.Write P_Data
			'***開始位置を指定
            ObjStream.Position = S_LOC + 1
			'***ストリームの読み取り
            B_Data = ObjStream.Read(E_LOC - (S_LOC + 2) - 2)
			'***現在の位置を初期化
            ObjStream.Position = 0
			'***現在の位置をストリームの終端に設定
            ObjStream.SetEOS
			'***バイナリデータをストリームへ書き込み
            ObjStream.Write B_Data
			'***ストリームをファイルへ保存(上書き)
'			strSaveFName = "zaikoglics.time"
			strSaveFName = SAVE_FN
            ObjStream.SaveToFile Server.MapPath(strSaveFName),2
'		    Response.Write "Upload file:" & Server.MapPath(strSaveFName)

            ObjStream.Close
	        Set ObjStream = Nothing
    End If
    UploadFile = strSaveFName
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
