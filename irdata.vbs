Option Explicit
'-----------------------------------------------------------------------
'���C���ďo���C���N���[�h
'-----------------------------------------------------------------------
Function Include( strFileName )
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file
	dim	strScriptPath
	strScriptPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	strFileName = strScriptPath & strFileName
    Set file = fso.OpenTextFile( strFileName )
    ExecuteGlobal file.ReadAll
    file.Close
	Set fso = Nothing
End Function
Call Include("const.vbs")

dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "�o�c�����f�[�^"
	Wscript.Echo "irdata.vbs [option]"
	Wscript.Echo "option:/ym:<�����N��>    �@�@�����N��...201304"
	Wscript.Echo "       /center:<�Z���^�[>    �Z���^�[...D"
	Wscript.Echo "       /syushi:<���x>    �@�@���x�@�@...110"
	Wscript.Echo "       /prev             �@�@�O�N����"
	Wscript.Echo "       /plan             �@�@�����v��"
	Wscript.Echo "       /update           �@�@�f�[�^���X�V"
	Wscript.Echo "       /debug           �@�@ �f�o�b�O���b�Z�[�W�\��"
End Sub
'-----------------------------------------------------------------------
'���C������
'-----------------------------------------------------------------------
Private Function Main()

	'���O�����I�v�V�����`�F�b�N
	select case WScript.Arguments.UnNamed.Count
	case 3
	case else
	end select
	'���O�t���I�v�V�����`�F�b�N
	dim	strArg
	for each strArg in WScript.Arguments.Named
		select case lcase(strArg)
		case "update"
		case "debug"
		case "ym"
		case "center"
		case "syushi"
		case "prev"
		case "plan"
		case "?"
			call usage()
			exit function
		case else
			call usage()
			exit function
		end select
	next
	select case GetFunction()
	case "prev"
		Call IrDataPrev()
	case "plan"
		Call IrDataPlan()
	end select
	Main = 0
End Function

Function GetFunction()
	GetFunction = "prev"
	if WScript.Arguments.Named.Exists("plan") then
		GetFunction = "plan"
	end if
End Function

Function IrDataPlan()
	Wscript.Echo "IrDataPlan()"
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	dim	strDbName
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	strDbName = "IR"
	Call objDb.Open(strDbName)
	'-------------------------------------------------------------------
	' �e�[�u��Open
	'-------------------------------------------------------------------
	dim	rsIrData
	Set rsIrData = Wscript.CreateObject("ADODB.Recordset")
	rsIrData.Open GetSqlIr(), objDb, adOpenKeyset, adLockOptimistic
	dim	lngCnt
	lngCnt = 0
	do while rsIrData.Eof = False
		dim	strMsg
		strMsg = makeMsg(GetFieldValue(rsIrData,"YM")		,-7) _
			   & makeMsg(GetFieldValue(rsIrData,"CenterCD")	,-7) _
			   & makeMsg(GetFieldValue(rsIrData,"SyushiCD")	,-7) _
			   & makeMsg(GetFieldValue(rsIrData,"KamokuCD")	,-7) _
			   & makeMsg(GetFieldValue(rsIrData,"Plan")		,10) _
			   & ""
		Wscript.Echo strMsg
		dim	rsAtt
		set rsAtt = ExecuteAdodb(objDb,GetSqlAtt(rsIrData))
		strMsg = makeMsg(GetFieldValue(rsAtt,"DT")			,-7) _
			   & makeMsg(GetFieldValue(rsAtt,"CenterCD")	,-7) _
			   & makeMsg(GetFieldValue(rsAtt,"SyushiCD")	,-7) _
			   & makeMsg(""	,-7) _
			   & makeMsg(GetFieldValue(rsAtt,"Plan")	,10) _
			   & ""
		dim	strDiff
		strDiff = ""
		if rsAtt.Eof = False then
			if GetFieldValue(rsIrData,"Plan") <> GetFieldValue(rsAtt,"Plan") then
				strDiff = "�~"
				lngCnt = lngCnt + 1
			end if
		end if
		if strDiff <> "" then
			strMsg = strMsg & " " & strDiff
			if WScript.Arguments.Named.Exists("update") then
				rsIrData.Fields("Plan") = GetFieldValue(rsAtt,"Plan")
				rsIrData.UpdateBatch
				strMsg = strMsg & " �X�V"
			end if
		end if
		Wscript.Echo strMsg
		rsIrData.MoveNext
	loop

	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̌㏈��
	'-------------------------------------------------------------------
	rsIrData.Close
	set rsIrData = Nothing
	objDb.Close
	set objDb = Nothing
	Wscript.Echo "�Č���=" & lngCnt
End Function

Function IrDataPrev()
	Wscript.Echo "IrDataPrev()"
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	dim	strDbName
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	strDbName = "IR"
	Call objDb.Open(strDbName)
	'-------------------------------------------------------------------
	' �e�[�u��Open
	'-------------------------------------------------------------------
	dim	rsIrData
	Set rsIrData = Wscript.CreateObject("ADODB.Recordset")
'	Set rsIrData = objDb.Execute("select * from IrData where YM='201304' and CenterCD='D' and KamokuCD like 'X%' order by YM,CenterCD,SyushiCD,KamokuCD")
'	Set rsIrData = objDb.Open(GetSqlIrPlan())
	rsIrData.Open GetSqlIr(), objDb, adOpenKeyset, adLockOptimistic
	dim	lngCnt
	lngCnt = 0
	do while rsIrData.Eof = False
		dim	strMsg
		strMsg = makeMsg(GetFieldValue(rsIrData,"YM")		,-7) _
			   & makeMsg(GetFieldValue(rsIrData,"CenterCD")	,-7) _
			   & makeMsg(GetFieldValue(rsIrData,"SyushiCD")	,-7) _
			   & makeMsg(GetFieldValue(rsIrData,"KamokuCD")	,-7) _
			   & makeMsg(GetFieldValue(rsIrData,"Prev")		,10) _
			   & ""
		Wscript.Echo strMsg
		dim	rsIrPrev
		set rsIrPrev = ExecuteAdodb(objDb,GetSqlIrLast(rsIrData))
		strMsg = makeMsg(GetFieldValue(rsIrPrev,"YM")		,-7) _
			   & makeMsg(GetFieldValue(rsIrPrev,"CenterCD")	,-7) _
			   & makeMsg(GetFieldValue(rsIrPrev,"SyushiCD")	,-7) _
			   & makeMsg(GetFieldValue(rsIrPrev,"KamokuCD")	,-7) _
			   & makeMsg(GetFieldValue(rsIrPrev,"Result")	,10) _
			   & ""
		dim	strDiff
		strDiff = ""
		if GetFieldValue(rsIrData,"Prev") <> GetFieldValue(rsIrPrev,"Result") then
			strDiff = "�~"
			lngCnt = lngCnt + 1
		end if
		if strDiff <> "" then
			strMsg = strMsg & " " & strDiff
			if WScript.Arguments.Named.Exists("update") then
				rsIrData.Fields("Prev") = GetFieldValue(rsIrPrev,"Result")
				rsIrData.UpdateBatch
				strMsg = strMsg & " �X�V"
			end if
		end if
		Wscript.Echo strMsg
		rsIrData.MoveNext
	loop

	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̌㏈��
	'-------------------------------------------------------------------
	rsIrData.Close
	set rsIrData = Nothing
	objDb.Close
	set objDb = Nothing
	Wscript.Echo "�Č���=" & lngCnt
End Function

Function GetSqlIrLast(rsIrData)
	dim	strSql

	dim	strWhere
	strWhere = makeWhere(strWhere,"YM"		,CLng(GetFieldValue(rsIrData,"YM"))-100,"")
	strWhere = makeWhere(strWhere,"CenterCD",GetFieldValue(rsIrData,"CenterCD"),"")
	strWhere = makeWhere(strWhere,"SyushiCD",GetFieldValue(rsIrData,"SyushiCD"),"")
	strWhere = makeWhere(strWhere,"KamokuCD",GetFieldValue(rsIrData,"KamokuCD"),"")
	strSql = "select"
	strSql = strSql & " *"
	strSql = strSql & " from IrData "
	strSql = strSql & strWhere
	strSql = strSql & " order by"
	strSql = strSql & " YM"
	strSql = strSql & ",CenterCD"
	strSql = strSql & ",SyushiCD"
	strSql = strSql & ",KamokuCD"
	GetSqlIrLast = strSql
End Function

Function GetSqlIr()
	dim	strSql

	dim	strWhere
	strWhere = makeWhere(strWhere,"YM"		,GetOption("ym",""),"")
	strWhere = makeWhere(strWhere,"CenterCD",GetOption("center",""),"")
	strWhere = makeWhere(strWhere,"SyushiCD",GetOption("syushi",""),"")
	strWhere = makeWhere(strWhere,"KamokuCD","X%","")
	strSql = "select"
	strSql = strSql & " *"
	strSql = strSql & " from IrData "
	strSql = strSql & strWhere
	strSql = strSql & " order by"
	strSql = strSql & " YM"
	strSql = strSql & ",CenterCD"
	strSql = strSql & ",SyushiCD"
	strSql = strSql & ",KamokuCD"
	GetSqlIr = strSql
End Function


Function IrData()
	Wscript.Echo "IrData()"
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	dim	strDbName
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	strDbName = "IR"
	Call objDb.Open(strDbName)
	'-------------------------------------------------------------------
	' �e�[�u��Open
	'-------------------------------------------------------------------
	dim	rsIrData
'	Set rsIrData = Wscript.CreateObject("ADODB.Recordset")
'	Set rsIrData = objDb.Execute("select * from IrData where YM='201304' and CenterCD='D' and KamokuCD like 'X%' order by YM,CenterCD,SyushiCD,KamokuCD")
	Set rsIrData = objDb.Execute(GetSqlIr())
	do while rsIrData.Eof = False
		Wscript.Echo GetFieldValue(rsIrData,"YM") 			& " " 	_
				   & GetFieldValue(rsIrData,"CenterCD")	& " "	_
				   & GetFieldValue(rsIrData,"SyushiCD")	& " "	_
				   & GetFieldValue(rsIrData,"C0000")	& " "	_
				   & GetFieldValue(rsIrData,"Plan")		& " "	_
				   & GetFieldValue(rsIrData,"X0100")	& " "	_
				   & GetFieldValue(rsIrData,"X0200")	& " "	_
				   & ""
		dim	rsAtt
		Set rsAtt = objDb.Execute(GetSqlAtt(rsIrData))
		Wscript.Echo GetFieldValue(rsAtt,"DT") 		& " " 	_
				   & GetFieldValue(rsAtt,"CenterCD")	& " "	_
				   & GetFieldValue(rsAtt,"SyushiCD")	& " "	_
				   & "     "	& " "	_
				   & GetFieldValue(rsAtt,"Plan")		& " "	_
				   & ""
		rsIrData.MoveNext
	loop

	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̌㏈��
	'-------------------------------------------------------------------
	rsIrData.Close
	set rsIrData = Nothing
	objDb.Close
	set objDb = Nothing

End Function

Function GetSqlIrxxx()
	dim	strSql

	dim	strWhere
	strWhere = makeWhere(strWhere,"YM"		,"201304","")
	strWhere = makeWhere(strWhere,"CenterCD","D","")
'	strWhere = makeWhere(strWhere,"KamokuCD","C%","")
	strSql = "select"
	strSql = strSql & " YM"
	strSql = strSql & ",CenterCD"
	strSql = strSql & ",SyushiCD"
	strSql = strSql & ",'C0000' C0000"
	strSql = strSql & ",sum(if(KamokuCD like 'C%',Plan,0)) Plan"
	strSql = strSql & ",sum(if(KamokuCD='X0100',Plan,0)) X0100"
	strSql = strSql & ",sum(if(KamokuCD='X0200',Plan,0)) X0200"
	strSql = strSql & " from IrData "
	strSql = strSql & strWhere
	strSql = strSql & " group by"
	strSql = strSql & " YM"
	strSql = strSql & ",CenterCD"
	strSql = strSql & ",SyushiCD"
	strSql = strSql & ",C0000"
	GetSqlIr = strSql
End Function


Function GetSqlAtt(rsIrData)
	dim	strSql

	dim	strWhere
	strWhere = makeWhere(strWhere,"DT"		,GetFieldValue(rsIrData,"YM"),"")
	strWhere = makeWhere(strWhere,"CenterCD",GetFieldValue(rsIrData,"CenterCD"),"")
	strWhere = makeWhere(strWhere,"SyushiCD",GetFieldValue(rsIrData,"SyushiCD"),"")
	select case GetFieldValue(rsIrData,"KamokuCD")
	case "C0000"
		strWhere = strWhere & " and (KamokuCD like 'ZZ%01' or KamokuCD like 'ZZ%02')"
	case "X0100"
		strWhere = makeWhere(strWhere,"KamokuCD","ZZ%01","")
	case "X0200"
		strWhere = makeWhere(strWhere,"KamokuCD","ZZ%02","")
	end select
	strSql = "select"
	strSql = strSql & " DT"
	strSql = strSql & ",CenterCD"
	strSql = strSql & ",SyushiCD"
	strSql = strSql & ",sum(Plan) Plan"
	strSql = strSql & " from Attendance "
	strSql = strSql & strWhere
	strSql = strSql & " group by"
	strSql = strSql & " DT"
	strSql = strSql & ",CenterCD"
	strSql = strSql & ",SyushiCD"
	GetSqlAtt = strSql
End Function


