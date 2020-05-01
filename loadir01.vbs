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
'-----------------------------------------------------------------------
dim	lngRet
lngRet = Main()
WScript.Quit lngRet
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'�g�p���@
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "�o�c����Excel�Ǎ�"
	Wscript.Echo "loadir.vbs [option] <filename.xls> [�Z���^�[]"
	Wscript.Echo " -?"
End Sub
'-----------------------------------------------------------------------
'���C��
'-----------------------------------------------------------------------
Private Function Main()
	dim	i
	dim	strArg
	dim	strFilename
	dim	strCenterCD
	dim	objFSO
	dim	objLog

	strFilename = ""
	strCenterCD = ""
	For Each strArg In WScript.Arguments
    	select case strArg
		case "-?"
			call usage()
			Main = 1
			exit Function
		case else
			if strFilename = "" then
				strFilename = strArg
			elseif strCenterCD = "" then
				strCenterCD = strArg
			else
				usage()
				Main = 1
				exit Function
			end if
		end select
	Next
	if strFilename = "" then
		usage()
		Main = 1
		exit Function
	end if
	set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	set objLog = OpenLogFile2(objFSO,"_" & strCenterCD)
	call LoadExcel(strFilename,objLog,strCenterCD)
	CloseLogFile(objLog)
	set objFSO = Nothing
	Main = 0
End Function

'-----------------------------------------------------------------------
'���O�t�@�C�� Open
'-----------------------------------------------------------------------
Private Function OpenLogFile(byval objFSO)
	dim	objFile
	dim	strFilename

	strFilename = Wscript.ScriptFullName
	strFilename = left(strFilename,len(strFilename)-3)
	strFilename = strFilename & "log"
	Set objFile = objFSO.OpenTextFile(strFilename, ForWriting, True)
	set OpenLogFile = objFile
End Function
'-----------------------------------------------------------------------
'���O�t�@�C�� Write
'-----------------------------------------------------------------------
Private Function WriteLogFile(byval objFile,byval strMsg)
	objFile.WriteLine strMsg
	Wscript.Echo strMsg
End Function
'-----------------------------------------------------------------------
'���O�t�@�C�� Err�\��
'-----------------------------------------------------------------------
Private Function ErrLogFile(byval objFile,byval objErr)
	dim	strMsg
	if objErr.Number <> 0 then
		strMsg = "Error.Number:" & objErr.Number
		Call WriteLogFile(objFile,strMsg)
		strMsg = "Error.Description:" & objErr.Description
		Call WriteLogFile(objFile,strMsg)
	end if
End Function
'-----------------------------------------------------------------------
'���O�t�@�C�� Close
'-----------------------------------------------------------------------
Private Function CloseLogFile(byval objFile)
	objFile.Close
	set CloseLogFile = Nothing
End Function
'-----------------------------------------------------------------------
'��Ǝ���Excel�Ǎ�
'-----------------------------------------------------------------------
Private Sub LoadExcel(byval strFilename,byval objLog,byVal strCenterCD)
	dim	objXL
	dim	objBk
	dim	objSt
	dim	objRg
	dim	lngRow
	dim	lngCol
	dim	rsWTM
'	dim	strYM
'	dim	strCenterCD
	dim	strPersonCD
	dim	strPersonName
	dim	strSyushiCD
	dim	lngWorkTM
	dim	lngMaxRow
	dim	strPassword

'	On Error Resume Next

	Call WriteLogFile(objLog,"LoadExcel(" & strFilename & "," & strCenterCD & ")")

	'-------------------------------------------------------------------
	'Excel�̏���
	'-------------------------------------------------------------------
	Set objXL = WScript.CreateObject("Excel.Application")
	Call ErrLogFile(objLog,Err)
'	objXL.Application.Visible = True
	strPassword = GetPassword(strCenterCD)
	Call WriteLogFile(objLog,"Workbooks.Open(" & strFilename & ")")
	Set objBk = objXL.Workbooks.Open(strFilename,False,True,,strPassword)
	Call ErrLogFile(objLog,Err)

	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̏���
	'-------------------------------------------------------------------
	dim	objDb
	dim	strDbName
	Set objDb = Wscript.CreateObject("ADODB.Connection")
												Call ErrLogFile(objLog,Err)
	strDbName = "IR"
	Call objDb.Open(strDbName)
												Call ErrLogFile(objLog,Err)
	'-------------------------------------------------------------------
	' �e�[�u��Open
	'-------------------------------------------------------------------
	dim	rsIrData
	Set rsIrData = Wscript.CreateObject("ADODB.Recordset")
												Call ErrLogFile(objLog,Err)
	rsIrData.MaxRecords = 1
	rsIrData.Open "IrData", objDb, adOpenKeyset, adLockBatchOptimistic, adCmdTableDirect
												Call ErrLogFile(objLog,Err)
	'-------------------------------------------------------------------
	' �o�c�����f�[�^�Ǎ����e�[�u���o�^
	'-------------------------------------------------------------------
	dim	strStat
	dim	strSheetName
	strStat = ""
	for each objSt in objBk.Worksheets
		strSheetName = objSt.Name
		Call WriteLogFile(objLog,"objSt.Name=" & strSheetName)
		select case strStat
		case ""
			select case strSheetName
			case "P1"
				strStat = strSheetName
			end select
		case "P1"
			select case strSheetName
			case "P2"
				strStat = strSheetName
				Call DaleteIrData(objBk,objDb,objLog,strCenterCD)
			end select
		case "P2"
			select case strSheetName
			case "P3"
				strStat = strSheetName
			case else
				Call LoadIrData(objBk,objSt,objDb,rsIrData,objLog,strCenterCD)
			end select
		case "P3"
			select case strSheetName
			case "P4"
				strStat = strSheetName
			end select
		case "P4"
		end select
	next
	'-------------------------------------------------------------------
	'�f�[�^�x�[�X�̌㏈��
	'-------------------------------------------------------------------
	rsIrData.Close
	set rsIrData = Nothing
	objDb.Close
	set objDb = Nothing
	'-------------------------------------------------------------------
	'Excel�̌㏈��
	'-------------------------------------------------------------------
	call objBk.Close(False)
	set objXL = Nothing
End Sub

Function GetPassword(byval strCenterCD)
	dim	strPassword
	strPassword = ""
	select case strCenterCD
	case "D"
		strPassword = "lioncat1962"
	case "E"
		strPassword = "SHIGAPC"
	case "H"
		strPassword = "fk2011"
	case else
	end select
	GetPassword = strPassword
End Function

Function LoadIrData(byval objBk,byval objSt,byval objDb,byval rsIrData,byval objLog,byval strCenterCD)
	Call WriteLogFile(objLog,"LoadIrData(" & objSt.Name & ")")
	dim	strCampYear
	strCampYear = objSt.Range("B2")
	Call WriteLogFile(objLog,"CampYear=" & strCampYear)

	dim	strYM
    dim	strSyushiCD

'	if strCenterCD <> "Z" then
	if strCenterCD <> "T" then
		strCenterCD = GetBookInfo(objBk,"CenterCD")
	end if
	strSyushiCD = objSt.Name

	dim	objSt000	' ���v
	set objSt000 = objBk.Worksheets("000")
	dim	i
	for i = 0 to 11
		strYM = GetBookInfo(objBk,"YM_" & (i + 1))
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"A0100",objSt000,objSt,i,10)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"A0200",objSt000,objSt,i,11)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"A0300",objSt000,objSt,i,12)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"A0400",objSt000,objSt,i,13)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"A0500",objSt000,objSt,i,14)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"A0600",objSt000,objSt,i,15)

		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"B0100",objSt000,objSt,i,17)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"B0200",objSt000,objSt,i,18)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"B0300",objSt000,objSt,i,19)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"B0400",objSt000,objSt,i,20)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"B0500",objSt000,objSt,i,21)

		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"B0600",objSt000,objSt,i,23)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"B0700",objSt000,objSt,i,24)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"B0800",objSt000,objSt,i,25)

		' �����v�� ����/�Ԑڐl����
		Call objSt.Unprotect("sdc2035")
'2018.11.30		objSt.Range("EA94").Offset(0,i) = GetCostPlan(objLog,objDb,strYM,strCenterCD,strSyushiCD,"X0100")
'2018.11.30		objSt.Range("EA95").Offset(0,i) = GetCostPlan(objLog,objDb,strYM,strCenterCD,strSyushiCD,"X0200")
		if objSt.Range("EA28").Offset(0,i) = 0 then
			'�l����Z�b�g����Ă��Ȃ��ꍇ
			objSt.Range("EA28").Offset(0,i) = objSt.Range("EA94").Offset(0,i) + objSt.Range("EA95").Offset(0,i)
		else
			'�l����Z�b�g����Ă���ꍇ
			if objSt.Range("EA94").Offset(0,i) = 0 then
				'���ڐl����=0�̏ꍇ
				'�Ԑڐl����=�l����
				objSt.Range("EA95").Offset(0,i) = objSt.Range("EA28").Offset(0,i) + objSt.Range("EA29").Offset(0,i)
			elseif objSt.Range("EA94").Offset(0,i) = 0 then
				'�Ԑڐl����=0�̏ꍇ
			end if
		end if
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"C0100",objSt000,objSt,i,28)	'�l����
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"C0200",objSt000,objSt,i,29)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"C0300",objSt000,objSt,i,30)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"C0400",objSt000,objSt,i,31)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"C0500",objSt000,objSt,i,32)
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"C0600",objSt000,objSt,i,33)

		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"D0100",objSt000,objSt,i,36)

		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"E9900",objSt000,objSt,i,36)	'���ʊǗ���

		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"X0100",objSt000,objSt,i, 94)	'���ڐl����
		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"X0200",objSt000,objSt,i, 95)	'�Ԑڐl����
'		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"Y0100",objSt000,objSt,i, 96)	'���ڍ��(��)
'		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"Y0200",objSt000,objSt,i, 97)	'�Ԑڍ��(��)
'		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"Y0010",objSt000,objSt,i, 98)	'�Ζ�����(��)
'		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"Y0300",objSt000,objSt,i, 99)	'����(��)
'		Call AddIdData(objLog,objDb,rsIrData,strYM,strCenterCD,strSyushiCD,"Y0400",objSt000,objSt,i,100)	'�L������(��)

	next

	LoadIrData = 0
End Function

Function GetCostPlan(objLog,objDb _
				  ,byval strYM _
				  ,byval strCenterCD _
				  ,byval strSyushiCD _
				  ,byval strKamokuCD _
					)
	strYM = CInt(Left(strYM,4)) + 1 & Right(strYM,2)
	select case strKamokuCD
	case "X0100"
		strKamokuCD = "'Z%1'"
	case "X0200"
		strKamokuCD = "'Z%2'"
	end select
	dim	strSql
	strSql = "select Sum(Plan) s" _
		   & " from Attendance" _
		   & " where DT = '" & strYM & "'" _
		   & " and  CenterCD = '" & strCenterCD & "'" _
		   & " and  SyushiCD = '" & strSyushiCD & "'" _
		   & " and  KamokuCD like " & strKamokuCD & "" _
		   & ""
	dim	rsAtt
	Call WriteLogFile(objLog,"GetCostPlan():Execute:" & strSql)
	set rsAtt = objDb.Execute(strSql)
	dim curCostPlan
	curCostPlan = 0
	if rsAtt.Eof = False then
		curCostPlan = rsAtt.Fields("s")
	end if
	set rsAtt = Nothing
	Call WriteLogFile(objLog,"GetCostPlan()=" & curCostPlan)
	GetCostPlan = curCostPlan
End Function
'--------------------------------------------------------------------
'�O�N���ђl��Ԃ�
'--------------------------------------------------------------------
Function GetCostLast(objLog _
				  ,objDb _
				  ,byval strYM _
				  ,byval strCenterCD _
				  ,byval strSyushiCD _
				  ,byval strKamokuCD _
					)
	dim	strSql
	dim	strWhere
	strWhere = makeWhere(strWhere,"YM"		,CLng(strYM)-100	,"")
	strWhere = makeWhere(strWhere,"CenterCD",strCenterCD		,"")
	strWhere = makeWhere(strWhere,"SyushiCD",strSyushiCD		,"")
	strWhere = makeWhere(strWhere,"KamokuCD",strKamokuCD		,"")
	strSql = "select"
	strSql = strSql & " *"
	strSql = strSql & " from IrData "
	strSql = strSql & strWhere
	dim	objRs
	Call WriteLogFile(objLog,"GetCostLast():Execute:" & strSql)
	set objRs = objDb.Execute(strSql)
	dim curCostLast
	curCostLast = 0
	if objRs.Eof = False then
		curCostLast = objRs.Fields("Result")
	end if
	set objRs = Nothing
	Call WriteLogFile(objLog,"GetCostLast()=" & curCostLast)
	GetCostLast = curCostLast
End Function

'--------------------------------------------------------------------
'�����v�� ��Ǝ��Ԃ�Ԃ�
'--------------------------------------------------------------------
Function GetCostAtt(objLog _
				  ,objDb _
				  ,byval strYM _
				  ,byval strCenterCD _
				  ,byval strSyushiCD _
				  ,byval strKamokuCD _
					)
	dim	strSql

	dim	strWhere
	strWhere = makeWhere(strWhere,"DT"		,strYM			,"")
	strWhere = makeWhere(strWhere,"CenterCD",strCenterCD	,"")
	strWhere = makeWhere(strWhere,"SyushiCD",strSyushiCD	,"")
	select case strKamokuCD
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

	dim	objRs
	Call WriteLogFile(objLog,"GetCostAtt():Execute:" & strSql)
	set objRs = objDb.Execute(strSql)
	dim curCostAtt
	curCostAtt = 0
	if objRs.Eof = False then
		curCostAtt = objRs.Fields("Plan")
	end if
	set objRs = Nothing
	Call WriteLogFile(objLog,"GetCostAtt()=" & curCostAtt)
	GetCostAtt = curCostAtt
End Function


Function AddIdData(objLog _
				  ,objDb _
				  ,rsIrData _
				  ,byval strYM _
				  ,byval strCenterCD _
				  ,byval strSyushiCD _
				  ,byval strKamokuCD _
				  ,objSt000 _
				  ,objSt _
				  ,byval intCol _
				  ,byval lngRow)
	dim	lngPlan
	dim	lngResult
	dim	lngPrev

	dim	i
	for i = 1 to 2
		if i = 1 then
			'�����̓o�^
			lngPlan		= GetValue(objSt.Range("CA" & lngRow).Offset(0,intCol))	'�����v��
			lngResult	= GetValue(objSt.Range("AA" & lngRow).Offset(0,intCol))	'��������
			lngPrev		= GetValue(objSt.Range("DA" & lngRow).Offset(0,intCol))	'�O������
			select case strKamokuCD
			case "X0100","X0200"
                if GetValue(objSt.Range("AA90").Offset(0,intCol)) = 0 then
                    lngResult = 0
                end if
				lngPlan	= GetCostAtt(objLog,objDb,strYM,strCenterCD,strSyushiCD,strKamokuCD)	'�����v�� ��Ǝ���
				lngPrev		= GetCostLast(objLog,objDb,strYM,strCenterCD,strSyushiCD,strKamokuCD)	'�O������
			end select
		else
			'�����̓o�^
			strYM = CInt(Left(strYM,4)) + 1 & Right(strYM,2)
			' �����v��
			lngPlan		= GetValue(objSt.Range("EA" & lngRow).Offset(0,intCol))
			lngResult	= 0
			' �O�N����
			if objSt000.Range("AA16").Offset(0,intCol) = 0 then
				' �����̔��オ0�̏ꍇ�́A�v����Z�b�g
				lngPrev		= GetValue(objSt.Range("CA" & lngRow).Offset(0,intCol))
			else
				' �����̑O�N���тɍ������т�o�^
				lngPrev		= GetValue(objSt.Range("AA" & lngRow).Offset(0,intCol))
			end if
		end if
		select case strKamokuCD
		case "B0500","B0800"
			lngPlan		= lngPlan * -1
			lngResult	= lngResult * -1
			lngPrev		= lngPrev * -1
		end select
	
		rsIrData.AddNew
								Call ErrLogFile(objLog,Err)
		rsIrData.Fields("YM")		= strYM
		rsIrData.Fields("CenterCD")	= strCenterCD
		rsIrData.Fields("SyushiCD")	= strSyushiCD
		rsIrData.Fields("KamokuCD")	= strKamokuCD
		rsIrData.Fields("Plan")		= lngPlan
		rsIrData.Fields("Result")	= lngResult
		rsIrData.Fields("Prev")		= lngPrev
		rsIrData.UpdateBatch
								Call ErrLogFile(objLog,Err)
	next
End Function

Function GetValue(byval r)
	dim	v
	v = 0
	if isempty(r) = False then
		if isnumeric(r) then
			v = r
		end if
	end if
	GetValue = v
End Function

Function DaleteIrData(objBk,objDb,objLog,byval strCenterCD)
	dim	strYM_S
	dim	strYM_E

'	if strCenterCD <> "Z" then
	if strCenterCD <> "T" then
		strCenterCD = GetBookInfo(objBk,"CenterCD")
	end if

	strYM_S = GetBookInfo(objBk,"YM_1")
	strYM_E = GetBookInfo(objBk,"YM_12")
	strYM_E = CInt(Left(strYM_E,4)) + 1 & Right(strYM_E,2)
	call DateleIdDataSub(objLog,objDb,strYM_S,strYM_E,strCenterCD)
End Function

Function GetBookInfo(byval objBk,byval strInfo)
	dim	strValue
	dim	lngValue
	dim	stFile
	set stFile = objBk.Worksheets("FILE")
	strValue = ""
	select case strInfo
	case "CenterCD"
		strValue = left(stFile.Range("K1"),1)
	case "YM_1","YM_2","YM_3","YM_4","YM_5","YM_6","YM_7","YM_8","YM_9"
		strValue = GetYM(stFile.Range("C2"),right(strInfo,1))
	case "YM_10","YM_11","YM_12"
		strValue = GetYM(stFile.Range("C2"),right(strInfo,2))
	end select
	GetBookInfo = strValue
End Function

Function GetYM(byval lngCampYear _
			  ,byval lngCampMonth _
			  )
	dim	lngYM
	lngYM = lngCampYear + 196600
	if lngCampMonth < 10 then
		lngYM = (lngCampYear + 1966) * 100 + lngCampMonth + 3
	else
		lngYM = (lngCampYear + 1967) * 100 + lngCampMonth - 9
	end if
	GetYM = lngYM & ""
End Function

Function DateleIdDataSub(byval objLog,byval objDb,byval strYM_S,byval strYM_E,byval strCenterCD)
	dim	strSql
	strSql = "delete " _
		   & " from IrData" _
		   & " where YM between '" & strYM_S & "' and '" & strYM_E & "'" _
		   & " and  CenterCD = '" & strCenterCD & "'" _
		   & ""
	Call WriteLogFile(objLog,"DateleIdData():" & strSql)
	call objDb.Execute(strSql)
	DateleIdDataSub = 0
End Function
