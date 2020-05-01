'Option Explicit
'-------------------------------
'const.vbs
'\\hs1\it\ir �p
'-------------------------------
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- CommandTypeEnum Values ----
Const adCmdUnspecified	= -1	' Unspecified type of command 
Const adCmdText			= 1		' Evaluates CommandText as a textual definition of a command or stored procedure call 
Const adCmdTable		= 2		' Evaluates CommandText as a table name whose columns are returned by an SQL query 
Const adCmdStoredProc	= 4		' Evaluates CommandText as a stored procedure name 
Const adCmdUnknown		= 8		' Default. Unknown type of command 
Const adCmdFile			= 256	' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only. 
Const adCmdTableDirect	= 512	' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect. Cannot be combined with the ExecuteOptionEnum value adAsyncExecute.  

Const ForReading		= 1
Const ForWriting		= 2
Const ForAppending		= 8
Const adSearchForward	= 1
' ObjectStateEnum
' �I�u�W�F�N�g���J���Ă��邩���Ă��邩�A�f�[�^ �\�[�X�ɐڑ������A
' �R�}���h�����s�����A�܂��̓f�[�^���擾�����ǂ�����\���܂��B
Const	adStateClosed		= 0 ' �I�u�W�F�N�g�����Ă��邱�Ƃ������܂��B 
Const	adStateOpen			= 1 ' �I�u�W�F�N�g���J���Ă��邱�Ƃ������܂��B 
Const	adStateConnecting	= 2 ' �I�u�W�F�N�g���ڑ����Ă��邱�Ƃ������܂��B 
Const	adStateExecuting	= 4 ' �I�u�W�F�N�g���R�}���h�����s���ł��邱�Ƃ������܂��B 
Const	adStateFetching		= 8 ' �I�u�W�F�N�g�̍s���擾����Ă��邱�Ƃ������܂��B 

const	xlToLeft = -4159
const	xlUp	 = -4162

Function makeMsg(byval sVal,byval iLen)
	sVal = RTrim(sVal)
	if iLen > 0 then
		sVal = Right(space(iLen) & sVal,iLen)
	else
		iLen = iLen * -1
		sVal = Left(sVal & space(iLen),iLen)
	end if
	makeMsg = sVal
End Function

function GetDateTime(dt)
	dim	tmpYYYYMMDD
	dim	tmpHHMMSS
	'/// �N���� �쐬
	tmpYYYYMMDD = year(dt) & Right(00 & month(dt), 2) & Right(00 & day(dt), 2)
	'/// ���� �쐬   
	tmpHHMMSS   = Right(00 & hour(dt), 2) & Right(00 & minute(dt), 2) & Right(00 & second(dt), 2)
	'/// ����   
	GetDateTime = tmpYYYYMMDD & tmpHHMMSS
end function

Sub DispMsg(strMsg)
	Wscript.Echo strMsg
End Sub
'-----------------------------------------------------------------------
'���O�t�@�C�� Open
'-----------------------------------------------------------------------
Function OpenLogFile(byval objFSO)
	dim	objFile
	dim	strFilename

	strFilename = Wscript.ScriptFullName
	strFilename = left(strFilename,len(strFilename)-3)
	strFilename = strFilename & "log"

	Debug "OpenLogFile():" & strFilename

	On Error Resume Next
		Set objFile = objFSO.OpenTextFile(strFilename, ForWriting, True)
		if Err.Number <> 0 then
			DispMsg " OpenTextFile() Error:" & Hex(Err.Number) & " " & Err.Description
		end if
	On Error Goto 0
	set OpenLogFile = objFile
End Function
'-----------------------------------------------------------------------
'���O�t�@�C�� Open
'-----------------------------------------------------------------------
Function OpenLogFile2(byval objFSO,byval strAdd)
	dim	objFile
	dim	strFilename

	strFilename = Wscript.ScriptFullName
	strFilename = left(strFilename,len(strFilename)-4)
	strFilename = strFilename & strAdd & ".log"

	Debug "OpenLogFile2():" & strFilename

	On Error Resume Next
		Set objFile = objFSO.OpenTextFile(strFilename, ForWriting, True)
		if Err.Number <> 0 then
			DispMsg " OpenLogFile2() Error:" & Hex(Err.Number) & " " & Err.Description
		end if
	On Error Goto 0
	set OpenLogFile2 = objFile
End Function
'-----------------------------------------------------------------------
'���O�t�@�C�� Write
'-----------------------------------------------------------------------
Private Function WriteLogFile(byval objFile,byval strMsg)
	if objFile is Nothing then
	else
		objFile.WriteLine strMsg
	end if
	Wscript.Echo strMsg
End Function
'-----------------------------------------------------------------------
'���O�t�@�C�� Err�\��
'-----------------------------------------------------------------------
Private Function ErrLogFile(byval objFile,byval objErr)
	dim	strMsg
	dim	intErrNumber
	intErrNumber = objErr.Number
	if intErrNumber <> 0 then
		strMsg = "Error.Number:0x" & Hex(intErrNumber)
		Call WriteLogFile(objFile,strMsg)
		strMsg = "Error.Description:" & objErr.Description
		Call WriteLogFile(objFile,strMsg)
	end if
	ErrLogFile = intErrNumber
'	if objErr.Number <> 0 then
'		strMsg = "Error.Number:" & objErr.Number
'		Call WriteLogFile(objFile,strMsg)
'		strMsg = "Error.Description:" & objErr.Description
'		Call WriteLogFile(objFile,strMsg)
'	end if
End Function
'-----------------------------------------------------------------------
'���O�t�@�C�� Close
'-----------------------------------------------------------------------
Private Function CloseLogFile(byval objFile)
	objFile.Close
	set CloseLogFile = Nothing
End Function

'-----------------------------------------------------------------------
'�f�[�^�x�[�X�I�[�v��
'-----------------------------------------------------------------------
Function OpenAdodb(byval strDbName)
	dim	objDb
	Set objDb = Wscript.CreateObject("ADODB.Connection")
	strDbName = "IR"
	Call objDb.Open(strDbName)
	Set OpenAdodb = objDb
End Function
'-----------------------------------------------------------------------
'�f�[�^�x�[�X�N���[�Y
'-----------------------------------------------------------------------
Function CloseAdodb(byval objDb)
	objDb.Close
	set CloseAdodb = Nothing
End Function
'-----------------------------------------------------------------------
'���R�[�h�Z�b�g�I�[�v��
'-----------------------------------------------------------------------
Function OpenRs(byval objDb,byval strTableName)
	dim	objRs
	Set objRs = Wscript.CreateObject("ADODB.Recordset")
	if strTableName <> "" then
		objRs.Open strTableName, objDb, adOpenKeyset, adLockOptimistic,adCmdTableDirect
	end if
	set OpenRs = objRs
End Function
Function UpdateOpenRs(byval objDb,byval objRs,byval strSql)
	if objRs.State <> adStateClosed then
		objRs.Close
	end if
	objRs.Open strSql, objDb, adOpenKeyset, adLockOptimistic
	UpdateOpenRs = objRs.EOF
End Function
'-----------------------------------------------------------------------
'���R�[�h�Z�b�gExecute
'-----------------------------------------------------------------------
Function ExecuteAdodb(byval objDb,byval strSql)
	set ExecuteAdodb = objDb.Execute(strSql)
End Function
'-----------------------------------------------------------------------
'���R�[�h�Z�b�g�N���[�Y
'-----------------------------------------------------------------------
Function CloseRs(byval objRs)
	if objRs.State <> adStateClosed then
		objRs.Close
	end if
	set CloseRs = Nothing
End Function
'-----------------------------------------------------------------------
'�t�B�[���h�l
'-----------------------------------------------------------------------
Function GetFieldValue(byval objRs _
					  ,byval strName _
					  )
	dim	v
	Debug "GetFieldValue(" & strName & "):Type=" & objRs.Fields(strName).Type
	On Error Resume Next
		v = objRs.Fields(strName)
		if Err.Number <> 0 then
			strMsg = strMsg & " GetFieldValue() Error:" & Hex(Err.Number) & " " & Err.Description
		end if
	On Error Goto 0
	
	select case objRs.Fields(strName).Type
	case 6
		if isnull(v) then
			v = 0
		end if
		if v = "" then
			v = 0
		end if
	case else
		if isnull(v) then
			v = ""
		end if
	end select
	GetFieldValue = Rtrim(v)
End Function
'-----------------------------------------------------------------------
'�O�N��
'-----------------------------------------------------------------------
Function GetPrevYM(byval strYYYYMM _
  			 	)
	dim	strNextYM
	dim	strYYYY
	strYYYY = left(strYYYYMM,4)
	dim	strMM
	strMM = right(strYYYYMM,2)
	strMM = CInt(strMM) - 1
	if CInt(strMM) < 1 then
		strYYYY = CInt(strYYYY) - 1
		strMM	= 12
	end if
	GetPrevYM = strYYYY & right("0" & strMM ,2)
End Function
'-----------------------------------------------------------------------
'���N��
'-----------------------------------------------------------------------
Function GetNextYM(byval strYYYYMM _
  			 	)
	dim	strNextYM
	dim	strYYYY
	strYYYY = left(strYYYYMM,4)
	dim	strMM
	strMM = right(strYYYYMM,2)
	strMM = CInt(strMM) + 1
	if CInt(strMM) > 12 then
		strYYYY = CInt(strYYYY) + 1
		strMM	= 1
	end if
	GetNextYM = strYYYY & right("0" & strMM ,2)
End Function

'-----------------------------------------------------------------------
'�N�x��Ԃ�
'-----------------------------------------------------------------------
Function GetNendo(byval strYYYYMM _
  			 	)
	dim	strNendo
'	Debug "GetNendo(" & strYYYYMM & ")"
	strNendo = left(strYYYYMM,4)
	if CInt(right(strYYYYMM,2)) < 4 then
		strNendo = "" & (CInt(strNendo) - 1)
	end if
'	Debug "GetNendo(" & strYYYYMM & ")=" & strNendo
	GetNendo = strNendo
End Function
'-----------------------------------------------------------------------
'�f�o�b�O���b�Z�[�W
'-----------------------------------------------------------------------
Function Debug(byval strMsg)
	if WScript.Arguments.Named.Exists("debug") then
		Wscript.Echo strMsg
	end if
End Function
Function isDebug()
	isDebug = WScript.Arguments.Named.Exists("debug")
End Function

'-----------------------------------------------------------------------
'�I�v�V�����`�F�b�N
'-----------------------------------------------------------------------
Function GetOption(byval strName _
				  ,byval strDefault _
				  )
	dim	strValue

	if strName = "" then
		strValue = ""
		if strDefault < WScript.Arguments.UnNamed.Count then
			strValue = WScript.Arguments.UnNamed(strDefault)
		end if
	else
		strValue = strDefault
		if WScript.Arguments.Named.Exists(strName) then
			strValue = WScript.Arguments.Named(strName)
		end if
	end if
	GetOption = strValue
End Function
'-----------------------------------------------------------------------
'select Where����
'-----------------------------------------------------------------------
Function makeWhere(byval strWhere _
				  ,byval strField _
				  ,byval strValue1 _
				  ,byval strValue2 _
				  )
	dim	strAnd
	dim	strNot
	dim	strCmp
	
	if len(strValue1) > 0 then
		if len(strWhere) > 0 then
			strAnd = " and "
		else
			strAnd = " where "
		end if
		if len(strValue2) > 0 then
			strCmp = "between"
			strWhere = strWhere & strAnd & " " & strField & " " & strCmp & " '" & strValue1 & "' and '" & strValue2 & "'"
		else
			select case left(strValue1,1)
			case "<"
				strValue1 = right(strValue1,len(strValue1)-1)
				strCmp = "<"
			case else
				strValue1 = "'" & strValue1 & "'"
				if instr(1,strValue1,"%") > 0 then
					strCmp = strNot & "like"
				elseif instr(strValue1,",") > 0 then
					strCmp = strNot & "in "
					strValue1 = "(" & replace(strValue1,",","','") & ")"
				else
					if strNot = "" Then
						strCmp = "="
					else
						strCmp = "<>"
					end if
				end if
			end select
			strWhere = strWhere & strAnd & " " & strField & " " & strCmp & " " & strValue1 & ""
		end if
	end if
	makeWhere = strWhere
End Function

'-----------------------------------------------------------------------
'select Where����
'-----------------------------------------------------------------------
Function GetSheet(byval objBk,byval strSheetName)
	Set GetSheet = Nothing
	dim	objSt
	for each objSt in objBk.Worksheets
		Debug "GetSheet(" & strSheetName & "):" & objSt.Name
		if Trim(objSt.Name) = strSheetName then
			Set GetSheet = objSt
			exit for
		end if
		if Replace(Trim(objSt.Name),".","") = strSheetName then
			Set GetSheet = objSt
			exit for
		end if
	next
End Function

'-----------------------------------------------------------------------
'������ �STrim
'-----------------------------------------------------------------------
Function AllTrim(byval strV)
	dim	strTrim
	strTrim = Trim(strV)
	strTrim = Replace(strTrim," ","")
	strTrim = Replace(strTrim,vbCr,"")
	strTrim = Replace(strTrim,vbLf,"")
	AllTrim = strTrim
End Function
