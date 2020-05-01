Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'Main() メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objAtt
	Set objAtt = New Att
	if objAtt.Init() <> "" then
		call usage()
		exit function
	end if
	call objAtt.Run()
End Function
'-----------------------------------------------------------------------
'usage() 使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "計画勤怠時間 配分処理"
	Wscript.Echo "att.vbs [option] <センター> [yyyymm] [yyyymm]"
	Wscript.Echo " /debug"
	Wscript.Echo " /load	(default)"
	Wscript.Echo " /list"
	Wscript.Echo " /?"
	Wscript.Echo "Ex."
	Wscript.Echo "cscript//nologo att.vbs B 201704"
End Sub
'-----------------------------------------------------------------------
'Class Att
'2017.05.02 新規(処理速度アップ)
'-----------------------------------------------------------------------
Class Att
	Private	strDBName
	Private	objDB
	Private	strCenterCD
	Private	strYM1
	Private	strYM2
	Private	strSql
	Private	strAction
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		strDBName = GetOption("db"	,"ir")
		set objDB = nothing
		strCenterCD = ""
		strYM1		= ""
		strYM2		= ""
		strAction	= ""
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			if strCenterCD = "" then
				strCenterCD = strArg
			elseif strYM1 = "" then
				strYM1 = strArg
			elseif strYM2 = "" then
				strYM2 = strArg
			else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end if
		Next
		if strCenterCD = "" then
			Init = "未指定:センター"
			Disp Init
			Exit Function
		end if
		if strYM1 = "" then
			Init = "未指定:年月"
			Disp Init
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "?"
				Init = strArg
				Exit Function
			case "db"
			case "debug"	' デバッグ
			case "load"
				strAction	= "load"
			case "list"
				strAction	= "list"
			case else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end select
		Next
	End Function
	'-----------------------------------------------------------------------
	'Run() 実行処理
	'-----------------------------------------------------------------------
    Public Function Run()
		Debug ".Run()"
		OpenDb
		Load
		List
		CloseDb
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
    Public Function List()
		if strAction <> "" and strAction <> "list" then
			exit function
		end if
		Debug ".List()"
		AddSql ""
		AddSql "select"
		AddSql " a.DT"
		AddSql ",a.CenterCD"
		AddSql ",a.SyushiCD"
		AddSql ",a.Plan"
		AddSql ",a.Drt"
		AddSql ",a.Ind"
		AddSql ",i.X0100"
		AddSql ",i.X0200"

		AddSql " from ("
		AddSql " select"
		AddSql " DT"
		AddSql ",CenterCD"
		AddSql ",SyushiCD"
		AddSql ",Sum(Plan) Plan"
		AddSql ",Sum(if(Right(RTrim(KamokuCD),1) = '1',Plan,0)) Drt"
		AddSql ",Sum(if(Right(RTrim(KamokuCD),1) = '2',Plan,0)) Ind"
		AddSql " from Attendance"
		AddSql " where CenterCD = '" & strCenterCD & "'"
		AddSql " and KamokuCD like 'ZZ%'"
		if strYM2 = "" then
			AddSql " and DT = '" & strYM1 & "'"
		else
			AddSql " and DT between '" & strYM1 & "' and '" & strYM2 & "'"
		end if
		AddSql " and Plan <> 0"
		AddSql " group by "
		AddSql " DT"
		AddSql ",CenterCD"
		AddSql ",SyushiCD"
		AddSql " ) a"

		AddSql " left outer join ("
		AddSql " select"
		AddSql "  YM DT"
		AddSql " ,CenterCD"
		AddSql " ,SyushiCD"
		AddSql " ,Sum(Plan) Plan"
		AddSql " ,Sum(if(KamokuCD = 'X0100',Plan,0)) X0100"
		AddSql " ,Sum(if(KamokuCD = 'X0200',Plan,0)) X0200"
		AddSql " from IrData"
		AddSql " where KamokuCD in ('X0100','X0200')"
		AddSql " and CenterCD = '" & strCenterCD & "'"
		if strYM2 = "" then
			AddSql " and YM = '" & strYM1 & "'"
		else
			AddSql " and YM between '" & strYM1 & "' and '" & strYM2 & "'"
		end if
		AddSql " group by"
		AddSql "  DT"
		AddSql " ,CenterCD"
		AddSql " ,SyushiCD"
		AddSql " ) i"
		AddSql " on (a.DT = i.DT and a.CenterCD = i.CenterCD and a.SyushiCD = i.SyushiCD)"
		AddSql " order by "
		AddSql " a.DT"
		AddSql ",a.CenterCD"
		AddSql ",a.SyushiCD"

		Wscript.StdOut.Write "List("
		Wscript.StdOut.Write "" & strCenterCD
		Wscript.StdOut.Write " " & strYM1
		Wscript.StdOut.Write " " & strYM2
		Wscript.StdOut.Write ")..."
		on error resume next
		set objRs = objDb.Execute(strSql)
		Wscript.StdOut.WriteLine "0x" & Hex(Err.Number) & ":" & Err.Description
		on error goto 0
		if Err.Number <> 0 then
			Wscript.StdOut.WriteLine
			Wscript.StdOut.WriteLine strSql
			Wscript.Quit
		end if
		do while objRs.Eof = False
			Wscript.StdOut.Write T(objRs.Fields("DT"),-7)
			Wscript.StdOut.Write T(objRs.Fields("CenterCD"),-2)
			Wscript.StdOut.Write T(objRs.Fields("SyushiCD"),-4)
			Wscript.StdOut.Write T(objRs.Fields("Plan"),8)
			Wscript.StdOut.Write T(objRs.Fields("Drt"),8)
			Wscript.StdOut.Write T(objRs.Fields("Ind"),8)
			Wscript.StdOut.Write T(objRs.Fields("X0100"),8)
			Wscript.StdOut.Write T(objRs.Fields("X0200"),8)
			Call SetIrData("Drt","X0100")
			Call SetIrData("Ind","X0200")
			Wscript.StdOut.WriteLine
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'SetIrData() update/insert
	'-----------------------------------------------------------------------
	Private Function SetIrData(byVal strAtt,byVal strIr)
		if Trim(objRs.Fields("SyushiCD")) = "" then
			exit function
		end if
		if Trim(objRs.Fields(strAtt)) = Trim(objRs.Fields(strIr)) then
			exit function
		end if
		Wscript.StdOut.Write " " & strIr
		dim iRet
		iRet = UpdateIrData(strAtt,strIr)
		if iRet > 0 then
			Wscript.StdOut.Write ":u" & iRet
			if iRet > 1 then
				Wscript.StdOut.Write ":異常"
			end if
			exit function
		end if
		iRet = InsertIrData(strAtt,strIr)
		Wscript.StdOut.Write ":i" & iRet
		if iRet <> 1 then
			Wscript.StdOut.Write ":異常"
		end if
	End Function
	'-----------------------------------------------------------------------
	'UpdateIrData() SQL
	'-----------------------------------------------------------------------
	Private Function UpdateIrData(byVal strAtt,byVal strIr)
		AddSql ""
		AddSql "update IrData"
		AddSql " set Plan = " & Trim(objRs.Fields(strAtt))
		AddSql " where YM = '" & Trim(objRs.Fields("DT")) & "'"
		AddSql " and CenterCD = '" & Trim(objRs.Fields("CenterCD")) & "'"
		AddSql " and SyushiCD = '" & Trim(objRs.Fields("SyushiCD")) & "'"
		AddSql " and KamokuCD = '" & strIr & "'"
'		Wscript.StdOut.WriteLine
'		Wscript.StdOut.WriteLine strSql
		objDb.Execute strSql
		UpdateIrData = RowCount()
	End Function
	'-----------------------------------------------------------------------
	'InsertIrData() SQL
	'-----------------------------------------------------------------------
	Private Function InsertIrData(byVal strAtt,byVal strIr)
		AddSql ""
		AddSql "insert IrData"
		AddSql "("
		AddSql " YM"
		AddSql ",CenterCD"
		AddSql ",SyushiCD"
		AddSql ",KamokuCD"
		AddSql ",Plan"
		AddSql ",Result"
		AddSql ",Prev"
		AddSql ") values ("
		AddSql " '" & Trim(objRs.Fields("DT")) & "'"
		AddSql ",'" & Trim(objRs.Fields("CenterCD")) & "'"
		AddSql ",'" & Trim(objRs.Fields("SyushiCD")) & "'"
		AddSql ",'" & Trim(objRs.Fields("KamokuCD")) & "'"
		AddSql "," & Trim(objRs.Fields(strAtt))
		AddSql ",0"
		AddSql ",0"
		AddSql ")"
'		Wscript.StdOut.WriteLine
'		Wscript.StdOut.WriteLine strSql
		objDb.Execute strSql
		UpdateIrData = RowCount()
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
	Private	objRs
    Public Function Load()
		if strAction <> "" and strAction <> "load" then
			exit function
		end if
		Debug ".Load()"
		AddSql ""
		AddSql "update attendance"
		AddSql " set Plan = 0"
		AddSql " where CenterCD = '" & strCenterCD & "'"
		if strYM2 = "" then
			AddSql "   and DT = '" & strYM1 & "'"
		else
			AddSql "   and DT between '" & strYM1 & "' and '" & strYM2 & "'"
		end if
		AddSql " and SyushiCD <> ''"
		AddSql " and Plan <> 0"
		Wscript.StdOut.Write "クリア("
		Wscript.StdOut.Write "" & strCenterCD
		Wscript.StdOut.Write " " & strYM1
		Wscript.StdOut.Write " " & strYM2
		Wscript.StdOut.Write ")..."
		objDb.Execute strSql
		dim	objRow
		set objRow = objDb.Execute("select @@rowcount")
		Wscript.StdOut.WriteLine objRow.Fields(0)
		set	objRow = nothing

		AddSql ""
		AddSql "select"
		AddSql " a.DT"
		AddSql ",a.CenterCD"
		AddSql ",a.HaibunCD"
		AddSql ",a.PersonCD"
		AddSql ",h.SyushiCD"
		AddSql ",a.KamokuCD"
		AddSql ",a.Plan * if(left(a.KamokuCD,2) = 'TM',60,1) Plan"
		AddSql ",Round(a.Plan * if(left(a.KamokuCD,2) = 'TM',60,1) * h.ChokuRatio,0)	PlanDrt"	' direct
		AddSql ",Round(a.Plan * if(left(a.KamokuCD,2) = 'TM',60,1) * h.KanRatio,0)	PlanInd"	' indirect
		AddSql ",h.ChokuRatio"
		AddSql ",h.KanRatio"
		AddSql ",if(h.ChokuRatio>h.KanRatio,h.ChokuRatio,h.KanRatio) MaxRatio"
		AddSql " from attendance a"
		AddSql " left outer join Haibun h"
		AddSql "  on (a.CenterCD = h.CenterCD and a.HaibunCD = h.HaibunCD and GetNendo(a.DT) = h.YM)"
		AddSql " where a.CenterCD = '" & strCenterCD & "'"
		if strYM2 = "" then
			AddSql "   and a.DT = '" & strYM1 & "'"
		else
			AddSql "   and a.DT between '" & strYM1 & "' and '" & strYM2 & "'"
		end if
		AddSql " and a.SyushiCD = ''"
		AddSql " and a.Plan <> 0"
		AddSql " and (h.ChokuRatio <> 0 or h.KanRatio <> 0)"
		AddSql " order by "
		AddSql " a.DT"
		AddSql ",a.CenterCD"
		AddSql ",a.HaibunCD"
		AddSql ",a.PersonCD"
		AddSql ",a.KamokuCD"
		AddSql ",MaxRatio"
		AddSql ",h.SyushiCD"
		Debug ".Load():" & strSql
		Wscript.StdOut.Write "検索中("
		Wscript.StdOut.Write "" & strCenterCD
		Wscript.StdOut.Write " " & strYM1
		Wscript.StdOut.Write " " & strYM2
		Wscript.StdOut.Write ")..."
		set objRs = objDb.Execute(strSql)
		Wscript.StdOut.WriteLine "0x" & Hex(Err.Number) & " " & Err.Description
		dim	cPlan
		cPlan = 0
		Call PrevCheck(0)
		do while True
			if PrevCheck(1) = True then
				if cPlan <> 0 then
					Wscript.StdOut.Write T(prvDT,-7)
					Wscript.StdOut.Write T(prvCenterCD,-2)
					Wscript.StdOut.Write T(prvHaibunCD,-6)
					Wscript.StdOut.Write T(prvPersonCD,-11)
					Wscript.StdOut.Write T(prvKamokuCD,-6)
					Wscript.StdOut.Write T(prvSyushiCD,-4)
					Wscript.StdOut.Write T("",8)
					Wscript.StdOut.Write T(cPlan,8)
					dim	strKamokuCD
					strKamokuCD = Left(prvKamokuCD,4) & "1"
					if prvPlanDrt < prvPlanInd then
						strKamokuCD = Left(prvKamokuCD,4) & "2"
					end if
					Wscript.StdOut.Write ":"
					Wscript.StdOut.Write _
						T(SetPlan(prvDT _
							,	prvCenterCD _
							,	prvHaibunCD _
							,	prvPersonCD _
							,	strKamokuCD _
							,	prvSyushiCD _
							,	cPlan _
							) _
						,-8)
					Wscript.StdOut.WriteLine
				end if
				if objRs.Eof = True then
					exit do
				end if
				cPlan = objRs.Fields("Plan")
			end if
			Wscript.StdOut.Write T(objRs.Fields("DT"),-7)
			Wscript.StdOut.Write T(objRs.Fields("CenterCD"),-2)
			Wscript.StdOut.Write T(objRs.Fields("HaibunCD"),-6)
			Wscript.StdOut.Write T(objRs.Fields("PersonCD"),-11)
			Wscript.StdOut.Write T(objRs.Fields("KamokuCD"),-6)
			Wscript.StdOut.Write T(objRs.Fields("SyushiCD"),-4)
'			Wscript.StdOut.Write T(objRs.Fields("Plan"),8)
			Wscript.StdOut.Write T(cPlan,8)
			Wscript.StdOut.Write T(objRs.Fields("PlanDrt"),8)
			Wscript.StdOut.Write T(objRs.Fields("PlanInd"),8)
'			Wscript.StdOut.Write T(objRs.Fields("ChokuRatio"),8)
'			Wscript.StdOut.Write T(objRs.Fields("KanRatio"),8)
			Wscript.StdOut.Write ":"
			Wscript.StdOut.Write _
				T(SetPlan(GetField("DT") _
					,	GetField("CenterCD") _
					,	GetField("HaibunCD") _
					,	GetField("PersonCD") _
					,	Left(GetField("KamokuCD"),4) & "1" _
					,	GetField("SyushiCD") _
					,	GetField("PlanDrt") _
					) _
				,-8)
			Wscript.StdOut.Write _
				T(SetPlan(GetField("DT") _
					,	GetField("CenterCD") _
					,	GetField("HaibunCD") _
					,	GetField("PersonCD") _
					,	Left(GetField("KamokuCD"),4) & "2" _
					,	GetField("SyushiCD") _
					,	GetField("PlanInd") _
					) _
				,-8)
			Wscript.StdOut.WriteLine
			cPlan = cPlan - objRs.Fields("PlanDrt")
			cPlan = cPlan - objRs.Fields("PlanInd")
			Call PrevCheck(2)
			objRs.MoveNext
		loop
	End Function
	'-----------------------------------------------------------------------
	'GetField()
	'-----------------------------------------------------------------------
	Private Function GetField(byVal strNm)
		GetField = RTrim(objRs.Fields(strNm))
	End Function
	'-----------------------------------------------------------------------
	'SetPlan() update/insert
	'-----------------------------------------------------------------------
	Private Function SetPlan(byVal strDT _
							,byVal strCenterCD _
							,byVal strHaibunCD _
							,byVal strPersonCD _
							,byVal strKamokuCD _
							,byVal strSyushiCD _
							,byVal curPlan _
							)
		if curPlan = 0 then
			SetPlan = ""
			exit function
		end if
		dim iRet
		iRet = UpdatePlan(strDT _
						,strCenterCD _
						,strHaibunCD _
						,strPersonCD _
						,strKamokuCD _
						,strSyushiCD _
						,curPlan _
						)
		if iRet = 1 then
			SetPlan = strKamokuCD & ":u" & iRet
			exit function
		end if
		iRet = InsertPlan(strDT _
						,strCenterCD _
						,strHaibunCD _
						,strPersonCD _
						,strKamokuCD _
						,strSyushiCD _
						,curPlan _
						)
		SetPlan = strKamokuCD & ":i" & iRet
	End Function
	'-----------------------------------------------------------------------
	'UpdatePlan() SQL
	'-----------------------------------------------------------------------
	Private Function UpdatePlan(byVal strDT _
							,byVal strCenterCD _
							,byVal strHaibunCD _
							,byVal strPersonCD _
							,byVal strKamokuCD _
							,byVal strSyushiCD _
							,byVal curPlan _
							)
		AddSql ""
		AddSql "update attendance"
		AddSql " set Plan = Plan + " & curPlan
		AddSql " where DT = '" & strDT & "'"
		AddSql " and CenterCD = '" & strCenterCD & "'"
		AddSql " and HaibunCD = '" & strHaibunCD & "'"
		AddSql " and PersonCD = '" & strPersonCD & "'"
		AddSql " and KamokuCD = '" & strKamokuCD & "'"
		AddSql " and SyushiCD = '" & strSyushiCD & "'"
		Debug "UpdatePlan():" & strSql
		objDb.Execute strSql
			Debug "0x" & Hex(Err.Number) & ":" & Err.Description
		UpdatePlan = RowCount()
	End Function
	'-----------------------------------------------------------------------
	'RowCount() select @@rowcount
	'-----------------------------------------------------------------------
	Private Function RowCount()
		dim objRow
		set objRow = objDb.Execute("select @@rowcount")
		RowCount = objRow.Fields(0)
		set	objRow = nothing
	End Function
	'-----------------------------------------------------------------------
	'InsertPlan() SQL insert
	'-----------------------------------------------------------------------
	Private Function InsertPlan(byVal strDT _
							,byVal strCenterCD _
							,byVal strHaibunCD _
							,byVal strPersonCD _
							,byVal strKamokuCD _
							,byVal strSyushiCD _
							,byVal curPlan _
							)
		AddSql ""
		AddSql "insert into attendance"
		AddSql " ("
		AddSql " DT"
		AddSql ",CenterCD"
		AddSql ",HaibunCD"
		AddSql ",PersonCD"
		AddSql ",KamokuCD"
		AddSql ",SyushiCD"
		AddSql ",Plan"
		AddSql ",Result"
		AddSql ",Prev"
		AddSql " ) values ("
		AddSql " '" & strDT & "'"
		AddSql ",'" & strCenterCD & "'"
		AddSql ",'" & strHaibunCD & "'"
		AddSql ",'" & strPersonCD & "'"
		AddSql ",'" & strKamokuCD & "'"
		AddSql ",'" & strSyushiCD & "'"
		AddSql "," & curPlan
		AddSql ",0"
		AddSql ",0"
		AddSql ")"
		Debug ".InsertPlan():" & strSql
		on error resume next
			objDb.Execute strSql
			Debug "0x" & Hex(Err.Number)
			Debug Err.Description
			InsertPlan = Err.Number
		on error goto 0
		InsertPlan = RowCount()
	End Function
	'-----------------------------------------------------------------------
	'PrevCheck() 前レコードと比較
	'-----------------------------------------------------------------------
	Private	prvDt
	Private	prvCenterCD
	Private	prvHaibunCD
	Private	prvPersonCD
	Private	prvKamokuCD
	Private	prvSyushiCD
	Private	prvPlanDrt
	Private	prvPlanInd
	Private	Function PrevCheck(byVal f)
		PrevCheck = False
		if f = 0 then
			prvDt 		= ""
			prvCenterCD = ""
			prvHaibunCD = ""
			prvPersonCD = ""
			prvKamokuCD = ""
			prvSyushiCD = ""
			prvPlanDrt = 0
			prvPlanInd = 0
			exit function
		end if
		if f = 2 then
			if objRs.Eof = True then
				PrevCheck = PrevCheck(0)
				exit function
			end if
			prvDt 		= objRs.Fields("DT")
			prvCenterCD = objRs.Fields("CenterCD")
			prvHaibunCD = objRs.Fields("HaibunCD")
			prvPersonCD = objRs.Fields("PersonCD")
			prvKamokuCD = objRs.Fields("KamokuCD")
			prvSyushiCD = objRs.Fields("SyushiCD")
			prvPlanDrt = objRs.Fields("PlanDrt")
			prvPlanInd = objRs.Fields("PlanInd")
			exit function
		end if
		if objRs.Eof = True then
			PrevCheck = True
			exit function
		end if
		if prvDt <> objRs.Fields("DT") then
			PrevCheck = True
		end if
		if prvCenterCD <> objRs.Fields("CenterCD") then
			PrevCheck = True
		end if
		if prvHaibunCD <> objRs.Fields("HaibunCD") then
			PrevCheck = True
		end if
		if prvPersonCD <> objRs.Fields("PersonCD") then
			PrevCheck = True
		end if
		if prvKamokuCD <> objRs.Fields("KamokuCD") then
			PrevCheck = True
		end if
	End Function
	'-----------------------------------------------------------------------
	'T() 文字列
	'-----------------------------------------------------------------------
	Private Function T(byVal v,byVal i)
		if i > 0 then
			T = right(space(i) & v,i)
		else
			i = i * -1
			T = LeftB(v & space(i),i)
		end if
	End Function
	'-----------------------------------------------------------------------
	'LeftB() 文字列
	'-----------------------------------------------------------------------
	Private Function LeftB(byVal a_Str, byVal a_int)
		Dim iCount, iAscCode, iLenCount, iLeftStr
		iLenCount = 0
		iLeftStr = ""
		If Len(a_Str) = 0 Then
			LeftB = ""
			Exit Function
		End If
		If a_int = 0 Then
			LeftB = ""
			Exit Function
		End If
		For iCount = 1 to Len(a_Str)
			'** Asc関数で文字コード取得
			iAscCode = Asc(Mid(a_Str, iCount, 1))
			'** 半角は文字コードの長さが2、全角は4(2以上)として判断
			If Len(Hex(iAscCode)) > 2 Then
				iLenCount = iLenCount + 2
			Else
				iLenCount = iLenCount + 1
			End If
			If iLenCount > Cint(a_int) Then
				Exit For
			Else
				iLeftStr = iLeftStr + Mid(a_Str, iCount, 1)
			End If
		Next
		if LenB(iLeftStr) < a_int then
			iLeftStr = iLeftStr & space(a_int - LenB(iLeftStr))
		end if
		LeftB = iLeftStr
	End Function
	'-----------------------------------------------------------------------
	'LenB() 文字列
	'-----------------------------------------------------------------------
	Function LenB(byVal a_Str)
		Dim iCount, iAscCode, iLenCount, iLeftStr
		iLenCount = 0
		iLeftStr = ""
		If Len(a_Str) = 0 Then
			LenB = 0
			Exit Function
		End If
		For iCount = 1 to Len(a_Str)
			'** Asc関数で文字コード取得
			iAscCode = Asc(Mid(a_Str, iCount, 1))
			'** 半角は文字コードの長さが2、全角は4(2以上)として判断
			If Len(Hex(iAscCode)) > 2 Then
				iLenCount = iLenCount + 2
			Else
				iLenCount = iLenCount + 1
			End If
		Next
		LenB = iLenCount
	End Function
	'-----------------------------------------------------------------------
	'GetYM() 年月を返す
	'-----------------------------------------------------------------------
	Private Function GetYM(byVal i)
		dim	intYear
		intYear = CInt(Left(strYYYY,4))
		if len(strYYYY) = 6 then
			if CInt(Right(strYYYY,2)) < 4 then
				intYear = intYear - 1
			end if
		end if
		dim	intMonth
		intMonth = 3 + i
		if intMonth > 12 then
			intMonth = intMonth - 12
			intYear = intYear + 1
		end if
		GetYM = "" & intYear & Right("0" & intMonth,2)
	End Function
	'-----------------------------------------------------------------------
	'CheckDt() 日付チェック
	'-----------------------------------------------------------------------
	Private Function CheckDt(byVal strDt)
		if len(strYYYY) = 4 then
			CheckDt = True
			exit function
		end if
		if len(strYYYY) = 6 then
			if strYYYY = strDt then
				CheckDt = True
				exit function
			end if
		end if
		CheckDt = False
	End Function
	'-----------------------------------------------------------------------
	'ATrim() 全ての空白を取り除く
	'-----------------------------------------------------------------------
	Private	Function ATrim(byVal strV)
		strV = Trim(strV)
		strV = Replace(strV," ","")
		strV = Replace(strV,"　","")
		ATrim = strV
	End Function
	'-------------------------------------------------------------------
	'OpenDB
	'-------------------------------------------------------------------
    Private Function OpenDB()
		Debug ".OpenDB():" & strDBName
		Set objDB = Wscript.CreateObject("ADODB.Connection")
		objDB.commandTimeout = 0
'		objDB.CursorLocation = adUseClient
		Call objDB.Open(strDbName)
    End Function
	'-------------------------------------------------------------------
	'CloseDB
	'-------------------------------------------------------------------
    Private Function CloseDB()
		Debug ".CloseDB():" & strDBName
		Call objDB.Close()
		set objDB = Nothing
    End Function
	'-------------------------------------------------------------------
	'文字列追加 strSql
	'-------------------------------------------------------------------
	Private	Function AddSql(byVal strV)
		if strV = "" then
			strSql = strV
		end if
		if strSql <> "" then
			strSql = strSql & " "
		end if
		strSql = strSql & strV
	End Function
	'-----------------------------------------------------------------------
	'デバッグ用 /debug
	'-----------------------------------------------------------------------
	Public Sub Debug(byVal strMsg)
		if WScript.Arguments.Named.Exists("debug") then
			Wscript.StdErr.WriteLine strMsg
		end if
	End Sub
	'-----------------------------------------------------------------------
	'メッセージ表示
	'-----------------------------------------------------------------------
	Public Sub Disp(byVal strMsg)
		Wscript.Echo strMsg
	End Sub
	'-----------------------------------------------------------------------
	'オプション取得
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
End Class
