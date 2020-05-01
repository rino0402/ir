Option Explicit
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
WScript.Quit Main()
'-----------------------------------------------------------------------
'使用方法
'-----------------------------------------------------------------------
Private Sub usage()
	Wscript.Echo "計画用作業時間ファイル(Excel)読込 ver 0.5 2017.04.27 2017年度版"
	Wscript.Echo "loadplan03.vbs [option] <センター> <年度>[mm] <filename.xls> [シート名]"
	Wscript.Echo " /visible Excelを画面表示"
	Wscript.Echo " /debug   デバッグ"
	Wscript.Echo "cscript//nologo loadplan03.vbs H 201704 計画用作業時間_H_2017.xls"
	Wscript.Echo "..."
	Wscript.Echo "cscript//nologo attendance.vbs H 201704"
	Wscript.Echo "cscript//nologo prev.vbs H 201604"
End Sub
'-----------------------------------------------------------------------
'Excel
'2017.04.27 新規
'-----------------------------------------------------------------------
Const xlUp = -4162

Class Excel
	Private	strDBName
	Private	objDB
	Private	objExcel
	Private	objBook
	Private	strCenterCD
	Private	strYYYY
	Private	strFilename
	Private	strSheetName
	Private	strBookName
	Private	objSheet
	Private	lngMaxRow
	Private	lngRow
	Private	strSheetType
	Private	flgDeletePlan
	Private	strSql
	'-----------------------------------------------------------------------
	'Class_Initialize
	'-----------------------------------------------------------------------
	Private Sub Class_Initialize
		Debug ".Class_Initialize()"
		set	objExcel = nothing
		set	objBook = nothing
		strDBName = GetOption("db"	,"ir")
		set objDB = nothing
	End Sub
	'-----------------------------------------------------------------------
	'Class_Terminate
	'-----------------------------------------------------------------------
    Private Sub Class_Terminate
		Debug ".Class_Terminate()"
		set	objBook = nothing
		set	objExcel = nothing
		set objDB = nothing
    End Sub
	'-----------------------------------------------------------------------
	'Init() オプションチェック
	'-----------------------------------------------------------------------
    Public Function Init()
		Debug ".Init()"
		strFilename = ""
		strCenterCD = ""
		strYYYY		= ""
		strSheetName = ""
		strPassword = ""
		flgDeletePlan	= False
		dim	strArg
		Init = ""
		For Each strArg In WScript.Arguments.UnNamed
			if strCenterCD = "" then
				strCenterCD = strArg
			elseif strYYYY = "" then
				strYYYY = strArg
			elseif strFilename = "" then
				strFilename = strArg
			elseif strPassword = "" then
				strPassword = strArg
			else
				Init = "オプションエラー:" & strArg
				Disp Init
				Exit Function
			end if
		Next
		if strFileName = "" then
			Init = "ファイルを指定して下さい."
			Disp Init
			Exit Function
		end if
		For Each strArg In WScript.Arguments.Named
	    	select case lcase(strArg)
			case "db"
			case "visible"	' Excelを画面表示
			case "debug"	' デバッグ
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
		CloseDb
	End Function
	'-----------------------------------------------------------------------
	'Load() 読込
	'-----------------------------------------------------------------------
    Public Function Load()
		Debug ".Load():" & strFileName
		Call CreateExcel()
		Call OpenBook(strFileName)
		Call LoadBook()
		Call CloseBook()
	End Function
	'-------------------------------------------------------------------
	'LoadBook() 作業時間Excel読込
	'-------------------------------------------------------------------
    Private Function LoadBook()
		Debug ".LoadBook()"
		for each objSheet in objBook.Worksheets
			Debug ".LoadBook():" & objSheet.Name
			LoadSheet
		next
    End Function
	'-------------------------------------------------------------------
	'LoadSheet() 作業時間Excelシート読込
	'-------------------------------------------------------------------
    Private Function LoadSheet()
		Debug ".LoadSheet():" & objSheet.Name
		select case GetSheetType()
		case "人件費"	'年間勤務予定
			LoadPlan
		case "工料仕入"
'			LoadKoryo
		case "経費配分表"
			LoadTable
		end select
'		lngMaxRow = objSheet.Range("C65535").End(xlUp).Row
'		Debug ".LoadSheet():MaxRow=" & lngMaxRow
'		DeleteSheet
'		for lngRow = 11 to lngMaxRow
'			InsertSql
'		next
    End Function
	'-----------------------------------------------------------------------
	'シートの種別をチェック
	'-----------------------------------------------------------------------
	Private Function GetSheetType()
		Debug ".GetSheetType():" & objSheet.Name
		strSheetType = ""
		select case objSheet.Name
		case "工料仕入時間計画 (48期)"
			' 滋賀DC
			strSheetType = "工料仕入"
			GetSheetType = strSheetType
			Exit Function
		end select
		if objSheet.Range("A1") = "経費配分表" then
			if objSheet.Range("A3") = "配分コード" then
				if objSheet.Range("B3") = "配分名称" then
					strSheetType = "経費配分表"
					GetSheetType = strSheetType
					Exit Function
				end if
			end if
		end if
		dim	strCol
		strCol = "B"
		if objSheet.Range("C5") = "勤務時間" then
			strCol = "C"
		end if
		if objSheet.Range(strCol & "5") = "残業時間" and _
		   objSheet.Range(strCol & "6") = "給与（ｾﾝﾀｰ）" and _
		   objSheet.Range(strCol & "7") = "残業手当" and _
		   objSheet.Range(strCol & "8") = "通勤手当" and _
		   objSheet.Range(strCol & "9") = "法定福利" and _
		   objSheet.Range(strCol & "10") = "労働保険" and _
		   objSheet.Range(strCol & "11") = "賞与引当" and _
		   objSheet.Range(strCol & "12") = "退職引当" then
			strSheetType = "人件費"
			GetSheetType = strSheetType
			Exit Function
		end if
		if objSheet.Range(strCol & "5") = "残業時間" and _
		   objSheet.Range(strCol & "6") = "給与手当" and _
		   objSheet.Range(strCol & "7") = "残業手当" and _
		   objSheet.Range(strCol & "8") = "通勤手当" and _
		   objSheet.Range(strCol & "9") = "法定福利" and _
		   objSheet.Range(strCol & "10") = "労働保険" and _
		   objSheet.Range(strCol & "11") = "賞与引当" and _
		   objSheet.Range(strCol & "12") = "退職引当" then
			strSheetType = "人件費"
			GetSheetType = strSheetType
			Exit Function
		end if
		if objSheet.Range(strCol & "5") = "勤務時間" and _
		   objSheet.Range(strCol & "6") = "給与手当" and _
		   objSheet.Range(strCol & "7") = "残業手当" and _
		   objSheet.Range(strCol & "8") = "通勤手当" and _
		   objSheet.Range(strCol & "9") = "法定福利" and _
		   objSheet.Range(strCol & "10") = "労働保険" then
			strSheetType = "人件費"
			GetSheetType = strSheetType
			Exit Function
		end if
		if objSheet.Range("B5") = "時間" and _
		   objSheet.Range("B6") = "仕入金額" then
			strSheetType = "人件費"
			GetSheetType = strSheetType
			Exit Function
		end if
	End Function
	'-----------------------------------------------------------------------
	'LoadPlan() 人件費 シート読込 → 勤怠テーブル(Attendance)へ登録
	'-----------------------------------------------------------------------
	Private Function LoadPlan()
		Debug ".LoadPlan():" & objSheet.Name
		DeletePlan
		lngMaxRow = objSheet.Range("A65536").End(xlUp).Row
		for lngRow = 5 to lngMaxRow
			Debug ".LoadPlan():" & objSheet.name & ":" & lngRow & "/" & lngMaxRow
			LoadPlanRow
		next
	End Function
	'-----------------------------------------------------------------------
	'DeletePlan() 人件費 シート読込 → 勤怠テーブル(Attendance)削除
	'-----------------------------------------------------------------------
	Private Function DeletePlan()
		Debug ".DeletePlan():" & objSheet.Name
		if flgDeletePlan = True then
			exit function
		end if
		dim	strSql
'		strSql = "delete from Attendance"
		strSql = "update Attendance set Plan = 0"
		strSql = strSql & " where CenterCD = '" & strCenterCD & "'"
		select case Len(strYYYY)
		case 4
			strSql = strSql & " and DT between '" & GetYM(1) & "' and '" & GetYM(12) & "'"
		case 6
			strSql = strSql & " and DT = '" & strYYYY & "'"
		case else
			strSql = ""
		end select
		if strSql = "" then
			exit function
		end if
		Debug ".DeletePlan():" & strSql
		Wscript.StdOut.WriteLine strSql
		objDb.Execute strSql
		flgDeletePlan = True
	End Function
	'-----------------------------------------------------------------------
	'LoadPlanRow() 人件費 行読込 → 勤怠テーブル(Attendance)へ登録
	'-----------------------------------------------------------------------
	Private Function LoadPlanRow()
		Debug ".LoadPlanRow():" & objSheet.Name & " " & lngRow
		dim	strHaibunCD
		strHaibunCD = GetHaibunCode()
		if strHaibunCD = "" then
			Exit Function
		end if
		Debug ".LoadPlanRow():" & objSheet.Name & " " & lngRow & " " & strHaibunCD & " " & strPersonCD
		dim	strPersonCD
		strPersonCD = GetPersonCode()
		if strPersonCD = "" then
			Exit Function
		end if
		Debug ".LoadPlanRow():" & objSheet.Name & " " & lngRow & " " & strHaibunCD & " " & strPersonCD
		dim	r
		dim	i
		if objSheet.Range("C5") = "勤務時間" then
			set r = objSheet.Range("D" & lngRow)
		else
			set r = objSheet.Range("C" & lngRow)
		end if
		for i = 1 to 12
			dim	strDt
			strDt = GetYM(i)
			if CheckDt(strDt) = True then
				Wscript.StdOut.Write strHaibunCD & " " & strPersonCD
				Wscript.StdOut.Write " " & strDt
				Wscript.StdOut.Write " " & CCur(r)
				Wscript.StdOut.Write " " & CCur(r.Offset(1,0))
				Wscript.StdOut.Write " " & CCur(r.Offset(2,0))
				Wscript.StdOut.Write " " & CCur(r.Offset(3,0))
				Wscript.StdOut.WriteLine
				AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"TM100",CCur(r)				'勤務時間
				AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"TM200",CCur(r.Offset(1,0))	'作業時間
				AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"TM300",CCur(r.Offset(2,0))	'非作業時間
				AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"TM400",CCur(r.Offset(3,0))	'有給時間

				if objSheet.Range("B6") = "給与手当" then
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ100",CCur(r.Offset(-8,0))	'給与手当
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ200",CCur(r.Offset(-7,0))	'残業手当
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ300",CCur(r.Offset(-6,0))	'通勤手当
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ400",CCur(r.Offset(-5,0))	'法定福利
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ500",CCur(r.Offset(-4,0))	'労働保険
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ600",CCur(r.Offset(-3,0))	'賞与引当
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ700",CCur(r.Offset(-2,0))	'退職引当
				else
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ100",CCur(r.Offset(-6,0))	'給与手当
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ200",CCur(r.Offset(-5,0))	'残業手当
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ300",CCur(r.Offset(-4,0))	'通勤手当
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ400",CCur(r.Offset(-3,0))	'法定福利
					AddPlan strCenterCD,strHaibunCD,strPersonCD,strDT,"ZZ500",CCur(r.Offset(-2,0))	'労働保険
				end if
			end if
			set	r = r.Offset(0,1)
		next
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
	'AddPlan() SQL insert
	'-----------------------------------------------------------------------
	Private Function AddPlan(byVal strCenterCD _
							,byVal strHaibunCD _
							,byVal strPersonCD _
							,byVal strDT _
							,byVal strKamokuCD _
							,byVal curPlan _
							)
		Debug ".AddPlan()" & strCenterCD & " " & strHaibunCD & " " & strPersonCD & " " & strDT & " " & strKamokuCD & " " & curPlan
		if InsertPlan(strCenterCD,strHaibunCD,strPersonCD,strDT,strKamokuCD,curPlan) <> 0 then
			Call UpdatePlan(strCenterCD,strHaibunCD,strPersonCD,strDT,strKamokuCD,curPlan)
		end if
	End Function
	'-----------------------------------------------------------------------
	'InsertPlan() SQL insert
	'-----------------------------------------------------------------------
	Private Function InsertPlan(byVal strCenterCD _
							,byVal strHaibunCD _
							,byVal strPersonCD _
							,byVal strDT _
							,byVal strKamokuCD _
							,byVal curPlan _
							)
		Debug ".InsertPlan()" & strCenterCD & " " & strHaibunCD & " " & strPersonCD & " " & strDT & " " & strKamokuCD & " " & curPlan
		dim	strSql
		strSql = "insert into attendance"
		strSql = strSql & " ("
		strSql = strSql & " DT"
		strSql = strSql & ",CenterCD"
		strSql = strSql & ",HaibunCD"
		strSql = strSql & ",PersonCD"
		strSql = strSql & ",SyushiCD"
		strSql = strSql & ",KamokuCD"
		strSql = strSql & ",Plan"
		strSql = strSql & ",Result"
		strSql = strSql & ",Prev"
		strSql = strSql & " ) values ("
		strSql = strSql & " '" & strDT & "'"
		strSql = strSql & ",'" & strCenterCD & "'"
		strSql = strSql & ",'" & strHaibunCD & "'"
		strSql = strSql & ",'" & strPersonCD & "'"
		strSql = strSql & ",''"
		strSql = strSql & ",'" & strKamokuCD & "'"
		strSql = strSql & "," & curPlan
		strSql = strSql & ",0"
		strSql = strSql & ",0"
		strSql = strSql & ")"
		Debug ".AddPlan():" & strSql
		on error resume next
			objDb.Execute strSql
			Debug "0x" & Hex(Err.Number)
			Debug Err.Description
			InsertPlan = Err.Number
		on error goto 0
	End Function
	'-----------------------------------------------------------------------
	'UpdatePlan() SQL insert
	'-----------------------------------------------------------------------
	Private Function UpdatePlan(byVal strCenterCD _
							,byVal strHaibunCD _
							,byVal strPersonCD _
							,byVal strDT _
							,byVal strKamokuCD _
							,byVal curPlan _
							)
		Debug ".UpdatePlan()" & strCenterCD & " " & strHaibunCD & " " & strPersonCD & " " & strDT & " " & strKamokuCD & " " & curPlan
		dim	strSql
		strSql = "update attendance"
		strSql = strSql & " set Plan =" & curPlan
		strSql = strSql & " where DT = '" & strDT & "'"
		strSql = strSql & " and CenterCD = '" & strCenterCD & "'"
		strSql = strSql & " and HaibunCD = '" & strHaibunCD & "'"
		strSql = strSql & " and PersonCD = '" & strPersonCD & "'"
		strSql = strSql & " and SyushiCD = ''"
		strSql = strSql & " and KamokuCD = '" & strKamokuCD & "'"
		Debug ".AddPlan():" & strSql
		on error resume next
			objDb.Execute strSql
			Debug "0x" & Hex(Err.Number)
			Debug Err.Description
			UpdatePlan = Err.Number
		on error goto 0
	End Function
	'-----------------------------------------------------------------------
	'GetHaibunCode() 人件費 配分コードを返す
	'-----------------------------------------------------------------------
	Private Function GetHaibunCode()
		Debug ".GetHaibunCode():" & objSheet.Name & " " & lngRow
		GetHaibunCode = ""
		dim	strCol
		strCol = "B"
		if objSheet.Range("C5") = "勤務時間" then
			strCol = "C"
		end if
		if objSheet.Range(strCol & lngRow) <> "勤務時間" then
			exit function
		end if
		if objSheet.Range(strCol & (lngRow + 1)) <> "作業時間" then
			exit function
		end if
		if objSheet.Range(strCol & (lngRow + 2)) <> "非作業時間" then
			exit function
		end if
		if objSheet.Range(strCol & (lngRow + 3)) <> "有給時間" then
			exit function
		end if
		GetHaibunCode = objSheet.Range("A" & lngRow)
	End Function
	'-----------------------------------------------------------------------
	'GetPersonCode() 人件費 従業員コード(名)を返す
	'-----------------------------------------------------------------------
	Private Function GetPersonCode()
		Debug ".GetPersonCode():" & objSheet.Name & " " & lngRow
		GetPersonCode = ""
		dim	strCol
		strCol = "B"
		if objSheet.Range("C5") = "勤務時間" then
			strCol = "C"
		end if
		if objSheet.Range(strCol & lngRow - 9) = "残業時間" then
			GetPersonCode = ATrim(objSheet.Range("A" & lngRow - 9))
			exit function
		end if
		if objSheet.Range(strCol & lngRow - 7) = "勤務時間" then
			GetPersonCode = ATrim(objSheet.Range("A" & lngRow - 7))
			exit function
		end if
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
	'-----------------------------------------------------------------------
	'LoadTable() 経費配分表シート読込 → 経費配分テーブル(Haibun)へ登録
	'-----------------------------------------------------------------------
	Private Function LoadTable()
		Debug ".LoadTable():" & objSheet.Name
		dim	strHaibunCode
		dim	strHaibunName
		dim	strSyushiCD
		dim	r
		dim	vCho
		dim	vKan
		dim	strYM

		strYM = left(strYYYY,4)
		if Len(strYYYY) = 6 then
			if right(strYYYY,2) <> "04" then
				Exit Function
			end if
		end if
		DeleteHaibun
		lngMaxRow = objSheet.Range("A65536").End(xlUp).Row
		for lngRow = 5 to lngMaxRow
			Debug ".LoadTable():" & objSheet.name & ":" & lngRow & "/" & lngMaxRow
			LoadTableRow
		next
	End Function
	'-----------------------------------------------------------------------
	'LoadTableRow() 経費配分表 １行読込
	'-----------------------------------------------------------------------
	Private Function LoadTableRow()
		dim	strHaibunCode
		strHaibunCode = objSheet.Range("A" & lngRow)
		dim	strHaibunName
		strHaibunName = objSheet.Range("B" & lngRow)
		Debug ".LoadTableRow():" & objSheet.Name & " " & lngRow & ":" & strHaibunCode & " " & strHaibunName
		if objSheet.Range("A" & lngRow).EntireRow.Hidden = True then
			Debug ".LoadTableRow():skip 非表示"
			Exit Function
		end if
		if strHaibunCode = "配分コード" then
			Debug ".LoadTableRow():skip " & strHaibunCode
			Exit Function
		end if
		if strHaibunCode = "" then
			Debug ".LoadTableRow():skip 空白"
			Exit Function
		end if
		dim	r
		Set r = objSheet.Range("C2")
		do While True
			r.UnMerge	' セルの結合を解除
			dim	strSyushiCD
			strSyushiCD = r
			Debug ".LoadTableRow():" & r.Address & "=" & strSyushiCD
			if strSyushiCD = "" then
				exit do
			end if
			if len(strSyushiCD) <> 3 then
				exit do
			end if
			dim	vCho
			dim	vKan
			vCho = GetChoKan(r,"直接")
			vKan = GetChoKan(r,"間接")
			Debug ".LoadTableRow():" & strHaibunCode & " " & strHaibunName & " " & strSyushiCD & ":" & vCho & "/" & vKan
			AddHaibun strYM,strHaibunCode,strSyushiCD,vCho,vKan
			Set r = r.Offset(0,2)
		loop
	End Function
	'-----------------------------------------------------------------------
	'GetChoKan() 経費配分表 値 直接/間接
	'-----------------------------------------------------------------------
	Private Function AddHaibun(byVal strYM,byVal strHaibunCode,byVal strSyushiCD,byVal vCho,byVal vKan)
		Debug ".AddHaibun():" & strCenterCD & " " & strYM & " " & strHaibunCode & " " & strSyushiCD & " " & vCho & " " & vKan
		dim	strSql
		strSql = "insert into Haibun"
		strSql = strSql & " ("
		strSql = strSql & " CenterCD"
		strSql = strSql & ",YM"
		strSql = strSql & ",HaibunCD"
		strSql = strSql & ",SyushiCD"
		strSql = strSql & ",ChokuRatio"
		strSql = strSql & ",KanRatio"
		strSql = strSql & " ) values ("
		strSql = strSql & " '" & strCenterCD & "'"
		strSql = strSql & ",'" & strYM & "'"
		strSql = strSql & ",'" & strHaibunCode & "'"
		strSql = strSql & ",'" & strSyushiCD & "'"
		strSql = strSql & "," & vCho
		strSql = strSql & "," & vKan
		strSql = strSql & ")"
		Debug ".AddHaibun():" & strSql
		objDb.Execute strSql
	End Function
	'-----------------------------------------------------------------------
	'GetChoKan() 経費配分表 値 直接/間接
	'-----------------------------------------------------------------------
	Private Function GetChoKan(objRng,byval strKubun)
		dim	rngTop
		set rngTop = objRng
		select case strKubun
		case "直接"
		case "間接"
			Set rngTop = rngTop.Offset(0,1)
		end select
		dim	strCol
		strCol = Split(rngTop.Address,"$")(1)
		v = GetValue(objSheet.Range(strCol & lngRow))
		dim	v
		GetChoKan = v
	End Function
	'-----------------------------------------------------------------------
	'GetValue() セルの値を返す
	'-----------------------------------------------------------------------
	Private Function GetValue(byval r)
		dim	v
		v = 0
		if isempty(r) = False then
			if isnumeric(r) then
				v = r
			end if
		end if
		GetValue = v
	End Function
	'-----------------------------------------------------------------------
	'DeleteHaibun() 経費配分テーブル(Haibun)削除
	'-----------------------------------------------------------------------
	Private	strYM
	Private Function DeleteHaibun()
		strYM = left(strYYYY,4)
		Debug ".DeleteHaibun():" & objSheet.Name & " " & strYM & " " & strCenterCD
		dim	strSql
		strSql = "delete " _
			   & " from Haibun" _
			   & " where YM = '" & strYM & "'" _
			   & " and  CenterCD = '" & strCenterCD & "'" _
			   & ""
		Debug ".DateleHaibun():" & strSql
		objDb.Execute strSql
	End Function
	'-------------------------------------------------------------------
	'DeleteSheet
	'-------------------------------------------------------------------
    Private Function DeleteSheet()
		Debug ".DeleteSheet()"
		Wscript.StdOut.Write objSheet.Name & ":削除中..."

		AddSql ""
		AddSql "delete from CsvTemp"
		AddSql " where Filename = '" & objSheet.Name & "'"
		Wscript.StdOut.Write ":" & strSql
		CallSql strSql
		Wscript.StdOut.WriteLine
    End Function
	'-------------------------------------------------------------------
	'InsertSql
	'-------------------------------------------------------------------
    Private Function InsertSql()
		Debug ".InsertSql()"
		Wscript.StdOut.Write objSheet.Name & ":" & lngRow & "/" & lngMaxRow

		dim	intCol
		intCol = 34

		AddSql ""
		AddSql "insert into CsvTemp"
		AddSql "(Filename"
		AddSql ",Row"
		dim	i
		for	i = 1 to intCol
			AddSql ",Col" & right("00" & i,2)
		next
		AddSql ",Col"
		AddSql ") values ("
		AddSql "'" & objSheet.Name & "'"
		AddSql "," & lngRow
		dim	objRange
		set objRange = objSheet.Range("A" & lngRow)
		for	i = 1 to intCol
			AddSql ",'" & Trim(objRange) & "'"
			set objRange = objRange.Offset(0,1)
		next
		AddSql "," & intCol
		AddSql ")"
		Wscript.StdOut.Write ":" & strSql
		CallSql strSql
		Wscript.StdOut.WriteLine
    End Function
	'-------------------------------------------------------------------
	'Sql実行
	'-------------------------------------------------------------------
	Public Function CallSql(byVal strSql)
		Debug ".CallSql():" & strSql
'		on error resume next
		Call objDB.Execute(strSql)
'		on error goto 0
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
	'-------------------------------------------------------------------
	'Excelの準備
	'-------------------------------------------------------------------
	Private Function CreateExcel()
		Debug ".CreateExcel()"
		if objExcel is nothing then
			Debug ".CreateExcel():CreateObject(Excel.Application)"
			Set objExcel = WScript.CreateObject("Excel.Application")
		end if
	end function
	'-------------------------------------------------------------------
	'AbsPath() 絶対パス
	'-------------------------------------------------------------------
	Private	Function AbsPath(byVal strPath)
		dim	objFso
		Set objFso = CreateObject("Scripting.FileSystemObject")
		AbsPath = objFso.GetAbsolutePathName(strPath)
		Set objFso = Nothing
	End Function
	'-------------------------------------------------------------------
	'Excel ファイルオープン
	'-------------------------------------------------------------------
	Private	strPassword
	Private Function OpenBook(byVal strBkName)
		Debug ".OpenBook()"
		if objBook is nothing then
			strBkName = AbsPath(strBkName)
			Debug ".OpenBook().Open:" & strBkName
			Wscript.StdOut.Write strBkName & " ..."
			on error resume next
			Set objBook = objExcel.Workbooks.Open(strBkName,False,True,,strPassword)
			Wscript.StdOut.WriteLine Err.Number & ":" & Err.Description
			on error goto 0
		end if
	end function
	'-------------------------------------------------------------------
	'Excel ファイルクローズ
	'-------------------------------------------------------------------
	Private Function CloseBook()
		Debug ".CloseBook()"
		if not objBook is nothing then
			Debug ".CloseBook().Close:" & objBook.Name
			Call objBook.Close(False)
			set objBook = nothing
		end if
	end function
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
'-----------------------------------------------------------------------
'メイン
'-----------------------------------------------------------------------
Private Function Main()
	dim	objExcel
	Set objExcel = New Excel
	if objExcel.Init() <> "" then
		call usage()
		exit function
	end if
	call objExcel.Run()
End Function
