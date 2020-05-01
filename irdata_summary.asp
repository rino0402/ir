<%
Option Explicit
Response.Buffer = false		'ページ出力をバッファに格納するかどうか
Response.Expires = -1		'ブラウザ上にキャッシュされるページの有効期限が切れるまでの時間
%>
<%
Function GetVersion()
	GetVersion = "2012.05.04 新規作成"
	GetVersion = "2012.05.07 「結果をコピー」で検索結果をクリップボードにコピーするように対応"
	GetVersion = "2012.05.07 事業区分別 計画差(売上/利益)の計算式を修正"
	GetVersion = "2012.05.07 事業区分概況書 の対応"
	GetVersion = "2012.05.08 累計の検索範囲が201104～になっていたのを201204～に修正"
	GetVersion = "2012.05.11 経営概況 事業区分別に「_その他」を追加"
	GetVersion = "2012.12.15 経営概況 2013年度での集計対応"
	GetVersion = "2012.12.20 インターフェース改善中...外部変数削除"
	GetVersion = "2012.12.20 出力形式:事業区分概況書(年間-月別) 追加"
	GetVersion = "2013.01.30 出力形式:事業区分概況書(着地) 計画1～3月の対応"
	GetVersion = "2013.05.10 出力形式:事業区分概況書(間接あり時間チェック) ※各種時間のチェックに使用して下さい"
	GetVersion = "2013.05.11 出力形式:一覧表 科目/事業区分ごとに値をチェックできます。"
	GetVersion = "2013.05.28 事業区分=空白 で検索した場合、各時間の値が大きくなる不具合を修正"
	GetVersion = "2013.07.23 収支での集計／検索に対応"
	GetVersion = "2013.11.15 出力形式:事業区分概況書(着地:計画11～3月)の対応"
	GetVersion = "2014.07.25 出力形式:事業区分概況書(着地:計画(7,8,9,10～3))の対応"
	GetVersion = "2014.11.28 出力形式:事業区分概況書(年間) 期の表示間違いを訂正"
	GetVersion = "2015.01.20 出力形式:年間事業区分別概況書(月別) 時間チェック欄を追加（作業中・・・)"
	GetVersion = "2015.01.23 出力形式:事業区分概況書(着地:計画)の不具合修正"
	GetVersion = "2015.09.17 バグ：年間事業区分別概況書 が 前年計画と今期計画になっている"
	GetVersion = "2015.09.18 事業区分概況書(年間)：今期見通しが前年実績になっていた不具合修正"
	GetVersion = "2017.10.30 経営概況:売上原価／粗利益／経費を比例費/限界利益／固定費に変更"
	GetVersion = "2017.11.14 修正中(80% 明日に続くm(__)m)...年間事業区分別概況書(月別) 今期見通し/来期計画"
	GetVersion = "2017.11.15 来期計画が正しく集計できない不具合を修正"
	GetVersion = "2017.11.15 来期計画「その他」でエラーが発生する不具合を修正"
	GetVersion = "2018.11.13 年間事業区分別概況書(月別)で「Null 値の使い方が不正です」が発生する不具合修正"
	GetVersion = "2020.01.27"
End Function

Function GetDbName()
	GetDbName	= "IR"
End Function

%>
<%
Function YKSub(byVal strCenterCD,byVal strYM,byVal strYM2,byval strJKubun,byval strKubun)
	dim	s
	s = "select"
	s = s & " sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst04"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst05"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst06"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst07"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst08"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst09"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst10"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst11"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst12"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst01"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst02"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) ARst03"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) APln04"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) APln05"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) APln06"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) APln07"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) APln08"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) APln09"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) APln10"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) APln11"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) APln12"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) APln01"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) APln02"
	s = s & ",sum(if(KamokuCD like 'A%' and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) APln03"

	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst04"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst05"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst06"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst07"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst08"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst09"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst10"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst11"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst12"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst01"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst02"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) BRst03"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) BPln04"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) BPln05"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) BPln06"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) BPln07"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) BPln08"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) BPln09"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) BPln10"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) BPln11"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) BPln12"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) BPln01"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) BPln02"
	s = s & ",sum(if(KamokuCD like 'B%' and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) BPln03"

	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst04"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst05"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst06"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst07"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst08"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst09"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst10"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst11"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst12"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst01"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst02"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C1Rst03"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) C1Pln04"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) C1Pln05"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) C1Pln06"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) C1Pln07"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) C1Pln08"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) C1Pln09"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) C1Pln10"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) C1Pln11"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) C1Pln12"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) C1Pln01"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) C1Pln02"
	s = s & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400') and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) C1Pln03"

	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst04"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst05"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst06"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst07"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst08"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst09"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst10"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst11"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst12"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst01"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst02"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C2Rst03"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) C2Pln04"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) C2Pln05"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) C2Pln06"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) C2Pln07"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) C2Pln08"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) C2Pln09"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) C2Pln10"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) C2Pln11"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) C2Pln12"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) C2Pln01"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) C2Pln02"
	s = s & ",sum(if(KamokuCD in ('C0500','C0600') and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) C2Pln03"

	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst04"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst05"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst06"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst07"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst08"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst09"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst10"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst11"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst12"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst01"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst02"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) C9Rst03"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) C9Pln04"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) C9Pln05"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) C9Pln06"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) C9Pln07"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) C9Pln08"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) C9Pln09"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) C9Pln10"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) C9Pln11"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) C9Pln12"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) C9Pln01"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) C9Pln02"
	s = s & ",sum(if(KamokuCD in ('C9999') and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) C9Pln03"

	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst04"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst05"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst06"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst07"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst08"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst09"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst10"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst11"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst12"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst01"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst02"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) DRst03"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) DPln04"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) DPln05"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) DPln06"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) DPln07"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) DPln08"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) DPln09"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) DPln10"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) DPln11"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) DPln12"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) DPln01"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) DPln02"
	s = s & ",sum(if(KamokuCD in ('D0100') and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) DPln03"

	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 4) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst04"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 5) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst05"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 6) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst06"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 7) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst07"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 8) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst08"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 9) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst09"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM,10) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst10"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM,11) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst11"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM,12) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst12"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 1) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst01"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 2) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst02"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM, 3) & "',if(YM <= '" & strYM & "',Result,Plan),0)) X2Rst03"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 4) & "',Plan,0)) X2Pln04"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 5) & "',Plan,0)) X2Pln05"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 6) & "',Plan,0)) X2Pln06"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 7) & "',Plan,0)) X2Pln07"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 8) & "',Plan,0)) X2Pln08"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 9) & "',Plan,0)) X2Pln09"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2,10) & "',Plan,0)) X2Pln10"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2,11) & "',Plan,0)) X2Pln11"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2,12) & "',Plan,0)) X2Pln12"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 1) & "',Plan,0)) X2Pln01"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 2) & "',Plan,0)) X2Pln02"
	s = s & ",sum(if(KamokuCD in ('X0200') and YM = '" & GetNendo(strYM2, 3) & "',Plan,0)) X2Pln03"

	s = s & " from IrData"
	s = s & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM2,3) &"'"
	s = s & "   AND CenterCD = '" & strCenterCD &"'"
	if strJKubun <> "" then
		s = s & vbcrlf & "   AND SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strJKubun & "')"
	else
		s = s & vbcrlf & "   AND SyushiCd <> ''"
	end if
	dim	strSyushiCD		  '12345678
	strSyushiCD = GetRequest("SyushiCD","")
	if left(strKubun,8) = "_Syushi_" then
		strSyushiCD = right(RTrim(strKubun),3)
	end if
	s = s & SqlWhere("and", "SyushiCd", strSyushiCD)
'	if strSyushiCD <> "" then
'		s = s & vbcrlf & "   AND SyushiCd = '" & strSyushiCD & "'"
'	end if
	YKSub = s
End Function
Function TMSub(byVal strCenterCD,byVal strYM,byVal strYM2,byval strJKubun,byval strKubun,byval strFlg)
	dim	s
	s = "select"
	s = s & " sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst04"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst05"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst06"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst07"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst08"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst09"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst10"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst11"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst12"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst01"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst02"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) ARst03"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) APln04"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) APln05"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) APln06"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) APln07"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) APln08"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) APln09"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) APln10"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) APln11"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) APln12"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) APln01"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) APln02"
	s = s & ",sum(if(KamokuCD in ('TM101','TM102') and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) APln03"
select case strFlg
case "1","3"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst04"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst05"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst06"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst07"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst08"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst09"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst10"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst11"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst12"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst01"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst02"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM101Rst03"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM101Pln04"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM101Pln05"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM101Pln06"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM101Pln07"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM101Pln08"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM101Pln09"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM101Pln10"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM101Pln11"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM101Pln12"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM101Pln01"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM101Pln02"
	s = s & ",sum(if(KamokuCD = 'TM101' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM101Pln03"
end select
select case strFlg
case "2","3"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst04"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst05"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst06"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst07"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst08"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst09"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst10"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst11"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst12"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst01"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst02"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM102Rst03"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM102Pln04"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM102Pln05"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM102Pln06"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM102Pln07"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM102Pln08"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM102Pln09"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM102Pln10"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM102Pln11"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM102Pln12"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM102Pln01"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM102Pln02"
	s = s & ",sum(if(KamokuCD = 'TM102' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM102Pln03"
end select
select case strFlg
case "1","3"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst04"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst05"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst06"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst07"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst08"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst09"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst10"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst11"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst12"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst01"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst02"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM201Rst03"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM201Pln04"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM201Pln05"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM201Pln06"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM201Pln07"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM201Pln08"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM201Pln09"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM201Pln10"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM201Pln11"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM201Pln12"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM201Pln01"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM201Pln02"
	s = s & ",sum(if(KamokuCD = 'TM201' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM201Pln03"
end select
select case strFlg
case "2","3"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst04"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst05"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst06"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst07"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst08"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst09"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst10"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst11"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst12"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst01"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst02"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM202Rst03"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM202Pln04"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM202Pln05"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM202Pln06"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM202Pln07"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM202Pln08"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM202Pln09"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM202Pln10"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM202Pln11"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM202Pln12"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM202Pln01"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM202Pln02"
	s = s & ",sum(if(KamokuCD = 'TM202' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM202Pln03"
end select
select case strFlg
case "1","3"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst04"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst05"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst06"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst07"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst08"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst09"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst10"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst11"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst12"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst01"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst02"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM301Rst03"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM301Pln04"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM301Pln05"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM301Pln06"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM301Pln07"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM301Pln08"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM301Pln09"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM301Pln10"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM301Pln11"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM301Pln12"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM301Pln01"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM301Pln02"
	s = s & ",sum(if(KamokuCD = 'TM301' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM301Pln03"
end select
select case strFlg
case "2","3"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst04"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst05"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst06"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst07"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst08"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst09"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst10"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst11"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst12"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst01"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst02"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM302Rst03"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM302Pln04"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM302Pln05"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM302Pln06"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM302Pln07"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM302Pln08"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM302Pln09"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM302Pln10"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM302Pln11"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM302Pln12"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM302Pln01"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM302Pln02"
	s = s & ",sum(if(KamokuCD = 'TM302' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM302Pln03"
end select
select case strFlg
case "1","3"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst04"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst05"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst06"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst07"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst08"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst09"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst10"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst11"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst12"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst01"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst02"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM401Rst03"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM401Pln04"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM401Pln05"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM401Pln06"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM401Pln07"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM401Pln08"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM401Pln09"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM401Pln10"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM401Pln11"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM401Pln12"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM401Pln01"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM401Pln02"
	s = s & ",sum(if(KamokuCD = 'TM401' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM401Pln03"
end select
select case strFlg
case "2","3"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 4) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst04"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 5) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst05"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 6) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst06"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 7) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst07"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 8) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst08"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 9) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst09"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM,10) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst10"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM,11) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst11"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM,12) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst12"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 1) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst01"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 2) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst02"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM, 3) & "',if(DT <= '" & strYM & "',Result,Plan),0)) TM402Rst03"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 4) & "',Plan,0)) TM402Pln04"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 5) & "',Plan,0)) TM402Pln05"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 6) & "',Plan,0)) TM402Pln06"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 7) & "',Plan,0)) TM402Pln07"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 8) & "',Plan,0)) TM402Pln08"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 9) & "',Plan,0)) TM402Pln09"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2,10) & "',Plan,0)) TM402Pln10"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2,11) & "',Plan,0)) TM402Pln11"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2,12) & "',Plan,0)) TM402Pln12"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 1) & "',Plan,0)) TM402Pln01"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 2) & "',Plan,0)) TM402Pln02"
	s = s & ",sum(if(KamokuCD = 'TM402' and DT = '" & GetNendo(strYM2, 3) & "',Plan,0)) TM402Pln03"
end select
	s = s & " from Attendance"
	s = s & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM2,3) &"'"
	s = s & "   AND CenterCD = '" & strCenterCD &"'"
	if strJKubun <> "" then
		s = s & vbcrlf & "   AND SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strJKubun & "')"
	else
		s = s & vbcrlf & "   AND SyushiCd <> ''"
	end if
	dim	strSyushiCD		  '12345678
	strSyushiCD = GetRequest("SyushiCD","")
	if left(strKubun,8) = "_Syushi_" then
		strSyushiCD = right(RTrim(strKubun),3)
	end if
	s = s & SqlWhere("and", "SyushiCD", strSyushiCD)
'	if strSyushiCD <> "" then
'		s = s & vbcrlf & "   AND SyushiCd = '" & strSyushiCD & "'"
'	end if
	TMSub = s
End Function
Function GetYm()
	dim	intYear
	dim	intMonth
	dim	intDay
	dim	dt

	dt = Now()
	intYear		= Year(dt)
	intMonth	= Month(dt)
	intDay		= Day(dt)

	if intDay < 25 then
		intMonth = intMonth - 1
		if intMonth < 1 then
			intMonth = 12
			intYear	= intYear - 1
		end if
	end if
	GetYm = intYear & right("0" & intMonth,2)
End Function
'-----------------------------------------------------------
'出力形式リストを返す
'-----------------------------------------------------------
Function GetPTypeList()
	dim	strPTypeList
	strPTypeList = ""
	strPTypeList = strPTypeList & "<!-- GetPTypeList() start -->" & vbCrLF

	dim	strPType
	strPType = GetRequest("ptype","")

	strPTypeList = strPTypeList & GetOptionTag("pTable"			,"経営概況"					,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableSyushi"	,"経営概況+収支別"			,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableJM"		,"事業区分別当月詳細"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableJK"		,"事業区分概況書"			,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableJKKan"	,"事業区分概況書(直間時間チェック)"	,strPType) & vbCrLF
	strPTypeList = strPTypeList & "<OPTGROUP label=""事業区分概況書(着地)"">"
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku7"	,"着地:計画(7～3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku8"	,"着地:計画(8～3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku9"	,"着地:計画(9～3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku10"	,"着地:計画(10～3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku11"	,"着地:計画(11～3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku"	,"着地:計画(12～3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku1"	,"着地:計画(1～3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku2"	,"着地:計画(2～3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableChaku3"	,"着地:計画(3)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & "</OPTGROUP>"
	strPTypeList = strPTypeList & GetOptionTag("pTableJKYear"	,"事業区分概況書(年間)"		,strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableJKYearMonth2"	,"事業区分概況書(年間-月別) 今期見通し/今期計画",strPType) & vbCrLF
	strPTypeList = strPTypeList & GetOptionTag("pTableJKYearMonth"	,"事業区分概況書(年間-月別) 今期見通し/来期計画",strPType) & vbCrLF

	strPTypeList = strPTypeList & GetOptionTag("pList"			,"一覧表"					,strPType) & vbCrLF

	strPTypeList = strPTypeList & "<!-- GetPTypeList() end -->"
	GetPTypeList = strPTypeList
End Function
%>
<!--#include file="makeWhere.asp" -->
<%
'------------------------------------------------------
'初期処理
'------------------------------------------------------
if GetRequest("submit1","") <> "" then
	Call SetCookie("CenterCD")
end if
'------------------------------------------------------
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=shift_jis">
<LINK REL=STYLESHEET TYPE="text/css" HREF="ir.css" TITLE="CSS">
<TITLE>経営資料-経営概況</TITLE>
<!-- jdMenu head用 include 開始 -->
<link href="jquery.jdMenu.css" rel="stylesheet" type="text/css" />
<script src="jquery.js" type="text/javascript"></script>
<script src="jquery.dimensions.js" type="text/javascript"></script>
<script src="jquery.positionBy.js" type="text/javascript"></script>
<script src="jquery.bgiframe.js" type="text/javascript"></script>
<script src="jquery.jdMenu.js" type="text/javascript"></script>
<!-- jdMenu head用 include 終了 -->
</HEAD>
<SCRIPT LANGUAGE="JavaScript"><!--
	function DoCopy(arg){
		var doc = document.body.createTextRange();
		doc.moveToElementText(document.all(arg));
		doc.execCommand("copy");
		window.alert("クリップボードへコピーしました。\n貼り付けできます。" );
	}
--></SCRIPT>
<BODY>
<!-- jdMenu body用 include 開始 -->
<!--#include file="jdmenu-sdc-ir.asp" -->
<!-- jdMenu body用 include 終了 -->
  <FORM name="sqlForm"> <!--accept-charset="UTF-8"-->
	<table id="sqlTbl">
		<caption style="text-align:left;">経営資料-経営概況</caption>
		<tr>
			<th>集計年月</th>
			<th>センター</th>
			<th>事業区分</th>
			<th>収支</th>
		</tr>
		<tr valign="top">
			<td align="center">
				<INPUT class="input" TYPE="text" NAME="YM" VALUE="<%=GetRequest("YM",GetYM())%>" size="8" style="text-align:center;" required pattern="^[0-9]+$"><!-- placeholder="年月(yyyymm)を入力"-->
			</td>
			<td align="center">
				<select class="input" NAME="CenterCD">
				<%=GetCenterList()%>
				</select>
			</td>
			<td align="center">
				<select class="input" NAME="JKubun">
				<%=GetJKubunList()%>
				</select>
			</td>
			<td align="center">
				<INPUT class="input" TYPE="text" NAME="SyushiCD" VALUE="<%=GetRequest("SyushiCD","")%>" style="text-align:center;" placeholder="収支 ex.111,112"><!-- size="8" -->
			</td>
		</tr>
		<tr>
			<td colspan="4" nowrap>
				<label for="ptype">　出力形式：</label>
				<select class="input" NAME="ptype" id="ptype">
				<%=GetPTypeList()%>
				</select>
			</td>
		</tr>
		<tr bordercolor=White>
			<td colspan="4">
				<INPUT class="cssbutton" TYPE="submit" value="検索" id=submit1 name=submit1>
				<INPUT class="cssbutton" TYPE="reset" value="リセット" id=reset1 name=reset1 onClick="location.href='<%=Request.ServerVariables("URL")%>';">
				<span class="info_new"><%=GetVersion()%></span>
			</td>
		</tr>
	</table>
  </FORM>
<%	if len(GetRequest("submit1","")) > 0 then
		Server.ScriptTimeout = 3000
%>
	<SCRIPT LANGUAGE=javascript><!--
		sqlForm.disabled = true;
	//--></SCRIPT>
	<div>
		<INPUT TYPE="button" onClick="DoCopy('resultDiv')"
			 value="検索中...ScriptTimeout=<%=Server.ScriptTimeout%>" id="cpTblBtn" disabled>
	</div>

	<div id='resultDiv'>
	<TABLE id="resultTbl">
		<caption  style="text-align:left;"><%=GetCaption(GetRequest("YM",GetYM()),GetRequest("CenterCd",""),GetRequest("JKubun",""),GetRequest("ptype","pTable"))%></caption>
	<%
'		Response.Flush
		dim	objDb
		Set objDb = Server.CreateObject("ADODB.Connection")
		objDb.Open GetDbName()
	%>
		<thead>
			<%=MakeHeader(objDb,GetRequest("CenterCd",""),GetRequest("YM",GetYM()),GetRequest("ptype","pTable"))%>
		</thead>
		<tbody>
			<%=MakeBody(objDb,GetRequest("CenterCd",""),GetRequest("JKubun",""),GetRequest("YM",GetYM()),GetRequest("ptype","pTable"))%>
		</tbody>
	</TABLE>
	</div>
	<SCRIPT LANGUAGE=javascript><!--
		sqlForm.disabled = false;
		cpTblBtn.disabled = false;
		cpTblBtn.value = "結果をコピー";
	//--></SCRIPT>
<%
		call closeDb(objDb)
	end if
%>
<%
	Call endHtml()
%>

<% sub	closeDb(objDb)
	objDb.Close
	set objDb = nothing
end sub %>

<% sub	endHtml() 	%>
	<!-- endHtml() start	-->
	</BODY>
	</HTML>
	<!-- endHtml() end		-->
<% end sub			%>

<%
'-------------------------------------------------------------
'Table Caption
'-------------------------------------------------------------
Function GetCaption(byval strYm,byval strCenterCD,byval strJKubun,byval strPType)
	dim	strCaption

	strCaption = ""
	select case strPType
	case "pList"
		strCaption = "データ一覧"
	case "pTable"
		strCaption = GetPeriod(strYM) & "期 " & Right(strYM,2) & "月切 経営概況書"
	case "pTableSyushi"
		strCaption = GetPeriod(strYM) & "期 " & Right(strYM,2) & "月切 経営概況書+収支別"
	case "pTableJK"
		strCaption = GetPeriod(strYM) & "期 " & Right(strYM,2) & "月切 事業区分概況書"
	case "pTableJKKan"
		strCaption = GetPeriod(strYM) & "期 " & Right(strYM,2) & "月切 事業区分概況書(直間時間チェック)"
	case "pTableJM"
		strCaption = GetPeriod(strYM) & "期 " & Right(strYM,2) & "月切 事業区分別当月詳細"
	case "pTableJKYear"
		strCaption = "年間事業区分別概況書"
	case "pTableJKYearMonth","pTableJKYearMonth2"
		strCaption = "年間事業区分別概況書(月別)"
	end select
	strCaption = strCaption & " " & strYm
	strCaption = strCaption & " " & GetCenterName(strCenterCD)
	strCaption = strCaption & " " & strJKubun
	dim	strSyushi
	strSyushi = GetRequest("SyushiCD","")
	if strSyushi <> "" then
		strCaption = strCaption & " " & strSyushi
	end if
	GetCaption = RTrim(strCaption)
End Function

'-------------------------------------------------------------
'テーブル内容
'-------------------------------------------------------------
Function MakeBody(byVal objDb,byVal strCenterCD,byval strJKubun,byVal strYM,byval strTableType)
	dim	strHTML
	dim	objRs
	dim	strPersonName
	dim	strPersonCD
	dim	i
	dim	iKubun
	dim	aryWorkTM
	dim	strSql
	dim	errNumber
	dim	lngTotalTM
	dim	lngWorkTM

	strHTML = vbCrLf
	strHTML = strHTML & "<!-- MakeBody(" & strCenterCD & "," & strJKubun & "," & strYM & "," & strTableType & ")-->" & vbCrLf

	'------------------------------------------------------------------------------
	'レコード内容からHTMLを作成
	'------------------------------------------------------------------------------
	select case strTableType
	case "pList"
		strHTML = strHTML & GetTdList(objDb,strCenterCD,strJKubun,strYM,strTableType)
	case "pTable","pTableSyushi"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"">売上</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"売上")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
'		strHTML = strHTML & "<TH rowspan=""4"">売上原価</TH>"
		strHTML = strHTML & "<TH rowspan=""5"">比例費</TH>"
		strHTML = strHTML & "<TH>資材費</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"資材費")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>工料仕入</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"工料仕入")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>その他</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"その他仕入")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TH>直接人件費</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"直接人件費")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>合計</TH>"
'		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"仕入")
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"比例費")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"">限界利益</TH>"
'		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"粗利益")
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"限界利益")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH rowspan=""5"">固定費</TH>"
		strHTML = strHTML & "<TH>間接人件費</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"間接人件費")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>通常管理費</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"通常管理費")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>特別管理費</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"特別管理費")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>システム費</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"システム費")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>合計</TH>"
'		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"経費")
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"固定費")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"">営業利益</TH>"
		strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"営業利益")
		strHTML = strHTML & "</TR>"

		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"" rowspan=""2"">"
		if strCenterCD = "A" then
			strHTML = strHTML & "センター別"
		else
			select case strTableType
			case "pTable"
				strHTML = strHTML & "事業区分別"
			case "pTableSyushi"
				strHTML = strHTML & "収支別"
			end select
		end if
		strHTML = strHTML & "</TH>"
		strHTML = strHTML & "<TH colspan=""2"">当月</TH>"
		strHTML = strHTML & "<TH colspan=""2"">計画差</TH>"
		strHTML = strHTML & "<TH colspan=""2"">累計</TH>"
		strHTML = strHTML & "<TH colspan=""2"">計画差</TH>"
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>売上</TH>"
		strHTML = strHTML & "<TH>利益</TH>"
		strHTML = strHTML & "<TH>売上</TH>"
		strHTML = strHTML & "<TH>利益</TH>"
		strHTML = strHTML & "<TH>売上</TH>"
		strHTML = strHTML & "<TH>利益</TH>"
		strHTML = strHTML & "<TH>売上</TH>"
		strHTML = strHTML & "<TH>利益</TH>"
		strHTML = strHTML & "</TR>"

		if strCenterCD = "A" then
			' B:小野PC/D:滋賀物流/E:滋賀PC/F:奈良/G:大阪PC/H:袋井PC/I:広島
			dim	strC
			for each strC in Array("B","E","H","D","G","F","I")
				strHTML = strHTML & "<TR>"
				strHTML = strHTML & "<TD colspan=""2"">" & GetCenterName(strC) & "</TD>"
				strHTML = strHTML & GetTdValue(objDb,strC,strJKubun,strYM,"_")
				strHTML = strHTML & "</TR>"
			next
		else
			select case strTableType
			case "pTable"
				strSql = "select distinct JigyoKubunName from JigyoKubun where CenterCD = '" & strCenterCD & "' order by JigyoKubunName"
				set objRs = objDb.Execute(strSql)
				do while objRs.Eof = False
					strHTML = strHTML & "<TR>"
					strHTML = strHTML & "<TD colspan=""2"">" & GetJKubunLink(strYm,strCenterCD,GetFields(objRs,"JigyoKubunName")) & "</TD>"
					strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"_" & GetFields(objRs,"JigyoKubunName"))
					strHTML = strHTML & "</TR>"
					objRs.MoveNext
				loop
					strHTML = strHTML & "<TR>"
					strHTML = strHTML & "<TD colspan=""2"">" & GetJKubunLink(strYm,strCenterCD,"_その他") & "</TD>"
					strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"__その他")
					strHTML = strHTML & "</TR>"
			case "pTableSyushi"
				strSql = "select distinct SyushiCD,SyushiName from JigyoKubun where CenterCD = '" & strCenterCD & "' order by SyushiCD"
				set objRs = objDb.Execute(strSql)
				do while objRs.Eof = False
					strHTML = strHTML & vbCrLf & "<TR>"
					strHTML = strHTML & vbCrLf & "<TD colspan=""2"">" & GetSyushiLink(strYm,strCenterCD,GetFields(objRs,"SyushiCD"),GetFields(objRs,"SyushiName")) & "</TD>"
					strHTML = strHTML & vbCrLf & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"_Syushi_" & GetFields(objRs,"SyushiCD"))
					strHTML = strHTML & vbCrLf & "</TR>"
					objRs.MoveNext
				loop
			end select
		end if
	case "pTableJM"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" rowspan=""2"">事業区分別</TH>"
		strHTML = strHTML & "<TH colspan=""4"">当月</TH>"
		strHTML = strHTML & "<TH colspan=""4"">計画</TH>"
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH>売上</TH>"
		strHTML = strHTML & "<TH>利益</TH>"
		strHTML = strHTML & "<TH>直接人件費</TH>"
		strHTML = strHTML & "<TH>間接人件費</TH>"
		strHTML = strHTML & "<TH>売上</TH>"
		strHTML = strHTML & "<TH>利益</TH>"
		strHTML = strHTML & "<TH>直接人件費</TH>"
		strHTML = strHTML & "<TH>間接人件費</TH>"
		strHTML = strHTML & "</TR>"

		strSql = "select distinct JigyoKubunName from JigyoKubun where CenterCD = '" & strCenterCD & "' order by JigyoKubunName"
		set objRs = objDb.Execute(strSql)
		do while objRs.Eof = False
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TD colspan=""1"">" & GetJKubunLink(strYm,strCenterCD,GetFields(objRs,"JigyoKubunName")) & "</TD>"
			strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"JM_" & GetFields(objRs,"JigyoKubunName"))
			strHTML = strHTML & "</TR>"
			objRs.MoveNext
		loop
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TD colspan=""1"">" & GetJKubunLink(strYm,strCenterCD,"_その他") & "</TD>"
			strHTML = strHTML & GetTdValue(objDb,strCenterCD,strJKubun,strYM,"JM_その他")
			strHTML = strHTML & "</TR>"
	case "pTableChaku7","pTableChaku8","pTableChaku9","pTableChaku10","pTableChaku11","pTableChaku","pTableChaku1","pTableChaku2","pTableChaku3","pTableJKYear","pTableJKYearMonth","pTableJKYearMonth2"	' 年間事業区分別概況書
		dim	strJK
		select case strTableType
		case "pTableChaku7","pTableChaku8","pTableChaku9","pTableChaku10","pTableChaku11","pTableChaku","pTableChaku1","pTableChaku2","pTableChaku3"
			strJK = "CK"
		case "pTableJKYear","pTableJKYearMonth","pTableJKYearMonth2"	' 年間事業区分別概況書
			strJK = "YK"
		end select
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"">売上</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "売上")
		strHTML = strHTML & "</TR>"
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH rowspan=""3"">比例費</TH>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>売上原価</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "売上原価")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>直接人件費</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "直接人件費")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>計</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "比例費")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"" nowrap>限界利益</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "限界利益")
		strHTML = strHTML & "</TR>"
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH rowspan=""5"">固定費</TH>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>間接人件費</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "間接人件費")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>通常管理費</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "通常管理費")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>特別管理費</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "特別管理費")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>システム費</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "システム費")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>計</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "固定費")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""2"" nowrap>営業利益</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "営業利益")
		strHTML = strHTML & "</TR>"
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "</TR>"
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH rowspan=""6"">直接作業Ｈ</TH>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>勤務時間</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "勤務時間")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>作業時間</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "作業時間")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>非作業時間</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "非作業時間")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>有給時間</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "有給時間")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>売上工数</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "売上工数")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>工数(余裕率除)</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "工数(余裕率除)")
		strHTML = strHTML & "</TR>"
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH rowspan=""4"">間接作業Ｈ</TH>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>勤務時間</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "勤務時間(間)")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>作業時間</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "作業時間(間)")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>非作業時間</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "非作業時間(間)")
		strHTML = strHTML & "</TR>"
		strHTML = strHTML & "<TR>"
		strHTML = strHTML & "<TH colspan=""1"" nowrap>有給時間</TH>"
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "有給時間(間)")
		strHTML = strHTML & "</TR>" & vbCrLf
		'-----------------------------------------------------------------------
		if strJK = "YK" then
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TH colspan=""2"" nowrap>勤務時間</TH>"
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "勤務時間(計)")
			strHTML = strHTML & "</TR>"
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TH colspan=""2"" nowrap>作業時間</TH>"
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "作業時間(計)")
			strHTML = strHTML & "</TR>"
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TH colspan=""2"" nowrap>非作業時間</TH>"
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "非作業時間(計)")
			strHTML = strHTML & "</TR>"
			strHTML = strHTML & "<TR>"
			strHTML = strHTML & "<TH colspan=""2"" nowrap>有給時間</TH>"
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,strJK & "有給時間(計)")
			strHTML = strHTML & "</TR>" & vbCrLf
		end if
	case "pTableJK","pTableJKKan"
		'-----------------------------------------------------------------------
		' pTableJK		事業区分概況書
		' pTableJKKan	事業区分概況書(間接あり)
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""2"">売上</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK売上")
		strHTML = strHTML & "</TR>" & vbCrlf
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH rowspan=""3"">比例費</TH>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">売上原価</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK売上原価")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">直接人件費</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK直接人件費")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">計</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK比例費")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""2"">限界利益</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK限界利益")
		strHTML = strHTML & "</TR>" & vbCrlf
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH rowspan=""5"">固定費</TH>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">間接人件費</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK間接人件費")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">通常管理費</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK通常管理費")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">特別管理費</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK特別管理費")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">システム費</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JKシステム費")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""1"">計</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK固定費")
		strHTML = strHTML & "</TR>" & vbCrlf
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TH colspan=""2"">営業利益</TH>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK営業利益")
		strHTML = strHTML & "</TR>" & vbCrlf
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>" & vbCrlf
		for i = 1 to 18
			strHTML = strHTML & "<TD></TD>"
		next
		strHTML = strHTML & vbCrlf
		strHTML = strHTML & "</TR>" & vbCrlf
		'-----------------------------------------------------------------------
		strHTML = strHTML & "<TR>" & vbCrlf
		strHTML = strHTML & "<TD rowspan=""6"">直接作業Ｈ</TD>" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">勤務時間</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK勤務時間")
		strHTML = strHTML & "</TR>" & vbCrlf

		strHTML = strHTML & "<TR bgcolor=""lightyellow"">" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">作業時間</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK作業時間")
		strHTML = strHTML & "</TR>" & vbCrlf

		strHTML = strHTML & "<TR bgcolor=""whitesmoke"">" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">非作業時間</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK非作業時間")
		strHTML = strHTML & "</TR>" & vbCrlf

		strHTML = strHTML & "<TR bgcolor=""lavenderblush"">" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">有給時間</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK有給時間")
		strHTML = strHTML & "</TR>" & vbCrlf

		strHTML = strHTML & "<TR bgcolor=""lightcyan"">" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">売上工数</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK売上工数")
		strHTML = strHTML & "</TR>" & vbCrlf

		strHTML = strHTML & "<TR bgcolor=""lightcyan"">" & vbCrlf
		strHTML = strHTML & "<TD colspan=""1"">工数(余裕率除)</TD>" & vbCrlf
		strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK工数(余裕率除)")
		strHTML = strHTML & "</TR>" & vbCrlf
		if strTableType = "pTableJKKan" then
			'-----------------------------------------------------------------------
			strHTML = strHTML & "<TR>" & vbCrlf
			strHTML = strHTML & "<TD rowspan=""4"">間接作業Ｈ</TD>" & vbCrlf
			strHTML = strHTML & "<TD colspan=""1"">勤務時間</TD>" & vbCrlf
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK勤務時間(間)")
			strHTML = strHTML & "</TR>" & vbCrlf

			strHTML = strHTML & "<TR bgcolor=""lightyellow"">" & vbCrlf
			strHTML = strHTML & "<TD colspan=""1"">作業時間</TD>" & vbCrlf
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK作業時間(間)")
			strHTML = strHTML & "</TR>" & vbCrlf

			strHTML = strHTML & "<TR bgcolor=""whitesmoke"">" & vbCrlf
			strHTML = strHTML & "<TD colspan=""1"">非作業時間</TD>" & vbCrlf
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK非作業時間(間)")
			strHTML = strHTML & "</TR>" & vbCrlf

			strHTML = strHTML & "<TR bgcolor=""lavenderblush"">" & vbCrlf
			strHTML = strHTML & "<TD colspan=""1"">有給時間</TD>" & vbCrlf
			strHTML = strHTML & GetTdValueJK(objDb,strCenterCD,strJKubun,strYM,"JK有給時間(間)")
			strHTML = strHTML & "</TR>" & vbCrlf
			'--------------------------------------------
			'各種時間内訳
			'--------------------------------------------
			strHTML = strHTML & GetTdList(objDb,strCenterCD,strJKubun,strYM,"JK時間内訳")
		end if
	end select
	set objRs = Nothing
	strHTML = strHTML & "<!-- MakeBody() End -->" & vbCrLf
	MakeBody = strHTML
End Function

Function GetJKubunLink(byval strYm,byval strCenterCD,byval strJKubun)
	dim	strLink
	strLink = "<A href=""" & Request.ServerVariables("URL")
	strLink = strLink & "?YM=" & strYm
	strLink = strLink & "&CenterCD=" & strCenterCD
	strLink = strLink & "&JKubun=" & Server.URLEncode(strJKubun)
	strLink = strLink & "&ptype=pTableJK"
	strLink = strLink & "&submit1=0"
	strLink = strLink & """>"
	strLink = strLink & strJKubun
	strLink = strLink & "</A>"
	GetJKubunLink = strLink
End Function

Function GetSyushiLink(byval strYm,byval strCenterCD,byval strSyushiCD,byval strSyushiName)
	dim	strLink
	strLink = "<A href=""" & Request.ServerVariables("URL")
	strLink = strLink & "?YM=" & strYm
	strLink = strLink & "&CenterCD=" & strCenterCD
	strLink = strLink & "&SyushiCD=" & Server.URLEncode(strSyushiCD)
	strLink = strLink & "&ptype=pTableJK"
	strLink = strLink & "&submit1=0"
	strLink = strLink & """>"
	strLink = strLink & strSyushiCD & " " & strSyushiName
	strLink = strLink & "</A>"
	GetSyushiLink = strLink
End Function
'------------------------------------------------------------------------------
'事業区分概況(明細)
'------------------------------------------------------------------------------
dim	objRsYK
Function GetTdValueYK(objDb _
				   ,byval strCenterCD _
				   ,byval strJKubun _
				   ,byval strYM _
				   ,byval strKubun)
	dim	strYM2
	dim	strSql
	GetTdValueYK = ""
	dim	strHTML
	strHTML = ""
	select case strKubun
	case "YK売上"
		strYM2 = strYM
		select case GetRequest("ptype","pTable")	'strTableType
		case "pTableJKYearMonth"	' 年間事業区分別概況書
			strYM2 = clng(strYM) + 100
		end select
		strSql = YKSub(strCenterCD,strYM,strYM2,strJKubun,strKubun)
objDb.commandTimeout=600
		on error resume next
			set objRsYK = objDb.Execute(strSql)
			if Err.Number <> 0 then
				strHTML = MakeError(Err)
			end if
		on error goto 0
	case "YK売上原価","YK直接人件費","YK比例費","YK限界利益","YK間接人件費","YK通常管理費","YK特別管理費","YKシステム費","YK固定費","YK営業利益"
	case "YK勤務時間"
		strYM2 = strYM
		select case GetRequest("ptype","pTable")	'strTableType
		case "pTableJKYearMonth"	' 年間事業区分別概況書
			strYM2 = clng(strYM) + 100
		end select
		strSql = TMSub(strCenterCD,strYM,strYM2,strJKubun,strKubun,"1")
'		on error resume next
			set	objRsYK = nothing
objDb.commandTimeout=600
			set objRsYK = objDb.Execute(strSql)
'			if Err.Number <> 0 then
'				strHTML = MakeError(Err)
'			end if
'		on error goto 0
	case "YK勤務時間(間)"
		strYM2 = strYM
		select case GetRequest("ptype","pTable")	'strTableType
		case "pTableJKYearMonth"	' 年間事業区分別概況書
			strYM2 = clng(strYM) + 100
		end select
		strSql = TMSub(strCenterCD,strYM,strYM2,strJKubun,strKubun,"2")
			set	objRsYK = nothing
objDb.commandTimeout=600
			set objRsYK = objDb.Execute(strSql)
	case "YK勤務時間(計)"
		strYM2 = strYM
		select case GetRequest("ptype","pTable")	'strTableType
		case "pTableJKYearMonth"	' 年間事業区分別概況書
			strYM2 = clng(strYM) + 100
		end select
		strSql = TMSub(strCenterCD,strYM,strYM2,strJKubun,strKubun,"3")
			set	objRsYK = nothing
objDb.commandTimeout=600
			set objRsYK = objDb.Execute(strSql)
	case "YK勤務時間(間)","YK勤務時間(計)","YK作業時間","YK作業時間(間)","YK作業時間(計)","YK非作業時間","YK非作業時間(間)","YK非作業時間(計)","YK有給時間","YK有給時間(間)","YK有給時間(計)","YK売上工数","YK工数(余裕率除)"
	case else
		exit function
	end select
	if strHTML <> "" then
		GetTdValueYK = strHTML
		exit function
	end if
	strHTML = vbCrLf & "<!--" & strCenterCD & "," & strJKubun & "," & strYM & "," & strYM2 & "," & strKubun & "-->" & vbCrLf
	strHTML = strHTML & "<!--" & strSql & "-->" & vbCrLf
	dim	lngRst
	dim	lngPln
	dim	lngRstSum
	dim	lngPlnSum
	dim	lngRstSumA
	dim	lngPlnSumA
	dim	strRst
	dim	strPln
	dim	i
	lngRstSum = 0
	lngPlnSum = 0
	lngRstSumA = 0
	lngPlnSumA = 0
	strRst = ""
	strPln = ""
	for i = 1 to 12
		dim	m
		m = i + 3
		if m > 12 then
			m = m - 12
		end if
		m = right("0" & m,2)
		select case strKubun
		case "YK売上"
			lngRst = CDbl(objRsYK.Fields("ARst" & m))
			lngPln = CDbl(objRsYK.Fields("APln" & m))
		case "YK売上原価"
			lngRst = CDbl(objRsYK.Fields("BRst" & m))
			lngPln = CDbl(objRsYK.Fields("BPln" & m))
		case "YK直接人件費"
			lngRst = CDbl(objRsYK.Fields("C1Rst" & m)) - CDbl(objRsYK.Fields("X2Rst" & m))
			lngPln = CDbl(objRsYK.Fields("C1Pln" & m)) - CDbl(objRsYK.Fields("X2Pln" & m))
		case "YK比例費"
			lngRst = CDbl(objRsYK.Fields("BRst" & m)) + CDbl(objRsYK.Fields("C1Rst" & m)) - CDbl(objRsYK.Fields("X2Rst" & m))
			lngPln = CDbl(objRsYK.Fields("BPln" & m)) + CDbl(objRsYK.Fields("C1Pln" & m)) - CDbl(objRsYK.Fields("X2Pln" & m))
		case "YK限界利益"
			lngRst = CDbl(objRsYK.Fields("ARst" & m)) - (CDbl(objRsYK.Fields("BRst" & m)) + CDbl(objRsYK.Fields("C1Rst" & m)) - CDbl(objRsYK.Fields("X2Rst" & m)))
			lngPln = CDbl(objRsYK.Fields("APln" & m)) - (CDbl(objRsYK.Fields("BPln" & m)) + CDbl(objRsYK.Fields("C1Pln" & m)) - CDbl(objRsYK.Fields("X2Pln" & m)))
		case "YK間接人件費"
			lngRst = CDbl(objRsYK.Fields("X2Rst" & m))
			lngPln = CDbl(objRsYK.Fields("X2Pln" & m))
		case "YK通常管理費"
			lngRst = CDbl(objRsYK.Fields("C2Rst" & m))
			lngPln = CDbl(objRsYK.Fields("C2Pln" & m))
		case "YK特別管理費"
			lngRst = CDbl(objRsYK.Fields("C9Rst" & m))
			lngPln = CDbl(objRsYK.Fields("C9Pln" & m))
		case "YKシステム費"
			lngRst = CDbl(objRsYK.Fields("DRst" & m))
			lngPln = CDbl(objRsYK.Fields("DPln" & m))
		case "YK固定費"
			lngRst = CDbl(objRsYK.Fields("X2Rst" & m)) + CDbl(objRsYK.Fields("C2Rst" & m)) + CDbl(objRsYK.Fields("C9Rst" & m)) + CDbl(objRsYK.Fields("DRst" & m))
			lngPln = CDbl(objRsYK.Fields("X2Pln" & m)) + CDbl(objRsYK.Fields("C2Pln" & m)) + CDbl(objRsYK.Fields("C9Pln" & m)) + CDbl(objRsYK.Fields("DPln" & m))
		case "YK営業利益"
			lngRst = CDbl(objRsYK.Fields("ARst" & m)) - (CDbl(objRsYK.Fields("BRst" & m)) + CDbl(objRsYK.Fields("C1Rst" & m)) - CDbl(objRsYK.Fields("X2Rst" & m)))
			lngPln = CDbl(objRsYK.Fields("APln" & m)) - (CDbl(objRsYK.Fields("BPln" & m)) + CDbl(objRsYK.Fields("C1Pln" & m)) - CDbl(objRsYK.Fields("X2Pln" & m)))
			lngRst = lngRst - (CDbl(objRsYK.Fields("X2Rst" & m)) + CDbl(objRsYK.Fields("C2Rst" & m)) + CDbl(objRsYK.Fields("C9Rst" & m)) + CDbl(objRsYK.Fields("DRst" & m)))
			lngPln = lngPln - (CDbl(objRsYK.Fields("X2Pln" & m)) + CDbl(objRsYK.Fields("C2Pln" & m)) + CDbl(objRsYK.Fields("C9Pln" & m)) + CDbl(objRsYK.Fields("DPln" & m)))
		case "YK勤務時間"
			lngRst = GetFieldVal(objRsYK, "TM101Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM101Pln" & m)
		case "YK勤務時間(間)"
			lngRst = GetFieldVal(objRsYK, "TM102Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM102Pln" & m)
		case "YK勤務時間(計)"
			lngRst = GetFieldVal(objRsYK, "ARst" & m)
			lngPln = GetFieldVal(objRsYK, "APln" & m)
		case "YK作業時間"
			lngRst = GetFieldVal(objRsYK, "TM201Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM201Pln" & m)
		case "YK作業時間(間)"
			lngRst = GetFieldVal(objRsYK, "TM202Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM202Pln" & m)
		case "YK作業時間(計)"
			lngRst = GetFieldVal(objRsYK, "TM201Rst" & m) + GetFieldVal(objRsYK, "TM202Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM201Pln" & m) + GetFieldVal(objRsYK, "TM202Pln" & m)
		case "YK非作業時間"
			lngRst = GetFieldVal(objRsYK, "TM301Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM301Pln" & m)
		case "YK非作業時間(間)"
			lngRst = GetFieldVal(objRsYK, "TM302Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM302Pln" & m)
		case "YK非作業時間(計)"
			lngRst = GetFieldVal(objRsYK, "TM301Rst" & m) + GetFieldVal(objRsYK, "TM302Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM301Pln" & m) + GetFieldVal(objRsYK, "TM302Pln" & m)
		case "YK有給時間"
			lngRst = GetFieldVal(objRsYK, "TM401Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM401Pln" & m)
		case "YK有給時間(間)"
			lngRst = GetFieldVal(objRsYK, "TM402Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM402Pln" & m)
		case "YK有給時間(計)"
			lngRst = GetFieldVal(objRsYK, "TM401Rst" & m) + GetFieldVal(objRsYK, "TM402Rst" & m)
			lngPln = GetFieldVal(objRsYK, "TM401Pln" & m) + GetFieldVal(objRsYK, "TM402Pln" & m)
		case "YK売上工数"
			lngRst = 0
			lngPln = 0
		case else
			lngRst = 0
			lngPln = 0
		end select
		lngRstSum = lngRstSum + lngRst
		lngPlnSum = lngPlnSum + lngPln
		lngRstSumA = lngRstSumA + GetFieldVal(objRsYK, "ARst" & m)
		lngPlnSumA = lngPlnSumA + GetFieldVal(objRsYK, "APln" & m)
		strRst = strRst & "<td class=""Number"">" & GetNumber(lngRst,"") & "</td>"
		strPln = strPln & "<td class=""Number"">" & GetNumber(lngPln,"") & "</td>"
	next
	dim	strRt
	strRt = ""
	if lngRstSumA <> 0 then
		strRt = GetNumber(lngRstSum/lngRstSumA*100,"")
	end if
	strHTML = strHTML & "<td class=""Number"">" & GetNumber(lngRstSum,"") & "</td><td class=""Number"">" & strRt & "</td>" & strRst
	strRt = ""
	if lngPlnSumA <> 0 then
		strRt = GetNumber(lngPlnSum/lngPlnSumA*100,"")
	end if
	strHTML = strHTML & "<td class=""Number"">" & GetNumber(lngPlnSum,"") & "</td><td class=""Number"">" & strRt & "</td>" & strPln
	strHTML = strHTML & "<td class=""Number"">" & GetNumber(lngRstSum - lngPlnSum,"") & "</td>" & vbCrLf
	GetTdValueYK = strHTML
End Function

Function GetTdValueJK(byval objDb _
				   ,byval strCenterCD _
				   ,byval strJKubun _
				   ,byval strYM _
				   ,byval strKubun)
	GetTdValueJK = GetTdValueYK(objDb,strCenterCD,strJKubun,strYM,strKubun)
	if GetTdValueJK <> "" then
		exit function
	end if
	dim	strSql
	dim	objRs
	dim	strHTML
	'------------------------------------------------------------------------------
	'事業区分概況(明細)
	'------------------------------------------------------------------------------
	strHTML = ""
	strHTML = strHTML & "<!-- GetTdValueJK(" & strCenterCD & "," & strJKubun & "," & strYM & "," & strKubun & ")-->" & vbCrLf
	strSql = MakeSql(strCenterCD,strJKubun,strYM,strKubun)
	strHTML = strHTML & "<!-- " & vbCrLf & strSql & vbCrLf & ")-->" & vbCrLf
	if strSql <> "" then
		on error resume next
			set objRs = objDb.Execute(strSql)
			if Err.Number <> 0 then
				strHTML = MakeError(objErr)
			end if
			if objRs is Nothing then
				strHTML = strHTML & "<tr><td>" & objDb.Errors.Count & "</td>" & vbCrlf
		        Dim errX
		        For Each errX In objDb.Errors
					strHTML = strHTML & "<td>" & errX.Description & "</td>" & vbCrlf
		        Next
				strHTML = strHTML & "<td>" & strSql & "</td></tr>" & vbCrlf
'			End if
'			if not objRs is Nothing then
			else
				if objRs.Eof = False then
					dim	f
					for each f in objRs.Fields
						strHTML = strHTML & vbTab & "<td class=""Number"">" & GetNumber(GetFields(objRs,f.Name),"") & "</td>" & vbCrlf
					next
				end if
				set objRs = nothing
			end if
		on error goto 0
	end if
	strHTML = strHTML & "<!-- GetTdValueJK() End -->" & vbCrLf
	GetTdValueJK = strHTML
End Function


Function GetTdValue(byval objDb _
				   ,byval strCenterCD _
				   ,byval strJKubun _
				   ,byval strYM _
				   ,byval strKubun)
	dim	strSql
	dim	objRs
	dim	strHTML
	strHTML = ""
	'------------------------------------------------------------------------------
	'SQL実行
	'------------------------------------------------------------------------------
	strSql = MakeSql(strCenterCD,strJKubun,strYM,strKubun)
	on error resume next
		set objRs = objDb.Execute(strSql)
		if Err.Number <> 0 then
			strHTML = MakeError(objErr)
		end if
		if objRs is Nothing then
			strHTML = strHTML & vbCrLf & "<tr><td>" & objDb.Errors.Count & "</td>"
	        Dim errX
	        For Each errX In objDb.Errors
				strHTML = strHTML & vbCrLf & "<td>" & errX.Description & "</td>"
	        Next
			strHTML = strHTML & vbCrLf & "<td>" & strSql & "</td></tr>"
		End if
	on error goto 0
	if strHTML = "" then
		strHTML = strHTML & vbCrLf & "<!-- "
		strHTML = strHTML & vbCrLf & strSql
		strHTML = strHTML & vbCrLf & " -->"
		if objRs.Eof = False then
			dim	f
			for each f in objRs.Fields
				strHTML = strHTML & vbCrLf & "<td class=""Number"">" & GetNumber(GetFields(objRs,f.Name),"") & "</td>"
			next
		end if
	end if
	GetTdValue = strHTML
End Function

Function GetTdList(byval objDb _
				   ,byval strCenterCD _
				   ,byval strJKubun _
				   ,byval strYM _
				   ,byval strKubun)
	dim	strSql
	dim	objRs
	dim	strHTML
	strHTML = ""
	'------------------------------------------------------------------------------
	'データ一覧
	'------------------------------------------------------------------------------
	strHTML = ""
	strHTML = strHTML & "<!-- GetTdList(" & strCenterCD & "," & strJKubun & "," & strYM & "," & strKubun & ")-->" & vbCrLf
	strSql = MakeSql(strCenterCD,strJKubun,strYM,strKubun)
	strHTML = strHTML & "<!-- " & vbCrLf & strSql & vbCrLf & ")-->" & vbCrLf
	'------------------------------------------------------------------------------
	'SQL実行
	'------------------------------------------------------------------------------
	on error resume next
		set objRs = objDb.Execute(strSql)
		if Err.Number <> 0 then
			strHTML = MakeError(objErr)
		end if
		if objRs is Nothing then
			strHTML = strHTML & "<tr><td>" & objDb.Errors.Count & "</td>"
	        Dim errX
	        For Each errX In objDb.Errors
				strHTML = strHTML & "<td>" & errX.Description & "</td>"
	        Next
			strHTML = strHTML & "<td>" & strSql & "</td></tr>"
		End if
	on error goto 0
	if not objRs is Nothing then
		dim	f
		if strJKubun = "Header" then
				strHTML = strHTML & "<TR>" & vbCrlf
				for each f in objRs.Fields
					strHTML = strHTML & "<th>" & f.Name & "</th>" & vbCrlf
				next
				strHTML = strHTML & "</TR>" & vbCrlf
		else
			do while objRs.Eof = False
				for each f in objRs.Fields
					strHTML = strHTML & "<!--bgcolor " & f & "-->" & vbCrlf
					select case RTrim(f)
					case "作業時間(直接)","作業時間(間接)"
						strHTML = strHTML & "<TR bgcolor=""lightyellow"">" & vbCrlf
					case "非作業時間(直接)","非作業時間(間接)"
						strHTML = strHTML & "<TR bgcolor=""whitesmoke"">" & vbCrlf
					case "有給時間(直接)","有給時間(間接)"
						strHTML = strHTML & "<TR bgcolor=""lavenderblush"">" & vbCrlf
					case else
						strHTML = strHTML & "<TR>" & vbCrlf
					end select
					exit for
				next
				for each f in objRs.Fields
'					strHTML = strHTML & "<td>" & GetFields(objRs,f.Name) & "</td>" & vbCrlf
					strHTML = strHTML & GetFieldTd(objRs,f) & vbCrlf
				next
				strHTML = strHTML & "</TR>" & vbCrlf
				objRs.MoveNext
			Loop
		end if
	end if
	strHTML = strHTML & "<!-- GetTdList() End -->" & vbCrLf
	GetTdList = strHTML
End Function
'-------------------------------------------------------------
'<td>を返す
'-------------------------------------------------------------
Function GetFieldTd(objRs,f)
	dim	strTd
	dim	strValue
	strValue = GetFields(objRs,f.Name)
	strTd = "<td"
	strTd = strTd & " title=""" & f.name & " " & f.type & """"
	select case f.type
	Case 2 , 3 , 5 , 6 ,131	' 数値(Integer)
		strTd = strTd & " class=""Number"""
		strValue = GetNumber(strValue,"")
	case else
		strTd = strTd & " class=""Character"""
	end select
'	strTd = strTd & " nowrap"
	strTd = strTd & ">"
	strTd = strTd & strValue
	strTd = strTd & "</td>"
	GetFieldTd = strTd
End Function
'-------------------------------------------------------------
'配列クリア
'-------------------------------------------------------------
Function ClearArray(aryV())
	dim	i
	for i = lbound(aryV) to ubound(aryV)
		aryV(i) = 0
	next
	ClearArray = i
End Function

'-------------------------------------------------------------
'配列の加算
'-------------------------------------------------------------
Function AddArray(aryV,objRs)
	dim	i
	dim	aryFld
	aryFld = Array("A_Prev","A_Plan","A_Result","A_Margin","T_Prev","T_Plan","T_Result","T_Margin")
	for i = lbound(aryFld) to ubound(aryFld)
		aryV(i) = aryV(i) + CLng(GetFields(objRs,aryFld(i)))
	next
	AddArray = i
End Function

'-------------------------------------------------------------
'配列のTD要素を返す
'-------------------------------------------------------------
Function GetTdArray(aryV())
	dim	i
	dim	strTd
	strTd = ""
	for i = lbound(aryV) to ubound(aryV)
		strTd = strTd & "<TD>" & GetNumber(aryV(i),"") & "/TD" & vbcrlf
	next
	GetTdArray = strTd
End Function


'-------------------------------------------------------------
'Fieldの値を返す
'-------------------------------------------------------------
Function GetFields(byval objRs _
				  ,byval strFieldName _
				  )
	dim	v
	if strFieldName = "" then
		v = 0
	else
		select case strFieldName
		case "Y_Prev_Hi"
			v = GetPercent(GetFieldValue(objRs,"Y_Prev"),GetFieldValue(objRs,"Y_Prev_Hi"))
		case "Y_Plan_Hi"
			v = GetPercent(GetFieldValue(objRs,"Y_Plan"),GetFieldValue(objRs,"Y_Plan_Hi"))
'		case "Y_Margin"
'			v = CLng(GetFields(objRs,"Y_Plan")) - CLng(GetFields(objRs,"Y_Prev"))
		case "A_Result_Hi"
			v = GetPercent(GetFieldValue(objRs,"A_Result"),GetFieldValue(objRs,"A_Result_Hi"))
		case "T_Result_Hi"
			v = GetPercent(GetFieldValue(objRs,"T_Result"),GetFieldValue(objRs,"T_Result_Hi"))
		case "A_Plan_Hi"
			v = GetPercent(GetFieldValue(objRs,"A_Plan"),GetFieldValue(objRs,"A_Plan_Hi"))
		case "T_Plan_Hi"
			v = GetPercent(GetFieldValue(objRs,"T_Plan"),GetFieldValue(objRs,"T_Plan_Hi"))
		case "A_Prev_Hi"
			v = GetPercent(GetFieldValue(objRs,"A_Prev"),GetFieldValue(objRs,"A_Prev_Hi"))
		case "T_Prev_Hi"
			v = GetPercent(GetFieldValue(objRs,"T_Prev"),GetFieldValue(objRs,"T_Prev_Hi"))
		case "A_Prev_Margin"
			v = CLng(GetFields(objRs,"A_Result")) - CLng(GetFields(objRs,"A_Prev"))
		case "T_Prev_Margin"
			v = CLng(GetFields(objRs,"T_Result")) - CLng(GetFields(objRs,"T_Prev"))
		case "A_Margin"
			v = CLng(GetFields(objRs,"A_Result")) - CLng(GetFields(objRs,"A_Plan"))
		case "T_Margin"
			v = CLng(GetFields(objRs,"T_Result")) - CLng(GetFields(objRs,"T_Plan"))
		case "A_UriSa"
			v = CLng(GetFieldValue(objRs,"A_Uri")) - CLng(GetFieldValue(objRs,"A_UriSa"))
		case "A_RiekiSa"
			v = CLng(GetFieldValue(objRs,"A_Rieki")) - CLng(GetFieldValue(objRs,"A_RiekiSa"))
		case "T_UriSa"
			v = CLng(GetFieldValue(objRs,"T_Uri")) - CLng(GetFieldValue(objRs,"T_UriSa"))
		case "T_RiekiSa"
			v = CLng(GetFieldValue(objRs,"T_Rieki")) - CLng(GetFieldValue(objRs,"T_RiekiSa"))
		case else
			v = RTrim(objRs.Fields(strFieldName))
			if isnull(v) then
				v = 0
			end if
		end select
	end if
	GetFields = v
End Function

Function GetPercent(byval v1,byval v2)
	dim	v
	v = 0
	if clng(v2) <> 0 then
		v = clng(v1) * 100 / clng(v2)
		v = Round(v,0)
	end if
	GetPercent = v
End Function

Function GetFieldVal(byval objRs _
				  ,byval strFieldName _
				  )
	dim	v
	v = GetFieldValue(objRs, strFieldName)
	GetFieldVal = cdbl(v)
End Function

Function GetFieldValue(byval objRs _
				  ,byval strFieldName _
				  )
	dim	v
	if strFieldName = "" then
		v = 0
	else
		v = objRs.Fields(strFieldName)
		if isnull(v) then
			v = 0
		end if
	end if
	GetFieldValue = v
End Function
'-------------------------------------------------------------
'数字を返す
'-------------------------------------------------------------
Function GetNumber(byVal v,byVal strFormat)
	dim	strNumber

	strNumber = ""
	if isnull(v) = False then
		select case strFormat
		case "%"
			if CLng(v) <> 0 then
				strNumber = formatnumber(v,0,,,-1) & "%"
			end if
		case else
			if CLng(v) <> 0 then
				strNumber = formatnumber(v,0,,,-1)
			end if
		end select
	end if
	GetNumber = strNumber
End Function

Function GetFromA(byVal strYM,byVal strCenterCD)
	GetFromA = " "
	GetFromA = GetFromA & vbcrlf & "from (select"
	GetFromA = GetFromA & vbcrlf & " sum(if(YM = '" & strYM & "' and KamokuCD like 'A%',Prev,0)) Prev_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD like 'B%',Prev,0)) Prev_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD in ('C0100','C0200','C0300','C0400'),Prev,0)) Prev_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0500',Prev,0)) Prev_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0600',Prev,0)) Prev_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C9999',Prev,0)) Prev_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'D0100',Prev,0)) Prev_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'X0200',Prev,0)) Prev_X2"

	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD like 'A%',Plan,0)) Plan_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD like 'B%',Plan,0)) Plan_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD in ('C0100','C0200','C0300','C0400'),Plan,0)) Plan_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0500',Plan,0)) Plan_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0600',Plan,0)) Plan_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C9999',Plan,0)) Plan_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'D0100',Plan,0)) Plan_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'X0200',Plan,0)) Plan_X2"

	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD like 'A%',Result,0)) Result_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD like 'B%',Result,0)) Result_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD in ('C0100','C0200','C0300','C0400'),Result,0)) Result_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0500',Result,0)) Result_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C0600',Result,0)) Result_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'C9999',Result,0)) Result_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'D0100',Result,0)) Result_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(YM = '" & strYM & "' and KamokuCD = 'X0200',Result,0)) Result_X2"

	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'A%',Prev,0)) tPrev_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'B%',Prev,0)) tPrev_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400'),Prev,0)) tPrev_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0500',Prev,0)) tPrev_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0600',Prev,0)) tPrev_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C9999',Prev,0)) tPrev_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'D0100',Prev,0)) tPrev_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'X0200',Prev,0)) tPrev_X2"

	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'A%',Plan,0)) tPlan_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'B%',Plan,0)) tPlan_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400'),Plan,0)) tPlan_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0500',Plan,0)) tPlan_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0600',Plan,0)) tPlan_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C9999',Plan,0)) tPlan_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'D0100',Plan,0)) tPlan_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'X0200',Plan,0)) tPlan_X2"

	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'A%',Result,0)) tResult_A"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD like 'B%',Result,0)) tResult_B"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD in ('C0100','C0200','C0300','C0400'),Result,0)) tResult_C"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0500',Result,0)) tResult_C5"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C0600',Result,0)) tResult_C6"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'C9999',Result,0)) tResult_C9"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'D0100',Result,0)) tResult_D1"
	GetFromA = GetFromA & vbcrlf & ",sum(if(KamokuCD = 'X0200',Result,0)) tResult_X2"

	GetFromA = GetFromA & vbcrlf & " from IrData"
	GetFromA = GetFromA & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
	GetFromA = GetFromA & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
	GetFromA = GetFromA & vbcrlf & ") a"
End Function

'-------------------------------------------------------------
'SQL where 条件
'-------------------------------------------------------------
Function SqlWhere(byVal strAnd, byVal strField, byVal strValue)
	SqlWhere = ""
	if strValue = "" then
		exit function
	end if
	dim	strCmp
	strCmp = "="
	if left(strValue,1) = "-" then
		strCmp = "<>"
		strValue = right(strValue,len(strValue)-1)
	end if
	if instr(strValue,"%") > 0 then
		if strCmp = "=" then
			strCmp = "like"
		else
			strCmp = "not like"
		end if
	elseif instr(strValue,",") > 0 then
		if strCmp = "=" then
			strCmp = "in"
		else
			strCmp = "not in"
		end if
		strValue = "('" & replace(strValue, ",", "','") & "')"
	end if
	SqlWhere = vbCrLf & strAnd & " " & strField & " " & strCmp & " " & strValue
End Function
'-------------------------------------------------------------
'検索SQL
'-------------------------------------------------------------
Function MakeSql(byVal strCenterCD,byval strJKubun,byVal strYM,byval strKubun)
	dim	strSql
	dim	i
	dim	strSqlAdd

	strSql		= ""
	strSqlAdd	= ""

	dim	strSyushiCD		  '12345678
	strSyushiCD = GetRequest("SyushiCD","")
	if left(strKubun,8) = "_Syushi_" then
		strSyushiCD = right(RTrim(strKubun),3)
	end if

	strKubun = rtrim(strKubun)
	if strKubun = "pList" then
		strSql = "select"
'YM 	CenterCD 	SyushiCD 	KamokuCD 	Plan 	Result 	Prev
		strSql = strSql & vbcrlf & " i.CenterCD ""センター"""
		strSql = strSql & vbcrlf & ",i.YM ""年月"""
'		strSql = strSql & vbcrlf & ",i.KamokuCD ""科目"""
		strSql = strSql & vbcrlf & ",i.KamokuCD + RTrim(' ' + ifnull(k.KamokuName,'')) ""科目"""
		strSql = strSql & vbcrlf & ",ifnull(j.JigyoKubunName,'') ""事業区分"""
		strSql = strSql & vbcrlf & ",i.SyushiCD ""収支"""
		strSql = strSql & vbcrlf & ",i.Result ""実績"""
		strSql = strSql & vbcrlf & ",i.Plan ""計画"""
		strSql = strSql & vbcrlf & ",i.Prev ""前年"""
		strSql = strSql & vbcrlf & " from IrData i"
		strSql = strSql & vbcrlf & " left outer join Kamoku k on (i.KamokuCD = k.KamokuCD)"
		strSql = strSql & vbcrlf & " left outer join JigyoKubun j on (i.CenterCD = j.CenterCD and i.SyushiCD = j.SyushiCD)"
		strSql = strSql & vbcrlf & " WHERE i.YM = '" & strYM &"'"
		strSql = strSql & vbcrlf & "   AND i.CenterCD = '" & strCenterCD &"'"
		strSql = strSql & vbcrlf & "   AND (i.Result <> 0 or i.Plan <> 0 or i.Prev <> 0)"
		strSql = strSql & SqlWhere("and", "i.SyushiCD", strSyushiCD)
'		if strSyushiCD <> "" then
'			strSql = strSql & vbcrlf & "   AND i.SyushiCD = '" & strSyushiCD & "'"
'		end if
		if left(strKubun,8) <> "_Syushi_" then
			if strJKubun = "_その他" then
				strSql = strSql & vbcrlf & "   AND i.SyushiCd not in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "')"
			elseif strJKubun <> "" then
				strSql = strSql & vbcrlf & "   AND i.SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strJKubun & "')"
			end if
		end if
'		if strSyushiCD <> "" then
'			strSql = strSql & vbcrlf & "   AND i.SyushiCd = '" & strSyushiCD & "'"
'		end if
		strSql = strSql & vbcrlf & " order by i.YM,i.CenterCD,i.KamokuCD,""事業区分"",i.SyushiCD"
	else
		if left(strKubun,3) = "JM_" then
			strKubun = right(strKubun,len(strKubun) - 3)
			strSql = "select"
			strSql = strSql & vbcrlf & " sum((Result * if(KamokuCD like 'A%',1,0))) R_Uri"
			strSql = strSql & vbcrlf & ",sum((Result * if(KamokuCD like 'A%',1,if(KamokuCD like 'B%',-1,0)))) R_Rieki"
			strSql = strSql & vbcrlf & ",sum((Result * if(KamokuCD in ('C0100','C0200','C0300','C0400'),1,if(KamokuCD in ('X0200'),-1,0)))) R_Choku"
			strSql = strSql & vbcrlf & ",sum((Result * if(KamokuCD in ('X0200'),1,0))) R_Kan"
			strSql = strSql & vbcrlf & ",sum((Plan   * if(KamokuCD like 'A%',1,0))) P_Uri"
			strSql = strSql & vbcrlf & ",sum((Plan   * if(KamokuCD like 'A%',1,if(KamokuCD like 'B%',-1,0)))) P_Rieki"
			strSql = strSql & vbcrlf & ",sum((Plan   * if(KamokuCD in ('C0100','C0200','C0300','C0400'),1,if(KamokuCD in ('X0200'),-1,0)))) P_Choku"
			strSql = strSql & vbcrlf & ",sum((Plan   * if(KamokuCD in ('X0200'),1,0))) P_Kan"
			strSql = strSql & vbcrlf & " from IrData"
			strSql = strSql & vbcrlf & " WHERE YM = '" & strYM &"'"
			strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
			strSql = strSql & vbcrlf & "   AND KamokuCD between 'A0000' and 'X9999'"
			if left(strKubun,8) <> "_Syushi_" then
				if strKubun = "_その他" then
					strSql = strSql & vbcrlf & "   AND SyushiCd not in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "')"
				else
					strSql = strSql & vbcrlf & "   AND SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strKubun & "')"
				end if
			end if
			strSql = strSql & SqlWhere("and", "SyushiCd", strSyushiCD)
'			if strSyushiCD <> "" then
'				strSql = strSql & vbcrlf & "   AND SyushiCd = '" & strSyushiCD & "'"
'			end if
		elseif left(strKubun,1) = "_" then
			strKubun = right(strKubun,len(strKubun) - 1)
			strSql = "select"
			strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Result * if(KamokuCD like 'A%',1, 0),0)) A_Uri"
			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result * if(KamokuCD like 'A%',1,-1),0)) A_Rieki"
			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan * if(KamokuCD like 'A%',1, 0),0)) A_UriSa"
			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan * if(KamokuCD like 'A%',1,-1),0)) A_RiekiSa"
			strSql = strSql & vbcrlf & ",sum((Result * if(KamokuCD like 'A%',1, 0))) T_Uri"
			strSql = strSql & vbcrlf & ",sum((Result * if(KamokuCD like 'A%',1,-1))) T_Rieki"
			strSql = strSql & vbcrlf & ",sum((Plan * if(KamokuCD like 'A%',1, 0))) T_UriSa"
			strSql = strSql & vbcrlf & ",sum((Plan * if(KamokuCD like 'A%',1,-1))) T_RiekiSa"
			strSql = strSql & vbcrlf & " from IrData"
			strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
			strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
			strSql = strSql & vbcrlf & "   AND KamokuCD between 'A0000' and 'D9999'"
			if left(strKubun,7) <> "Syushi_" then
				if strKubun = "_その他" then
					strSql = strSql & vbcrlf & "   AND SyushiCd not in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "')"
				elseif strKubun <> "" then
					strSql = strSql & vbcrlf & "   AND SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strKubun & "')"
				end if
			end if
			strSql = strSql & SqlWhere("and", "SyushiCd", strSyushiCD)
'			if strSyushiCD <> "" then
'				strSql = strSql & vbcrlf & "   AND SyushiCd = '" & strSyushiCD & "'"
'			end if
		elseif left(strKubun,2) = "YK" then
			dim	strYM2
			strYM2 = strYM
			select case GetRequest("ptype","pTable")	'strTableType
			case "pTableJKYearMonth"	' 年間事業区分別概況書
				strYM2 = clng(strYM) + 100
			end select
			select case strKubun
			case "YK売上"
				strSql = "select"
				strSql = strSql & vbcrlf & " (ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",ARst04"
				strSql = strSql & vbcrlf & ",ARst05"
				strSql = strSql & vbcrlf & ",ARst06"
				strSql = strSql & vbcrlf & ",ARst07"
				strSql = strSql & vbcrlf & ",ARst08"
				strSql = strSql & vbcrlf & ",ARst09"
				strSql = strSql & vbcrlf & ",ARst10"
				strSql = strSql & vbcrlf & ",ARst11"
				strSql = strSql & vbcrlf & ",ARst12"
				strSql = strSql & vbcrlf & ",ARst01"
				strSql = strSql & vbcrlf & ",ARst02"
				strSql = strSql & vbcrlf & ",ARst03"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",APln04"
				strSql = strSql & vbcrlf & ",APln05"
				strSql = strSql & vbcrlf & ",APln06"
				strSql = strSql & vbcrlf & ",APln07"
				strSql = strSql & vbcrlf & ",APln08"
				strSql = strSql & vbcrlf & ",APln09"
				strSql = strSql & vbcrlf & ",APln10"
				strSql = strSql & vbcrlf & ",APln11"
				strSql = strSql & vbcrlf & ",APln12"
				strSql = strSql & vbcrlf & ",APln01"
				strSql = strSql & vbcrlf & ",APln02"
				strSql = strSql & vbcrlf & ",APln03"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03)"
				strSql = strSql & vbcrlf & "-(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK売上原価"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",BRst04"
				strSql = strSql & vbcrlf & ",BRst05"
				strSql = strSql & vbcrlf & ",BRst06"
				strSql = strSql & vbcrlf & ",BRst07"
				strSql = strSql & vbcrlf & ",BRst08"
				strSql = strSql & vbcrlf & ",BRst09"
				strSql = strSql & vbcrlf & ",BRst10"
				strSql = strSql & vbcrlf & ",BRst11"
				strSql = strSql & vbcrlf & ",BRst12"
				strSql = strSql & vbcrlf & ",BRst01"
				strSql = strSql & vbcrlf & ",BRst02"
				strSql = strSql & vbcrlf & ",BRst03"
				strSql = strSql & vbcrlf & ",(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",BPln04"
				strSql = strSql & vbcrlf & ",BPln05"
				strSql = strSql & vbcrlf & ",BPln06"
				strSql = strSql & vbcrlf & ",BPln07"
				strSql = strSql & vbcrlf & ",BPln08"
				strSql = strSql & vbcrlf & ",BPln09"
				strSql = strSql & vbcrlf & ",BPln10"
				strSql = strSql & vbcrlf & ",BPln11"
				strSql = strSql & vbcrlf & ",BPln12"
				strSql = strSql & vbcrlf & ",BPln01"
				strSql = strSql & vbcrlf & ",BPln02"
				strSql = strSql & vbcrlf & ",BPln03"
				strSql = strSql & vbcrlf & ",(BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03)"
				strSql = strSql & vbcrlf & "-(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"

			case "YK直接人件費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan) * if(KamokuCD = 'X0200',-1,1),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev * if(KamokuCD = 'X0200',-1,1),0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'"," * if(KamokuCD = 'X0200',-1,1)",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'"," * if(KamokuCD = 'X0200',-1,1)",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(if(YM<='" & strYM & "',Result,Plan)-Plan) * if(KamokuCD = 'X0200',-1,1),0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",C1Rst04-X2Rst04"
				strSql = strSql & vbcrlf & ",C1Rst05-X2Rst05"
				strSql = strSql & vbcrlf & ",C1Rst06-X2Rst06"
				strSql = strSql & vbcrlf & ",C1Rst07-X2Rst07"
				strSql = strSql & vbcrlf & ",C1Rst08-X2Rst08"
				strSql = strSql & vbcrlf & ",C1Rst09-X2Rst09"
				strSql = strSql & vbcrlf & ",C1Rst10-X2Rst10"
				strSql = strSql & vbcrlf & ",C1Rst11-X2Rst11"
				strSql = strSql & vbcrlf & ",C1Rst12-X2Rst12"
				strSql = strSql & vbcrlf & ",C1Rst01-X2Rst01"
				strSql = strSql & vbcrlf & ",C1Rst02-X2Rst02"
				strSql = strSql & vbcrlf & ",C1Rst03-X2Rst03"
				strSql = strSql & vbcrlf & ",(C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",C1Pln04-X2Pln04"
				strSql = strSql & vbcrlf & ",C1Pln05-X2Pln05"
				strSql = strSql & vbcrlf & ",C1Pln06-X2Pln06"
				strSql = strSql & vbcrlf & ",C1Pln07-X2Pln07"
				strSql = strSql & vbcrlf & ",C1Pln08-X2Pln08"
				strSql = strSql & vbcrlf & ",C1Pln09-X2Pln09"
				strSql = strSql & vbcrlf & ",C1Pln10-X2Pln10"
				strSql = strSql & vbcrlf & ",C1Pln11-X2Pln11"
				strSql = strSql & vbcrlf & ",C1Pln12-X2Pln12"
				strSql = strSql & vbcrlf & ",C1Pln01-X2Pln01"
				strSql = strSql & vbcrlf & ",C1Pln02-X2Pln02"
				strSql = strSql & vbcrlf & ",C1Pln03-X2Pln03"
				strSql = strSql & vbcrlf & ",((C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03))"
				strSql = strSql & vbcrlf & "-((C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03)) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK比例費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan) * if(KamokuCD = 'X0200',-1,1),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev * if(KamokuCD = 'X0200',-1,1),0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'"," * if(KamokuCD = 'X0200',-1,1)",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'"," * if(KamokuCD = 'X0200',-1,1)",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(if(YM<='" & strYM & "',Result,Plan)-Plan) * if(KamokuCD = 'X0200',-1,1),0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03)"
				strSql = strSql & vbcrlf & "+(C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",BRst04+C1Rst04-X2Rst04"
				strSql = strSql & vbcrlf & ",BRst05+C1Rst05-X2Rst05"
				strSql = strSql & vbcrlf & ",BRst06+C1Rst06-X2Rst06"
				strSql = strSql & vbcrlf & ",BRst07+C1Rst07-X2Rst07"
				strSql = strSql & vbcrlf & ",BRst08+C1Rst08-X2Rst08"
				strSql = strSql & vbcrlf & ",BRst09+C1Rst09-X2Rst09"
				strSql = strSql & vbcrlf & ",BRst10+C1Rst10-X2Rst10"
				strSql = strSql & vbcrlf & ",BRst11+C1Rst11-X2Rst11"
				strSql = strSql & vbcrlf & ",BRst12+C1Rst12-X2Rst12"
				strSql = strSql & vbcrlf & ",BRst01+C1Rst01-X2Rst01"
				strSql = strSql & vbcrlf & ",BRst02+C1Rst02-X2Rst02"
				strSql = strSql & vbcrlf & ",BRst03+C1Rst03-X2Rst03"
				strSql = strSql & vbcrlf & ",(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03)"
				strSql = strSql & vbcrlf & "+(C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",BPln04+C1Pln04-X2Pln04"
				strSql = strSql & vbcrlf & ",BPln05+C1Pln05-X2Pln05"
				strSql = strSql & vbcrlf & ",BPln06+C1Pln06-X2Pln06"
				strSql = strSql & vbcrlf & ",BPln07+C1Pln07-X2Pln07"
				strSql = strSql & vbcrlf & ",BPln08+C1Pln08-X2Pln08"
				strSql = strSql & vbcrlf & ",BPln09+C1Pln09-X2Pln09"
				strSql = strSql & vbcrlf & ",BPln10+C1Pln10-X2Pln10"
				strSql = strSql & vbcrlf & ",BPln11+C1Pln11-X2Pln11"
				strSql = strSql & vbcrlf & ",BPln12+C1Pln12-X2Pln12"
				strSql = strSql & vbcrlf & ",BPln01+C1Pln01-X2Pln01"
				strSql = strSql & vbcrlf & ",BPln02+C1Pln02-X2Pln02"
				strSql = strSql & vbcrlf & ",BPln03+C1Pln03-X2Pln03"
				strSql = strSql & vbcrlf & ",((BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03)"
				strSql = strSql & vbcrlf & "-(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03))"
				strSql = strSql & vbcrlf & "+((C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03))"
				strSql = strSql & vbcrlf & "-((C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03)) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK限界利益"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM<='" & strYM & "',Result,Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(Prev * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan",""," * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan",""," * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,"y")
				strSql = strSql & vbcrlf & ",sum((if(YM<='" & strYM & "',Result,Plan)-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03)"
				strSql = strSql & vbcrlf & "-((BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03)"
				strSql = strSql & vbcrlf & "+(C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03)) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",ARst04-(BRst04+C1Rst04-X2Rst04)"
				strSql = strSql & vbcrlf & ",ARst05-(BRst05+C1Rst05-X2Rst05)"
				strSql = strSql & vbcrlf & ",ARst06-(BRst06+C1Rst06-X2Rst06)"
				strSql = strSql & vbcrlf & ",ARst07-(BRst07+C1Rst07-X2Rst07)"
				strSql = strSql & vbcrlf & ",ARst08-(BRst08+C1Rst08-X2Rst08)"
				strSql = strSql & vbcrlf & ",ARst09-(BRst09+C1Rst09-X2Rst09)"
				strSql = strSql & vbcrlf & ",ARst10-(BRst10+C1Rst10-X2Rst10)"
				strSql = strSql & vbcrlf & ",ARst11-(BRst11+C1Rst11-X2Rst11)"
				strSql = strSql & vbcrlf & ",ARst12-(BRst12+C1Rst12-X2Rst12)"
				strSql = strSql & vbcrlf & ",ARst01-(BRst01+C1Rst01-X2Rst01)"
				strSql = strSql & vbcrlf & ",ARst02-(BRst02+C1Rst02-X2Rst02)"
				strSql = strSql & vbcrlf & ",ARst03-(BRst03+C1Rst03-X2Rst03)"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03)"
				strSql = strSql & vbcrlf & "-((BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03)"
				strSql = strSql & vbcrlf & "+(C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03)) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",APln04-(BPln04+C1Pln04-X2Pln04)"
				strSql = strSql & vbcrlf & ",APln05-(BPln05+C1Pln05-X2Pln05)"
				strSql = strSql & vbcrlf & ",APln06-(BPln06+C1Pln06-X2Pln06)"
				strSql = strSql & vbcrlf & ",APln07-(BPln07+C1Pln07-X2Pln07)"
				strSql = strSql & vbcrlf & ",APln08-(BPln08+C1Pln08-X2Pln08)"
				strSql = strSql & vbcrlf & ",APln09-(BPln09+C1Pln09-X2Pln09)"
				strSql = strSql & vbcrlf & ",APln10-(BPln10+C1Pln10-X2Pln10)"
				strSql = strSql & vbcrlf & ",APln11-(BPln11+C1Pln11-X2Pln11)"
				strSql = strSql & vbcrlf & ",APln12-(BPln12+C1Pln12-X2Pln12)"
				strSql = strSql & vbcrlf & ",APln01-(BPln01+C1Pln01-X2Pln01)"
				strSql = strSql & vbcrlf & ",APln02-(BPln02+C1Pln02-X2Pln02)"
				strSql = strSql & vbcrlf & ",APln03-(BPln03+C1Pln03-X2Pln03)"
				strSql = strSql & vbcrlf & ",((ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03)"
				strSql = strSql & vbcrlf & "-(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03))"
				strSql = strSql & vbcrlf & "-(((BRst04+BRst05+BRst06+BRst07+BRst08+BRst09+BRst10+BRst11+BRst12+BRst01+BRst02+BRst03)"
				strSql = strSql & vbcrlf & "-(BPln04+BPln05+BPln06+BPln07+BPln08+BPln09+BPln10+BPln11+BPln12+BPln01+BPln02+BPln03))"
				strSql = strSql & vbcrlf & "+((C1Rst04+C1Rst05+C1Rst06+C1Rst07+C1Rst08+C1Rst09+C1Rst10+C1Rst11+C1Rst12+C1Rst01+C1Rst02+C1Rst03)"
				strSql = strSql & vbcrlf & "-(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03))"
				strSql = strSql & vbcrlf & "-((C1Pln04+C1Pln05+C1Pln06+C1Pln07+C1Pln08+C1Pln09+C1Pln10+C1Pln11+C1Pln12+C1Pln01+C1Pln02+C1Pln03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03))) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK間接人件費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'X0200'"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",X2Rst04"
				strSql = strSql & vbcrlf & ",X2Rst05"
				strSql = strSql & vbcrlf & ",X2Rst06"
				strSql = strSql & vbcrlf & ",X2Rst07"
				strSql = strSql & vbcrlf & ",X2Rst08"
				strSql = strSql & vbcrlf & ",X2Rst09"
				strSql = strSql & vbcrlf & ",X2Rst10"
				strSql = strSql & vbcrlf & ",X2Rst11"
				strSql = strSql & vbcrlf & ",X2Rst12"
				strSql = strSql & vbcrlf & ",X2Rst01"
				strSql = strSql & vbcrlf & ",X2Rst02"
				strSql = strSql & vbcrlf & ",X2Rst03"
				strSql = strSql & vbcrlf & ",(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",X2Pln04"
				strSql = strSql & vbcrlf & ",X2Pln05"
				strSql = strSql & vbcrlf & ",X2Pln06"
				strSql = strSql & vbcrlf & ",X2Pln07"
				strSql = strSql & vbcrlf & ",X2Pln08"
				strSql = strSql & vbcrlf & ",X2Pln09"
				strSql = strSql & vbcrlf & ",X2Pln10"
				strSql = strSql & vbcrlf & ",X2Pln11"
				strSql = strSql & vbcrlf & ",X2Pln12"
				strSql = strSql & vbcrlf & ",X2Pln01"
				strSql = strSql & vbcrlf & ",X2Pln02"
				strSql = strSql & vbcrlf & ",X2Pln03"
				strSql = strSql & vbcrlf & ",(X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03)"
				strSql = strSql & vbcrlf & "-(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK通常管理費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0500','C0600')"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (C2Rst04+C2Rst05+C2Rst06+C2Rst07+C2Rst08+C2Rst09+C2Rst10+C2Rst11+C2Rst12+C2Rst01+C2Rst02+C2Rst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",C2Rst04"
				strSql = strSql & vbcrlf & ",C2Rst05"
				strSql = strSql & vbcrlf & ",C2Rst06"
				strSql = strSql & vbcrlf & ",C2Rst07"
				strSql = strSql & vbcrlf & ",C2Rst08"
				strSql = strSql & vbcrlf & ",C2Rst09"
				strSql = strSql & vbcrlf & ",C2Rst10"
				strSql = strSql & vbcrlf & ",C2Rst11"
				strSql = strSql & vbcrlf & ",C2Rst12"
				strSql = strSql & vbcrlf & ",C2Rst01"
				strSql = strSql & vbcrlf & ",C2Rst02"
				strSql = strSql & vbcrlf & ",C2Rst03"
				strSql = strSql & vbcrlf & ",(C2Pln04+C2Pln05+C2Pln06+C2Pln07+C2Pln08+C2Pln09+C2Pln10+C2Pln11+C2Pln12+C2Pln01+C2Pln02+C2Pln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",C2Pln04"
				strSql = strSql & vbcrlf & ",C2Pln05"
				strSql = strSql & vbcrlf & ",C2Pln06"
				strSql = strSql & vbcrlf & ",C2Pln07"
				strSql = strSql & vbcrlf & ",C2Pln08"
				strSql = strSql & vbcrlf & ",C2Pln09"
				strSql = strSql & vbcrlf & ",C2Pln10"
				strSql = strSql & vbcrlf & ",C2Pln11"
				strSql = strSql & vbcrlf & ",C2Pln12"
				strSql = strSql & vbcrlf & ",C2Pln01"
				strSql = strSql & vbcrlf & ",C2Pln02"
				strSql = strSql & vbcrlf & ",C2Pln03"
				strSql = strSql & vbcrlf & ",(C2Rst04+C2Rst05+C2Rst06+C2Rst07+C2Rst08+C2Rst09+C2Rst10+C2Rst11+C2Rst12+C2Rst01+C2Rst02+C2Rst03)"
				strSql = strSql & vbcrlf & "-(C2Pln04+C2Pln05+C2Pln06+C2Pln07+C2Pln08+C2Pln09+C2Pln10+C2Pln11+C2Pln12+C2Pln01+C2Pln02+C2Pln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK特別管理費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'C9999'"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (C9Rst04+C9Rst05+C9Rst06+C9Rst07+C9Rst08+C9Rst09+C9Rst10+C9Rst11+C9Rst12+C9Rst01+C9Rst02+C9Rst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",C9Rst04"
				strSql = strSql & vbcrlf & ",C9Rst05"
				strSql = strSql & vbcrlf & ",C9Rst06"
				strSql = strSql & vbcrlf & ",C9Rst07"
				strSql = strSql & vbcrlf & ",C9Rst08"
				strSql = strSql & vbcrlf & ",C9Rst09"
				strSql = strSql & vbcrlf & ",C9Rst10"
				strSql = strSql & vbcrlf & ",C9Rst11"
				strSql = strSql & vbcrlf & ",C9Rst12"
				strSql = strSql & vbcrlf & ",C9Rst01"
				strSql = strSql & vbcrlf & ",C9Rst02"
				strSql = strSql & vbcrlf & ",C9Rst03"
				strSql = strSql & vbcrlf & ",(C9Pln04+C9Pln05+C9Pln06+C9Pln07+C9Pln08+C9Pln09+C9Pln10+C9Pln11+C9Pln12+C9Pln01+C9Pln02+C9Pln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",C9Pln04"
				strSql = strSql & vbcrlf & ",C9Pln05"
				strSql = strSql & vbcrlf & ",C9Pln06"
				strSql = strSql & vbcrlf & ",C9Pln07"
				strSql = strSql & vbcrlf & ",C9Pln08"
				strSql = strSql & vbcrlf & ",C9Pln09"
				strSql = strSql & vbcrlf & ",C9Pln10"
				strSql = strSql & vbcrlf & ",C9Pln11"
				strSql = strSql & vbcrlf & ",C9Pln12"
				strSql = strSql & vbcrlf & ",C9Pln01"
				strSql = strSql & vbcrlf & ",C9Pln02"
				strSql = strSql & vbcrlf & ",C9Pln03"
				strSql = strSql & vbcrlf & ",(C9Rst04+C9Rst05+C9Rst06+C9Rst07+C9Rst08+C9Rst09+C9Rst10+C9Rst11+C9Rst12+C9Rst01+C9Rst02+C9Rst03)"
				strSql = strSql & vbcrlf & "-(C9Pln04+C9Pln05+C9Pln06+C9Pln07+C9Pln08+C9Pln09+C9Pln10+C9Pln11+C9Pln12+C9Pln01+C9Pln02+C9Pln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YKシステム費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "       )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (DRst04+DRst05+DRst06+DRst07+DRst08+DRst09+DRst10+DRst11+DRst12+DRst01+DRst02+DRst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",DRst04"
				strSql = strSql & vbcrlf & ",DRst05"
				strSql = strSql & vbcrlf & ",DRst06"
				strSql = strSql & vbcrlf & ",DRst07"
				strSql = strSql & vbcrlf & ",DRst08"
				strSql = strSql & vbcrlf & ",DRst09"
				strSql = strSql & vbcrlf & ",DRst10"
				strSql = strSql & vbcrlf & ",DRst11"
				strSql = strSql & vbcrlf & ",DRst12"
				strSql = strSql & vbcrlf & ",DRst01"
				strSql = strSql & vbcrlf & ",DRst02"
				strSql = strSql & vbcrlf & ",DRst03"
				strSql = strSql & vbcrlf & ",(DPln04+DPln05+DPln06+DPln07+DPln08+DPln09+DPln10+DPln11+DPln12+DPln01+DPln02+DPln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",DPln04"
				strSql = strSql & vbcrlf & ",DPln05"
				strSql = strSql & vbcrlf & ",DPln06"
				strSql = strSql & vbcrlf & ",DPln07"
				strSql = strSql & vbcrlf & ",DPln08"
				strSql = strSql & vbcrlf & ",DPln09"
				strSql = strSql & vbcrlf & ",DPln10"
				strSql = strSql & vbcrlf & ",DPln11"
				strSql = strSql & vbcrlf & ",DPln12"
				strSql = strSql & vbcrlf & ",DPln01"
				strSql = strSql & vbcrlf & ",DPln02"
				strSql = strSql & vbcrlf & ",DPln03"
				strSql = strSql & vbcrlf & ",(DRst04+DRst05+DRst06+DRst07+DRst08+DRst09+DRst10+DRst11+DRst12+DRst01+DRst02+DRst03)"
				strSql = strSql & vbcrlf & "-(DPln04+DPln05+DPln06+DPln07+DPln08+DPln09+DPln10+DPln11+DPln12+DPln01+DPln02+DPln03) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK固定費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD not like 'A%'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD not like 'A%',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD not like 'A%'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'X0200'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0500','C0600')"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'C9999'"
				strSql = strSql & vbcrlf & "     OR KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "     )"
				'----2017.11.14
				strSql = "select"
				strSql = strSql & vbcrlf & " (X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03)"
				strSql = strSql & vbcrlf & "+(C2Rst04+C2Rst05+C2Rst06+C2Rst07+C2Rst08+C2Rst09+C2Rst10+C2Rst11+C2Rst12+C2Rst01+C2Rst02+C2Rst03)"
				strSql = strSql & vbcrlf & "+(C9Rst04+C9Rst05+C9Rst06+C9Rst07+C9Rst08+C9Rst09+C9Rst10+C9Rst11+C9Rst12+C9Rst01+C9Rst02+C9Rst03)"
				strSql = strSql & vbcrlf & "+(DRst04+DRst05+DRst06+DRst07+DRst08+DRst09+DRst10+DRst11+DRst12+DRst01+DRst02+DRst03) Y_Prev"
				strSql = strSql & vbcrlf & ",(ARst04+ARst05+ARst06+ARst07+ARst08+ARst09+ARst10+ARst11+ARst12+ARst01+ARst02+ARst03) Y_Prev_Hi"
				strSql = strSql & vbcrlf & ",X2Rst04+C2Rst04+C9Rst04+DRst04"
				strSql = strSql & vbcrlf & ",X2Rst05+C2Rst05+C9Rst05+DRst05"
				strSql = strSql & vbcrlf & ",X2Rst06+C2Rst06+C9Rst06+DRst06"
				strSql = strSql & vbcrlf & ",X2Rst07+C2Rst07+C9Rst07+DRst07"
				strSql = strSql & vbcrlf & ",X2Rst08+C2Rst08+C9Rst08+DRst08"
				strSql = strSql & vbcrlf & ",X2Rst09+C2Rst09+C9Rst09+DRst09"
				strSql = strSql & vbcrlf & ",X2Rst10+C2Rst10+C9Rst10+DRst10"
				strSql = strSql & vbcrlf & ",X2Rst11+C2Rst11+C9Rst11+DRst11"
				strSql = strSql & vbcrlf & ",X2Rst12+C2Rst12+C9Rst12+DRst12"
				strSql = strSql & vbcrlf & ",X2Rst01+C2Rst01+C9Rst01+DRst01"
				strSql = strSql & vbcrlf & ",X2Rst02+C2Rst02+C9Rst02+DRst02"
				strSql = strSql & vbcrlf & ",X2Rst03+C2Rst03+C9Rst03+DRst03"
				strSql = strSql & vbcrlf & ",(X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03)"
				strSql = strSql & vbcrlf & "+(C2Pln04+C2Pln05+C2Pln06+C2Pln07+C2Pln08+C2Pln09+C2Pln10+C2Pln11+C2Pln12+C2Pln01+C2Pln02+C2Pln03)"
				strSql = strSql & vbcrlf & "+(C9Pln04+C9Pln05+C9Pln06+C9Pln07+C9Pln08+C9Pln09+C9Pln10+C9Pln11+C9Pln12+C9Pln01+C9Pln02+C9Pln03)"
				strSql = strSql & vbcrlf & "+(DPln04+DPln05+DPln06+DPln07+DPln08+DPln09+DPln10+DPln11+DPln12+DPln01+DPln02+DPln03) Y_Plan"
				strSql = strSql & vbcrlf & ",(APln04+APln05+APln06+APln07+APln08+APln09+APln10+APln11+APln12+APln01+APln02+APln03) Y_Plan_Hi"
				strSql = strSql & vbcrlf & ",X2Pln04+C2Pln04+C9Pln04+DPln04"
				strSql = strSql & vbcrlf & ",X2Pln05+C2Pln05+C9Pln05+DPln05"
				strSql = strSql & vbcrlf & ",X2Pln06+C2Pln06+C9Pln06+DPln06"
				strSql = strSql & vbcrlf & ",X2Pln07+C2Pln07+C9Pln07+DPln07"
				strSql = strSql & vbcrlf & ",X2Pln08+C2Pln08+C9Pln08+DPln08"
				strSql = strSql & vbcrlf & ",X2Pln09+C2Pln09+C9Pln09+DPln09"
				strSql = strSql & vbcrlf & ",X2Pln10+C2Pln10+C9Pln10+DPln10"
				strSql = strSql & vbcrlf & ",X2Pln11+C2Pln11+C9Pln11+DPln11"
				strSql = strSql & vbcrlf & ",X2Pln12+C2Pln12+C9Pln12+DPln12"
				strSql = strSql & vbcrlf & ",X2Pln01+C2Pln01+C9Pln01+DPln01"
				strSql = strSql & vbcrlf & ",X2Pln02+C2Pln02+C9Pln02+DPln02"
				strSql = strSql & vbcrlf & ",X2Pln03+C2Pln03+C9Pln03+DPln03"
				strSql = strSql & vbcrlf & " ((X2Rst04+X2Rst05+X2Rst06+X2Rst07+X2Rst08+X2Rst09+X2Rst10+X2Rst11+X2Rst12+X2Rst01+X2Rst02+X2Rst03)"
				strSql = strSql & vbcrlf & "+(C2Rst04+C2Rst05+C2Rst06+C2Rst07+C2Rst08+C2Rst09+C2Rst10+C2Rst11+C2Rst12+C2Rst01+C2Rst02+C2Rst03)"
				strSql = strSql & vbcrlf & "+(C9Rst04+C9Rst05+C9Rst06+C9Rst07+C9Rst08+C9Rst09+C9Rst10+C9Rst11+C9Rst12+C9Rst01+C9Rst02+C9Rst03)"
				strSql = strSql & vbcrlf & "+(DRst04+DRst05+DRst06+DRst07+DRst08+DRst09+DRst10+DRst11+DRst12+DRst01+DRst02+DRst03))"
				strSql = strSql & vbcrlf & "-((X2Pln04+X2Pln05+X2Pln06+X2Pln07+X2Pln08+X2Pln09+X2Pln10+X2Pln11+X2Pln12+X2Pln01+X2Pln02+X2Pln03)"
				strSql = strSql & vbcrlf & "+(C2Pln04+C2Pln05+C2Pln06+C2Pln07+C2Pln08+C2Pln09+C2Pln10+C2Pln11+C2Pln12+C2Pln01+C2Pln02+C2Pln03)"
				strSql = strSql & vbcrlf & "+(C9Pln04+C9Pln05+C9Pln06+C9Pln07+C9Pln08+C9Pln09+C9Pln10+C9Pln11+C9Pln12+C9Pln01+C9Pln02+C9Pln03)"
				strSql = strSql & vbcrlf & "+(DPln04+DPln05+DPln06+DPln07+DPln08+DPln09+DPln10+DPln11+DPln12+DPln01+DPln02+DPln03)) Y_Margin"
				strSql = strSql & vbcrlf & " from ( " & YKSub(strCenterCD,strYM,strYM2) & " ) i"
			case "YK営業利益"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM<='" & strYM & "',Result,Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(Prev * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan",""," * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan",""," * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,"y")
				strSql = strSql & vbcrlf & ",sum((if(YM<='" & strYM & "',Result,Plan)-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD between 'A0000' and 'D9999'"
			case "YK勤務時間","YK勤務時間(間)","YK勤務時間(計)"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(DT<='" & strYM & "',Result,Plan)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(DT<='" & strYM & "',Result,Plan)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(Prev",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Result,Plan","","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(Plan) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(Plan) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(Plan",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Plan","","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(DT<='" & strYM & "',Result,Plan)-Plan) Y_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				select case strKubun
				case "YK勤務時間"
					strSql = strSql & vbcrlf & "   AND (KamokuCD = 'TM101')"
				case "YK勤務時間(間)"
					strSql = strSql & vbcrlf & "   AND (KamokuCD = 'TM102')"
				case "YK勤務時間(計)"
					strSql = strSql & vbcrlf & "   AND (KamokuCD in ('TM101','TM102'))"
				end select
			case "YK作業時間","YK作業時間(間)","YK作業時間(計)"
				dim	strTM101
				dim	strTM201
				select case strKubun
				case "YK作業時間"
					strTM101 = "'TM101'"
					strTM201 = "'TM201'"
				case "YK作業時間(間)"
					strTM101 = "'TM102'"
					strTM201 = "'TM202'"
				case "YK作業時間(計)"
					strTM101 = "'TM101','TM102'"
					strTM201 = "'TM201','TM202'"
				end select
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Result,Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in (" & strTM201 & "," & strTM101 & ")"
				strSql = strSql & vbcrlf & "       )"
			case "YK非作業時間","YK非作業時間(間)","YK非作業時間(計)"
				dim	strTM301
				select case strKubun
				case "YK非作業時間"
					strTM101 = "'TM101'"
					strTM301 = "'TM301'"
				case "YK非作業時間(間)"
					strTM101 = "'TM102'"
					strTM301 = "'TM302'"
				case "YK非作業時間(計)"
					strTM101 = "'TM101','TM102'"
					strTM301 = "'TM301','TM301'"
				end select
	
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),if(DT='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Result,Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in (" & strTM301 & "," & strTM101 & ")"
				strSql = strSql & vbcrlf & "       )"
			case "YK有給時間","YK有給時間(間)","YK有給時間(計)"
				dim	strTM401
				select case strKubun
				case "YK有給時間"
					strTM101 = "'TM101'"
					strTM401 = "'TM401'"
				case "YK有給時間(間)"
					strTM101 = "'TM102'"
					strTM401 = "'TM402'"
				case "YK有給時間(計)"
					strTM101 = "'TM101','TM102'"
					strTM401 = "'TM401','TM402'"
				end select
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Result,Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  in (" & strTM101 & "),Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("DT",",sum(if(KamokuCD not in (" & strTM101 & "),Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("DT","Plan"," and KamokuCD not in (" & strTM101 & ")","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not in (" & strTM101 & "),if(DT<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in (" & strTM401 & "," & strTM101 & ")"
				strSql = strSql & vbcrlf & "       )"
			case "YK売上工数"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> 'Y9999',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD <> 'Y9999',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD <> 'Y9999'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD <> 'Y9999',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD <> 'Y9999'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD = 'Y9999'"
	'			strSql = strSql & vbcrlf & "     or KamokuCD = 'Y0100'"
				strSql = strSql & vbcrlf & "       )"
			case "YK工数(余裕率除)"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> 'Y9999',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',if(YM<='" & strYM & "',Result,Plan),0)) Y_Prev_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD <> 'Y9999',Prev,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Result,Plan"," and KamokuCD <> 'Y9999'","",strYM,"x")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Plan,0)) Y_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Plan,0)) Y_Plan_Hi"
'				strSql = strSql & vbcrlf & GetYearMonthSF("YM",",sum(if(KamokuCD <> 'Y9999',Plan,0)",GetNendo(strYM,3),0)
				strSql = strSql & vbcrlf & GetYearMonthYK("YM","Plan"," and KamokuCD <> 'Y9999'","",strYM,"y")
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',if(YM<='" & strYM & "',Result,Plan)-Plan,0)) Y_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD = 'Y9999'"
	'			strSql = strSql & vbcrlf & "     or KamokuCD = 'Y0100'"
				strSql = strSql & vbcrlf & "       )"
			end select
		elseif left(strKubun,2) = "CK" then
			select case strKubun
			case "CK売上"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD like 'A%'"
			case "CK売上原価"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD like 'B%'"
			case "CK直接人件費"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result * if(KamokuCD = 'X0200',-1,1)",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
			case "CK比例費"
				strSql = "select" 
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result * if(KamokuCD = 'X0200',-1,1)",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
			case "CK限界利益"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
			case "CK間接人件費"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'X0200'"
			case "CK通常管理費"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD in ('C0500','C0600')"
			case "CK特別管理費"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'C9999'"
			case "CKシステム費"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'D0100'"
			case "CK固定費"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD = 'X0200'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0500','C0600')"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'C9999'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "     )"
			case "CK営業利益"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("YM","Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)",strYM,0)
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD between 'A0000' and 'D9999'"
			case "CK勤務時間"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM101'"
			case "CK勤務時間(間)"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM102'"
			case "CK作業時間"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM201'"
			case "CK作業時間(間)"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM202'"
			case "CK非作業時間"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM301'"
			case "CK非作業時間(間)"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM302'"
			case "CK有給時間"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM401'"
			case "CK有給時間(間)"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'TM402'"
			case "CK売上工数"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'Y9999'"
			case "CK工数(余裕率除)"
				strSql = "select"
				strSql = strSql & vbcrlf & GetYearMonthSF("DT","Result",strYM,0)
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & GetNendo(strYM,3) &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD = 'Y9999'"
			end select
		else
			select case strKubun
			case "JK売上"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(Result-Plan) T_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(Result-Prev) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD like 'A%'"
			case "売上"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum(Result-Plan) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","A%","")
			case "JK売上原価"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "       )"
			case "JK直接人件費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result        * if(KamokuCD = 'X0200',-1,1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan          * if(KamokuCD = 'X0200',-1,1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',(Result-Plan) * if(KamokuCD = 'X0200',-1,1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev          * if(KamokuCD = 'X0200',-1,1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',(Result-Prev) * if(KamokuCD = 'X0200',-1,1),0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result * if(KamokuCD = 'X0200',-1,1),0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(Result-Plan) * if(KamokuCD = 'X0200',-1,1),0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev * if(KamokuCD = 'X0200',-1,1),0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(Result-Prev) * if(KamokuCD = 'X0200',-1,1),0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
			case "JK比例費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result        * if(KamokuCD = 'X0200',-1,1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan          * if(KamokuCD = 'X0200',-1,1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',(Result-Plan) * if(KamokuCD = 'X0200',-1,1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev          * if(KamokuCD = 'X0200',-1,1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',(Result-Prev) * if(KamokuCD = 'X0200',-1,1),0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result * if(KamokuCD = 'X0200',-1,1),0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan * if(KamokuCD = 'X0200',-1,1),0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(Result-Plan) * if(KamokuCD = 'X0200',-1,1),0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev * if(KamokuCD = 'X0200',-1,1),0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',(Result-Prev) * if(KamokuCD = 'X0200',-1,1),0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
			case "JK限界利益"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan   * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev          * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Prev) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum((Result-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum((Result-Prev) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0100','C0200','C0300','C0400','X0200')"
				strSql = strSql & vbcrlf & "       )"
			case "JK間接人件費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'X0200'"
				strSql = strSql & vbcrlf & "       )"
			case "JK通常管理費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0500','C0600')"
				strSql = strSql & vbcrlf & "       )"
			case "JK特別管理費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'C9999'"
				strSql = strSql & vbcrlf & "       )"
			case "JKシステム費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "       )"
			case "JK固定費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD not like 'A%',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'X0200'"
				strSql = strSql & vbcrlf & "     or KamokuCD in ('C0500','C0600')"
				strSql = strSql & vbcrlf & "     or KamokuCD = 'C9999'"
				strSql = strSql & vbcrlf & "     OR KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "     )"
			case "JK営業利益"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan   * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev          * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Prev) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(Result * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum((Result-Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD     like 'A%',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum((Result-Prev) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND KamokuCD between 'A0000' and 'D9999'"
			case "JK勤務時間","JK勤務時間(間)","JK勤務時間(計)"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(DT = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(Result-Plan) T_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(Result-Prev) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				select case strKubun
				case "JK勤務時間"
					strSql = strSql & vbcrlf & "   AND (KamokuCD = 'TM101'"
				case "JK勤務時間(間)"
					strSql = strSql & vbcrlf & "   AND (KamokuCD = 'TM102'"
				case "JK勤務時間(計)"
					strSql = strSql & vbcrlf & "   AND (KamokuCD = 'TM100'"
				end select
				strSql = strSql & vbcrlf & "       )"
	'			strSql = "select"
	'			strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Result,0)) A_Result"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result_Hi"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Plan,0)) A_Margin"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
	'			strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
	'			strSql = strSql & vbcrlf & ",sum(Result) T_Result"
	'			strSql = strSql & vbcrlf & ",sum(Result) T_Result_Hi"
	'			strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
	'			strSql = strSql & vbcrlf & ",sum(Plan) T_Plan_Hi"
	'			strSql = strSql & vbcrlf & ",sum(Result-Plan) T_Margin"
	'			strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
	'			strSql = strSql & vbcrlf & ",sum(Prev) T_Prev_Hi"
	'			strSql = strSql & vbcrlf & ",sum(Result-Prev) T_Prev_Margin"
	'			strSql = strSql & vbcrlf & " from IrData"
	'			strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
	'			strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
	'			strSql = strSql & vbcrlf & "   AND (KamokuCD = 'Y0010'"
	'			strSql = strSql & vbcrlf & "       )"
			case "JK作業時間","JK作業時間(間)"
	'			dim	strTM101
	'			dim	strTM201
				select case strKubun
				case "JK作業時間"
					strTM101 = "TM101"
					strTM201 = "TM201"
				case "JK作業時間(間)"
					strTM101 = "TM102"
					strTM201 = "TM202"
				end select
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in ('" & strTM201 & "','" & strTM101 & "')"
				strSql = strSql & vbcrlf & "       )"
			case "JK非作業時間","JK非作業時間(間)"
	'			dim	strTM301
				select case strKubun
				case "JK非作業時間"
					strTM101 = "TM101"
					strTM301 = "TM301"
				case "JK非作業時間(間)"
					strTM101 = "TM102"
					strTM301 = "TM302"
				end select
	
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in ('" & strTM301 & "','" & strTM101 & "')"
				strSql = strSql & vbcrlf & "       )"
			case "JK有給時間","JK有給時間(間)"
	'			dim	strTM401
				select case strKubun
				case "JK有給時間"
					strTM101 = "TM101"
					strTM401 = "TM401"
				case "JK有給時間(間)"
					strTM101 = "TM102"
					strTM401 = "TM402"
				end select
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "' and DT = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "' and DT = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = '" & strTM101 & "',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> '" & strTM101 & "',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from Attendance"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD in ('" & strTM401 & "','" & strTM101 & "')"
				strSql = strSql & vbcrlf & "       )"
			case "JK時間内訳"
				strSql = "select"
				strSql = strSql & vbcrlf & " k.KamokuName"
				strSql = strSql & vbcrlf & ",PersonCD"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(0) A_Result_Hi_0"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(0) A_Plan_Hi_0"
				strSql = strSql & vbcrlf & ",sum(0) A_Margin_0"
				strSql = strSql & vbcrlf & ",sum(if(DT = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(0) A_Prev_Hi_0"
				strSql = strSql & vbcrlf & ",sum(0) A_Prev_Margin_0"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum(0) T_Result_Hi_0"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(0) T_Plan_Hi_0"
				strSql = strSql & vbcrlf & ",sum(0) T_Margin_0"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(0) T_Prev_Hi_0"
				strSql = strSql & vbcrlf & ",sum(0) T_Prev_Margin_0"
				strSql = strSql & vbcrlf & " from Attendance a"
				strSql = strSql & vbcrlf & " inner join Kamoku k on (a.KamokuCD = k.KamokuCD)"
				strSql = strSql & vbcrlf & " WHERE DT between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (a.KamokuCD like 'TM%1' or a.KamokuCD like 'TM%2')"

				strSqlAdd = strSqlAdd & vbcrlf & " group by"
				strSqlAdd = strSqlAdd & vbcrlf & " a.KamokuCD"
				strSqlAdd = strSqlAdd & vbcrlf & ",k.KamokuName"
				strSqlAdd = strSqlAdd & vbcrlf & ",a.PersonCD"
				strSqlAdd = strSqlAdd & vbcrlf & " having sum(Result) <> 0 or sum(Prev) <> 0 or sum(Plan) <> 0"
				strSqlAdd = strSqlAdd & vbcrlf & " order by"
				strSqlAdd = strSqlAdd & vbcrlf & " a.KamokuCD"
				strSqlAdd = strSqlAdd & vbcrlf & ",a.PersonCD"
			case "JK売上工数"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD = 'Y9999'"
	'			strSql = strSql & vbcrlf & "     or KamokuCD = 'Y0100'"
				strSql = strSql & vbcrlf & "       )"
			case "JK工数(余裕率除)"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Result,0)) A_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Plan,0)) A_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100' and YM = '" & strYM & "',Prev,0)) A_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999' and YM = '" & strYM & "',Result-Prev,0)) A_Prev_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result,0)) T_Result"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Result,0)) T_Result_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Plan,0)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Plan,0)) T_Plan_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result-Plan,0)) T_Margin"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Prev,0)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD  = 'Y0100',Prev,0)) T_Prev_Hi"
				strSql = strSql & vbcrlf & ",sum(if(KamokuCD <> 'Y9999',Result-Prev,0)) T_Prev_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & " WHERE YM between '" & GetNendo(strYM,4) & "' and '" & strYM &"'"
				strSql = strSql & vbcrlf & "   AND CenterCD = '" & strCenterCD &"'"
				strSql = strSql & vbcrlf & "   AND (KamokuCD = 'Y9999'"
	'			strSql = strSql & vbcrlf & "     or KamokuCD = 'Y0100'"
				strSql = strSql & vbcrlf & "       )"
			case "資材費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","B0100,B0200,B0500","")
			case "工料仕入"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","B0300","")
			case "その他仕入"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","B%","")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","-B0100,B0200,B0300,B0500","")
			case "直接人件費"
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_C - a.Prev_X2) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_C - a.Plan_X2) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_C - a.Result_X2) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_C - a.tPrev_X2) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_C - a.tPlan_X2) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_C - a.tResult_X2) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "比例費"	'合計
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_B + (a.Prev_C - a.Prev_X2)) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_B + (a.Plan_C - a.Plan_X2)) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_B + (a.Result_C - a.Result_X2)) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_B + (a.tPrev_C - a.tPrev_X2)) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_B + (a.tPlan_C - a.tPlan_X2)) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_B + (a.tResult_C - a.tResult_X2)) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "限界利益"	'売上－比例費
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_A - a.Prev_B - (a.Prev_C - a.Prev_X2)) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_A - a.Plan_B - (a.Plan_C - a.Plan_X2)) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_A - a.Result_B - (a.Result_C - a.Result_X2)) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_A - a.tPrev_B - (a.tPrev_C - a.tPrev_X2)) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_A - a.tPlan_B - (a.tPlan_C - a.tPlan_X2)) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_A - a.tResult_B - (a.tResult_C - a.tResult_X2)) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "間接人件費"
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_X2) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_X2) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_X2) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_X2) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_X2) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_X2) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "通常管理費"
'				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","C0500,C0600","")
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_C5 + a.Prev_C6) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_C5 + a.Plan_C6) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_C5 + a.Result_C6) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_C5 + a.tPrev_C6) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_C5 + a.tPlan_C6) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_C5 + a.tResult_C6) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "特別管理費"
'				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","C9999","")
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_C9) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_C9) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_C9) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_C9) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_C9) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_C9) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "システム費"
'				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","D0100","")
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_D1) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_D1) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_D1) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_D1) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_D1) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_D1) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "固定費"	'合計
				strSql = "select"
				strSql = strSql & vbcrlf & " (a.Prev_X2 + a.Prev_C5 + a.Prev_C9 + a.Prev_D1) A_Prev"
				strSql = strSql & vbcrlf & ",(a.Plan_X2 + a.Plan_C5 + a.Plan_C9 + a.Plan_D1) A_Plan"
				strSql = strSql & vbcrlf & ",(a.Result_X2 + a.Result_C5 + a.Result_C9 + a.Result_D1) A_Result"
				strSql = strSql & vbcrlf & ",0 A_Margin"
				strSql = strSql & vbcrlf & ",(a.tPrev_X2 + a.tPrev_C5 + a.tPrev_C9 + a.tPrev_D1) T_Prev"
				strSql = strSql & vbcrlf & ",(a.tPlan_X2 + a.tPlan_C5 + a.tPlan_C9 + a.tPlan_D1) T_Plan"
				strSql = strSql & vbcrlf & ",(a.tResult_X2 + a.tResult_C5 + a.tResult_C9 + a.tResult_D1) T_Result"
				strSql = strSql & vbcrlf & ",0 T_Margin"
				MakeSql = strSql & GetFromA(strYM,strCenterCD)
				exit function
			case "経費"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'C%'"
				strSql = strSql & vbcrlf & "     OR KamokuCD = 'D0100'"
				strSql = strSql & vbcrlf & "     )"
			case "仕入"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","B%","")
			case "粗利益"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev * if(KamokuCD like 'B%',-1,1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan * if(KamokuCD like 'B%',-1,1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result * if(KamokuCD like 'B%',-1,1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan) * if(KamokuCD like 'B%',-1,1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev * if(KamokuCD like 'B%',-1,1)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD like 'B%',-1,1)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result * if(KamokuCD like 'B%',-1,1)) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan) * if(KamokuCD like 'B%',-1,1)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & "   AND (KamokuCD like 'A%'"
				strSql = strSql & vbcrlf & "     OR KamokuCD like 'B%'"
				strSql = strSql & vbcrlf & "     )"
			case "営業利益"
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev * if(KamokuCD like 'A%',1,-1),0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan * if(KamokuCD like 'A%',1,-1),0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result * if(KamokuCD like 'A%',1,-1),0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',(Result-Plan) * if(KamokuCD like 'A%',1,-1),0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev * if(KamokuCD like 'A%',1,-1)) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan * if(KamokuCD like 'A%',1,-1)) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result * if(KamokuCD like 'A%',1,-1)) T_Result"
				strSql = strSql & vbcrlf & ",sum((Result-Plan) * if(KamokuCD like 'A%',1,-1)) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","A0000","D9999")
			case else
				strSql = "select"
				strSql = strSql & vbcrlf & " sum(if(YM = '" & strYM & "',Prev,0)) A_Prev"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Plan,0)) A_Plan"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result,0)) A_Result"
				strSql = strSql & vbcrlf & ",sum(if(YM = '" & strYM & "',Result-Plan,0)) A_Margin"
				strSql = strSql & vbcrlf & ",sum(Prev) T_Prev"
				strSql = strSql & vbcrlf & ",sum(Plan) T_Plan"
				strSql = strSql & vbcrlf & ",sum(Result) T_Result"
				strSql = strSql & vbcrlf & ",sum(Result-Plan) T_Margin"
				strSql = strSql & vbcrlf & " from IrData"
				strSql = strSql & vbcrlf & makeWhere("","YM",GetNendo(strYM,4),strYM)
				strSql = strSql & vbcrlf & makeWhere(" ","CenterCD",strCenterCD,"")
				strSql = strSql & vbcrlf & makeWhere(" ","KamokuCD","99999","")
			end select
		end if
		if right(strSql,3) <> ") i" then
			if strJKubun <> "" then
				strSql = strSql & vbcrlf & "   AND SyushiCd in (select SyushiCd from JigyoKubun where CenterCD = '" & strCenterCD & "' and JigyoKubunName = '" & strJKubun & "')"
			else
				strSql = strSql & vbcrlf & "   AND SyushiCd <> ''"
			end if
			strSql = strSql & SqlWhere("and", "SyushiCD", strSyushiCD)
'			if strSyushiCD <> "" then
'				strSql = strSql & vbcrlf & "   AND SyushiCd = '" & strSyushiCD & "'"
'			end if
		end if
	end if
	MakeSql = strSql & strSqlAdd
End Function
'-------------------------------------------------------------
'1年間のselect のfield分を返す
'-------------------------------------------------------------
Function GetYearMonthYK(byVal strFld,byVal strSum,byval strAnd,byval strMul,byVal strYM,byVal strY)
	select case GetRequest("ptype","")
	case "pTableJKYearMonth"
'		if strSum = "Plan" then
'			strYM = clng(strYM) + 100
'		end if
	case "pTableJKYearMonth2"
	case else
		GetYearMonthYK = ""
		exit function
	end select
	dim	strYearMonthYK
	strYearMonthYK = ""
	dim	i
	for	i = 1 to 12
		dim	m
		m = i + 3
		if m > 12 then
			m = m - 12
		end if
		dim	strYYYYMM
		strYYYYMM = GetNendo(strYM,m)
		dim	strF
		strF = strSum
		if strSum = "Result,Plan" then
			if CLng(strYYYYMM) <= CLng(strYM) then
				strF = split(strSum,",")(0)
			else
				strF = split(strSum,",")(1)
			end if
		end if
		dim	strSelect
		strSelect = ",sum(if(" & strFld & " = '" & strYYYYMM & "'" & strAnd & "," & strF & ",0)" & strMul & ")"
		strYearMonthYK = strYearMonthYK & strSelect & vbCrLf
	next
	GetYearMonthYK = strYearMonthYK
End Function

'-------------------------------------------------------------
'1年間のselect のfield分を返す
'-------------------------------------------------------------
Function GetYearMonthSF(byVal strFld,byVal strSum,byVal strYM,byVal iOffset)
	dim	strYearMonthSF
	strYearMonthSF = ""
	select case GetRequest("ptype","")
	case "pTableJKYearMonth","pTableJKYearMonth2"
		dim	lngYYYYMM
		lngYYYYMM = CLng(strYM)
		lngYYYYMM = lngYYYYMM + (iOffset * 100)
	
		dim	strY
		if inStr(strSum,"Prev") > 0 then
			strY = "x"
		else
			strY = "y"
		end if
		dim	i
		for	i = 1 to 12
			dim	m
			m = i + 3
			if m > 12 then
				m = m - 12
			end if
			dim	strYYYYMM
			strYYYYMM = GetNendo(CStr(lngYYYYMM),m)
			strYearMonthSF = strYearMonthSF & strSum & " * (if(" & strFld & "='" & strYYYYMM & "',1,0))) " & strY & strYYYYMM & vbCrLf
		next
	case "pTableChaku7","pTableChaku8","pTableChaku9","pTableChaku10","pTableChaku11","pTableChaku","pTableChaku1","pTableChaku2","pTableChaku3"
		dim	iKeikaku
		iKeikaku = 0
		dim	strM
		if right(GetRequest("ptype",""),2) = "11" then
			iKeikaku = 8
		else
			strM = right(GetRequest("ptype",""),1)
			if strM <> "u" then
				iKeikaku = CLng(strM)
			end if
			if iKeikaku > 4 then
				' 4月=1 4-3
				' 5月=2 5-3
				' 6月=3 6-3
				' 7月=4 7-3
				' 8月=5 8-3
				' 9月=6 9-3
				'10月=7 10-3
				'11月=8 11-3
				'12月=9 12-3
				' 1月=10 1+9
				' 2月=11 2+9
				' 3月=12 3+9
				iKeikaku = iKeikaku - 3
			else
				iKeikaku = iKeikaku + 9
			end if
		end if
		dim	strKeikakuYM
		strKeikakuYM = ""
		dim	strComma
		strComma = " "
		for	i = 1 to 12
			m = i + 3
			if m > 12 then
				m = m - 12
			end if
			strYYYYMM = GetNendo(strYM,m)
			if i >= iKeikaku then
				if strKeikakuYM = "" then
					strKeikakuYM = strYYYYMM
				end if
				strSum = Replace(strSum,"Result","Plan")
			end if
			strYearMonthSF = strYearMonthSF & strComma & "sum( Round(if(" & strFld & " = '" & strYYYYMM & "'," & strSum & ",0),0) ) A_Result_" & m & vbCrLf
			strComma = ","
		next
		if strSum = "Plan" then
			strSum = ""
		else
			strSum = Replace(strSum,"Plan","")
		end if
		strYearMonthSF = strYearMonthSF & ",sum( Round(if(" & strFld & " < '" & strKeikakuYM & "',Result,Plan)" & strSum & ",0) ) T_Result" & vbCrLf
	end select
	GetYearMonthSF = strYearMonthSF
End Function

'-------------------------------------------------------------
'1年間の<TH>タグを返す
'-------------------------------------------------------------
Function GetYearMonthTH(byVal strYM,byVal iOffset)
	dim	strYearMonthTH
	strYearMonthTH = ""
	select case GetRequest("ptype","")
	case "pTableJKYearMonth","pTableJKYearMonth2"
		dim	lngYYYYMM
		lngYYYYMM = CLng(strYM)
		lngYYYYMM = lngYYYYMM + (iOffset * 100)
	
		dim	i
		for	i = 1 to 12
			dim	m
			m = i + 3
			if m > 12 then
				m = m - 12
			end if
			dim	strYYYYMM
			strYYYYMM = GetNendo(CStr(lngYYYYMM),m)
			strYearMonthTH = strYearMonthTH & "<TH>" & strYYYYMM & "</TH>" & vbCrLf
		next
	case "pTableChaku7","pTableChaku8","pTableChaku9","pTableChaku10","pTableChaku11","pTableChaku","pTableChaku1","pTableChaku2","pTableChaku3"
		dim	iKeikaku
		iKeikaku = 0
		if right(GetRequest("ptype",""),2) = "11" then
			iKeikaku = 8
		else
			strM = right(GetRequest("ptype",""),1)
			if strM <> "u" then
				iKeikaku = CLng(strM)
			end if
			if iKeikaku > 4 then
				' 4月=1 4-3
				' 5月=2 5-3
				' 6月=3 6-3
				' 7月=4 7-3
				' 8月=5 8-3
				' 9月=6 9-3
				'10月=7 10-3
				'11月=8 11-3
				'12月=9 12-3
				' 1月=10 1+9
				' 2月=11 2+9
				' 3月=12 3+9
				iKeikaku = iKeikaku - 3
			else
				iKeikaku = iKeikaku + 9
			end if
		end if
		for	i = 1 to 12
			m = i + 3
			if m > 12 then
				m = m - 12
			end if
			dim	strM
			strM = m & "月"
			if i >= iKeikaku then
				strM = strM & "<br>計画"
			end if
			strYearMonthTH = strYearMonthTH & "<TH>" & strM & "</TH>" & vbCrLf
		next
	end select
	GetYearMonthTH = strYearMonthTH
End Function

'-------------------------------------------------------------
'テーブルヘッダー
'-------------------------------------------------------------
Function MakeHeader(byVal objDb,byVal strCenterCD,byVal strYM,byval strTableType)
	dim	strHeader
	dim	objRs
	dim	strSyushiCD
	dim	strSyushiName
	dim	i
	strHeader = vbCrLf
	strHeader = strHeader & "<!-- MakeHear(" & strCenterCD & "," & strYM & "," & strTableType & ")-->" & vbCrLf

	select case strTableType
	case "pList"
		strHeader = strHeader & GetTdList(objDb,strCenterCD,"Header",strYM,strTableType)
	case "pTableJKYear","pTableJKYearMonth","pTableJKYearMonth2"
		dim	strPeriod1
		dim	strPeriod2
		dim	intYM1
		dim	intYM2
		if strTableType = "pTableJKYearMonth2" then
			intYM1 = 0	' 今期見通し
			intYM2 = 0	' 今期計画
		else
			intYM1 = 0	' 今期見通し
			intYM2 = 1	' 来期計画
		end if
		strPeriod1 = GetPeriod(strYM) + intYM1
		strPeriod2 = GetPeriod(strYM) + intYM2

		strHeader = strHeader & "<TR>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"" rowspan=""1"">" & strYM & "</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"" rowspan=""1"">" & strPeriod1 & "期見通し</TH>" & vbCrLf
		strHeader = strHeader & GetYearMonthTH(strYM,intYM1)
		strHeader = strHeader & "<TH colspan=""2"" rowspan=""1"">" & strPeriod2 & "期計画</TH>" & vbCrLf
		strHeader = strHeader & GetYearMonthTH(strYM,intYM2)
		strHeader = strHeader & "<TH colspan=""1"" rowspan=""1"">差</TH>" & vbCrLf
		strHeader = strHeader & "</TR>" & vbCrLf
	case "pTable"
		strHeader = strHeader & "<TR>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"" rowspan=""2""></TH>" & vbCrLf
		strHeader = strHeader & "<TH>" & GetPeriod(strYM) - 1 & "期</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""3"" title=""" & strYM & """>" & GetPeriod(strYM) & "期</TH>" & vbCrLf
		strHeader = strHeader & "<TH>" & GetPeriod(strYM) - 1 & "期累計</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""3"">" & GetPeriod(strYM) & "期累計</TH>" & vbCrLf
		strHeader = strHeader & "</TR>" & vbCrLf
		strHeader = strHeader & "<TR>" & vbCrLf
		strHeader = strHeader & "<TH>実績</TH>" & vbCrLf
		strHeader = strHeader & "<TH>計画</TH>" & vbCrLf
		strHeader = strHeader & "<TH>実績</TH>" & vbCrLf
		strHeader = strHeader & "<TH>差</TH>" & vbCrLf
		strHeader = strHeader & "<TH>実績</TH>" & vbCrLf
		strHeader = strHeader & "<TH>計画</TH>" & vbCrLf
		strHeader = strHeader & "<TH>実績</TH>" & vbCrLf
		strHeader = strHeader & "<TH>差</TH>" & vbCrLf
		strHeader = strHeader & "</TR>" & vbCrLf
	case "pTableJK","pTableJKKan"
		strHeader = strHeader & "<TR>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2""></TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">当月実績</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">事業計画</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""1"">差</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">前年実績</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""1"">差</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">累計実績</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">事業計画</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""1"">差</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2"">前年実績</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""1"">差</TH>" & vbCrLf
		strHeader = strHeader & "</TR>" & vbCrLf
	case "pTableChaku7","pTableChaku8","pTableChaku9","pTableChaku10","pTableChaku11","pTableChaku","pTableChaku1","pTableChaku2","pTableChaku3"
		strHeader = strHeader & "<TR>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""2""></TH>" & vbCrLf
		strHeader = strHeader & GetYearMonthTH(strYM,0)
'		strHeader = strHeader & "<TH colspan=""1"">4月</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">5月</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">6月</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">7月</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">8月</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">9月</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">10月</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">11月</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">12月<br>計画</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">1月<br>計画</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">2月<br>計画</TH>" & vbCrLf
'		strHeader = strHeader & "<TH colspan=""1"">3月<br>計画</TH>" & vbCrLf
		strHeader = strHeader & "<TH colspan=""1"">合計</TH>" & vbCrLf
		strHeader = strHeader & "</TR>" & vbCrLf
	end select

	strHeader = strHeader & "<!-- MakeHear() End -->" & vbCrLf

	MakeHeader = strHeader
End Function
'-------------------------------------------------------------
'エラーメッセージHTML
'-------------------------------------------------------------
Function MakeError(byVal objErr)
	dim	strHTML
	strHTML = strHTML & "<tr><td>Err.Number:</td>"
	strHTML = strHTML & "<td>0x" & Hex(objErr.Number) & "(" & objErr.Number & ")</td></tr>"
	strHTML = strHTML & "<tr><td>Err.Description:</td>"
	strHTML = strHTML & "<td>" & objErr.Description & "</td></tr>"
	strHTML = strHTML & "<tr><td>Err.Source:</td>"
	strHTML = strHTML & "<td>" & objErr.Source & "</td></tr>"
	MakeError = strHTML
End Function
'-------------------------------------------------------------
'収支マスター検索SQL
'-------------------------------------------------------------
Function SyushiSql(byVal strCenterCD)
	dim	strSql

	strSql = "select"
	strSql = strSql & " s.CenterCD CenterCD"
	strSql = strSql & ",s.SyushiKB"
	strSql = strSql & ",sk.SyushiKBName"
	strSql = strSql & ",s.SyushiCD SyushiCD"
	strSql = strSql & ",s.SyushiName SyushiName"
	strSql = strSql & " FROM Syushi s"
	strSql = strSql & " left outer join SyushiKB sk on (s.SyushiKB = sk.SyushiKB)"
	strSql = strSql & " WHERE s.CenterCD = '" & strCenterCD &"'"
	strSql = strSql & " ORDER BY"
	strSql = strSql & " s.CenterCD"
	strSql = strSql & ",s.SyushiKB"
	strSql = strSql & ",s.SyushiCD"

	SyushiSql = strSql
End Function
'-------------------------------------------------------------
'着地用 select フィールド
'sum(if(YM <= '201211',Result,Plan) * if(KamokuCD = 'X0200',-1,1) * if(KamokuCD like 'A%',1,-1),0)) A_Total
'-------------------------------------------------------------
Function ChakuSelect(byVal strSql,byVal strSum1,byVal strSum2,byVal strFld)
	dim	i
	dim	sYM
	dim	strV
	for i = 1 to 12
		strSql = strSql & vbcrlf
		if i = 1 then
			strSql = strSql & " "
		else
			strSql = strSql & ","
		end if
		if i < 10 then
			sYM = "2012" & right("0" & i + 3,2)
		else
			sYM = "2013" & right("0" & i - 9,2)
		end if
		if sYM <= "201211" then
			strV = "Result"
		else
			strV = "Plan"
		end if
		strV = "if(" & strFld & " = '" & sYM & "'," & strV & ",0)"
		strSql = strSql & strSum1 & strV & strSum2 & " A_" & sYM
	next
	strSql = strSql & ","
	strV = "if(" & strFld & " <= '201211',Result,Plan)"
	strSql = strSql & strSum1 & strV & strSum2 & " A_Total"
	ChakuSelect = strSql
End Function
'-------------------------------------------------------------
'年月
'-------------------------------------------------------------
Function GetNendo(byVal strYM,byVal intM)
	dim	intYear
	dim	intMonth
	intYear = CInt(left(strYM,4))
	intMonth = CInt(right(strYM,2))
	if intMonth < 4 then
		intYear = intYear - 1
	end if
	if intM < 4 then
		intYear = intYear + 1
	end if
	GetNendo = intYear & right("0" & intM,2)
End Function
%>
