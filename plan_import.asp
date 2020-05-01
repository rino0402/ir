<%@LANGUAGE=VBScript%>
<%Option Explicit%>
<% Response.Buffer = True %>
<% Response.Expires = -1 %>
<%
'--------------------------------------------------------------------
'バージョン情報
'--------------------------------------------------------------------
Function GetVersion()
	GetVersion = "2012.05.19 新規作成"
	GetVersion = "2012.12.14 2013年度対応...作業中..."
	GetVersion = "2013.05.11 年度の初期値に現在の年をセットするように変更"
	GetVersion = "2017.04.28 人件費(パート) にB列が追加された形式に対応"
End Function
'--------------------------------------------------------------------
%>
<%
'--------------------------------------------------------------------
'POSTデータの受取
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
<title>計画用作業時間 登録(アップロード)</title>
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
<form name="fileForm" METHOD="post" enctype="multipart/form-data" ACTION="plan_import_call.asp">
	<table>
	<caption style="text-align:left;">計画用作業時間ファイル(Excel)登録(アップロード)</caption>
	<tr>
		<td align="right">年度</td>
		<td align="left">
			<INPUT TYPE="text" NAME="YYYY" id="YYYY" VALUE="<%=GetRequest("YYYY")%>" size="10" class="input" style="text-align:center;" required pattern="^[0-9]+$">
			<span>登録する事業計画の年度を入力</span>
		</td>
	</tr>
	<tr>
		<td align="right">センター</td>
		<td align="left">
			<div><INPUT class="input" TYPE="text" NAME="CenterCD" id="CenterCD" VALUE="<%=GetRequest("CenterCD")%>" size="2" style="text-align:center;" required>
			B:小野PC/D:滋賀物流/E:滋賀PC/F:奈良/G:大阪PC/H:袋井PC/I:広島</div>
		</td>
	</tr>
	<tr>
		<td align="right">ファイル名</td>
		<td><input class="input" type="file" name="fName" size="100" required><br>
		計画用作業時間ファイル(Excel)を指定して下さい。例：経営概況書事業計画時間分配.xls
		</td>
	</tr>
	<tr>
		<td align="right">処理</td>
		<td>
			<!--input class="cssbutton" type="submit" value="実行"		id="btnSubmit"	onclick="return submitNew();"-->
			<input class="cssbutton" type="submit" value="全て実行"				id="pAll"		onclick="return submitAction(this);">
			<input class="" 		 type="submit" value="1.アップロードのみ"	id="pFile"		onclick="return submitAction(this);">
			<input class="" 		 type="submit" value="2.作業時間登録のみ"	id="pLoad"		onclick="return submitAction(this);">
			<input class="" 		 type="submit" value="3.配分処理のみ"		id="pHaibun"	onclick="return submitAction(this);">
			<INPUT class="cssbutton" TYPE="reset"  value="リセット"	id="btnReset"	onClick="location.href='<%=Request.ServerVariables("URL")%>';">
		</td>
	</tr>
	</table>
	<div class="info"><%=getVersion()%></div>
</form>
<h3>アップロードするファイルのサンプルイメージ</h3>
*.xlsx は使用できません。*.xls で保存したファイルを使用して下さい。
<h4>人件費(社員)</h4>
<div><img src="jinkenhi.jpg" alt="人件費(社員)"><div>
<h4 class="info_new">人件費(パート)2017年度 B列追加</h4>
<div><img src="pert2017.jpg" alt="人件費(パート)B列追加"><div>
<h4>人件費(パート)</h4>
<div><img src="pert.jpg" alt="人件費(パート)"><div>
<h4>経費配分表</h4>
<div><img src="haibun.jpg" alt="経費配分表"><div>
</body>
</html>
<script type="text/javascript">
    <!--
        function submitNew() {
			var w = 600;
			var h = 480;
			var x = (screen.width - w) / 2;
			var y = (screen.height - h) / 2;
            //① blankウィンドウを開く。
            var win = window.open("","myNewWnd","screenX="+x+",screenY="+y+",left="+x+",top="+y+",width="+w+",height="+h+",status=1,scrollbar=yes");
            
            //② myformのaction属性を取り出し保存
            var action = document.fileForm.action;
            //③ 押下されたボタンがbtnSubmitボタンであることをサーバに通知
            document.fileForm.action = action + "?btnSubmit=dummy";
            //④ レスポンスはmyNewWnd画面が受け取るように設定
            document.fileForm.target = "myNewWnd";
            //⑤ フォームのサブミット
            document.fileForm.submit();
            //⑥ myNewWnd画面がフォーカスをあてる
            win.focus();
            //⑦ 他のボタンに影響をしないようにtarget及びactionの値を元に戻す
            document.fileForm.target = "_self";
            document.fileForm.action = action;
            return false;
        }
        function submitAction(t) {
			var w = 600;
			var h = 480;
			var x = (screen.width - w) / 2;
			var y = (screen.height - h) / 2;
            //① blankウィンドウを開く。
            var win = window.open("","myNewWnd","screenX="+x+",screenY="+y+",left="+x+",top="+y+",width="+w+",height="+h+",status=1,scrollbar=yes");
            
            //② myformのaction属性を取り出し保存
			var action = document.fileForm.action;
//			var action = "plan_import_action.asp";
            //③ 押下されたボタンがbtnSubmitボタンであることをサーバに通知
            document.fileForm.action = "plan_import_action.asp?btnSubmit="+ t.id;
            //④ レスポンスはmyNewWnd画面が受け取るように設定
            document.fileForm.target = "myNewWnd";
            //⑤ フォームのサブミット
            document.fileForm.submit();
            //⑥ myNewWnd画面がフォーカスをあてる
            win.focus();
            //⑦ 他のボタンに影響をしないようにtarget及びactionの値を元に戻す
            document.fileForm.target = "_self";
            document.fileForm.action = action;
            return false;
        }
    //-->
</script>

