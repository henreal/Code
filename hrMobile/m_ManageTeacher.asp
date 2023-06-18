<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl : SiteTitle = "教师管理"

Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "GetTeacher" Call GetTeacherList()
	Case "View" Call View()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim Count1		'汇总教师

	Set rsTmp = Conn.Execute("Select count(TeacherID) From HR_Teacher Where Cast(YGDM As int)>0 And XMJP<>''")
		Count1 = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing

	tmpHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	tmpHtml = tmpHtml & "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-box {border:0px solid #F60;background-color: rgba(255,255,255,1);box-sizing: border-box;overflow-y:auto;position:absolute;top:140px;left:0;bottom:0;right:0;z-index:10}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-letter {position:fixed;height:100%;width:25px;right:0;top:140px;background-color:#e3e3e3;z-index:11}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-letter ul {display:flex;text-align:center;flex-wrap:wrap;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-letter li {display:block;text-align:center;width:100%;font-size:0.8rem}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-search-bar{background-color:#1ab4c7;} .weui-search-bar__cancel-btn {color:#fff}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-search-bar__form{background-color: transparent;}" & vbCrlf
	tmpHtml = tmpHtml & "		.addInfo {padding:3px 10px; display:flex;align-items: center;}" & vbCrlf
	tmpHtml = tmpHtml & "		.addInfo .headimg {padding-right:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.addInfo .headimg em {text-align: center;width:35px;height:35px;line-height:35px;background:#607d8b;color:#fff;border-radius:4px}" & vbCrlf
	tmpHtml = tmpHtml & "		.addInfo .num_name {line-height:1.1rem;} .addInfo .num_name b {font-size:0.7rem;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", tmpHtml)
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-cell"">" & vbCrlf
	Response.Write "	<div class=""weui-cell__bd"">全部教师</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell__ft"">" & Count1 & " 名</div>"  & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""weui-search-bar"" id=""searchBar"">" & vbCrlf
	Response.Write "	<form class=""weui-search-bar__form"">" & vbCrlf
	Response.Write "		<div class=""weui-search-bar__box"">" & vbCrlf
	Response.Write "			<i class=""weui-icon-search""></i><input type=""search"" class=""weui-search-bar__input"" id=""searchInput"" placeholder=""搜索"" required="""">" & vbCrlf
	Response.Write "			<a href=""javascript:"" class=""weui-icon-clear"" id=""searchClear""></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<label class=""weui-search-bar__label"" id=""searchText""><i class=""weui-icon-search""></i><span>搜索姓名/拼音/工号</span></label>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	<a href=""#"" class=""weui-search-bar__cancel-btn"" id=""searchCancel"">取消</a>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-addbook-box"">" & vbCrlf
	Response.Write "	<div class=""hr-addbook-letter""><ul>" & vbCrlf
	Response.Write "	<li data-id=""*""><i class=""hr-icon"">&#xef3e;</i></li>" & vbCrlf
	For i = 65 To 90
		Response.Write "	<li data-id=""" & Chr(i) & """>" & Chr(i) & "</li>" & vbCrlf
	Next
	Response.Write "	<li data-id=""#"">#</li>" & vbCrlf
	Response.Write "	</ul></div>" & vbCrlf
	Response.Write "	<div class=""hr-addbook-list"" id=""listGroup"">" & vbCrlf
	Response.Write "		<div class=""weui-loadmore""><i class=""weui-loading""></i><span class=""weui-loadmore__tips"">正在加载</span></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf


	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Manage/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	var boxHeight = $("".hr-addbook-box"").height()-10;" & vbCrlf
	tmpHtml = tmpHtml & "	var liH = boxHeight/28;" & vbCrlf
	tmpHtml = tmpHtml & "	$('.hr-addbook-letter li').height(liH);" & vbCrlf
	tmpHtml = tmpHtml & "	getTeacherList("""");" & vbCrlf
	
	tmpHtml = tmpHtml & "	$("".hr-addbook-letter li"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var lmd =$(this).data(""id""); getTeacherList(lmd);" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#searchInput"").bind(""input propertychange"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		getTeacherList($(this).val());" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	function getTeacherList(fLetter){" & vbCrlf		'取教师数据
	tmpHtml = tmpHtml & "		$.get(""" & ParmPath & "ManageTeacher/GetTeacher.html"",{Type:0, Letter:fLetter}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#listGroup"").html(rsStr);" & vbCrlf
	tmpHtml = tmpHtml & "			$("".num_name"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "				location.href=""" & ParmPath & "ManageTeacher/View.html?ygdm=""+ $(this).data(""ygdm"");" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	}" & vbCrlf

	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strFootHtml = Replace(strFootHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

Sub GetTeacherList()
	Dim sqlList, rsList, tFirstWord, tTel, tTel1, tPRZC
	Dim listType : listType = HR_Clng(Request("Type"))
	Dim tLetter : tLetter = Trim(ReplaceBadChar(Request("Letter")))
	sqlList = "Select Top 2000 * From HR_Teacher Where Cast(YGDM As int)>0 And XMJP<>''"
	If listType > 0 Then sqlList = sqlList & " And ApiType=" & listType
	If HR_IsNull(tLetter) = False Then
		tFirstWord = UCase(Left(tLetter, 1))
		If Asc(tFirstWord) > 64 And Asc(tFirstWord) < 91 Then
			sqlList = sqlList & " And XMJP like('" & tLetter & "%')"
		ElseIf HR_Clng(tFirstWord) > 0 Then
			sqlList = sqlList & " And YGDM like('" & tLetter & "%')"
		Else
			sqlList = sqlList & " And YGXM like('" & tLetter & "%')"
		End If
	End If
	sqlList = sqlList & " Order By YGXM collate Chinese_PRC_CS_AS_KS_WS"
	Set rsList = Server.CreateObject("ADODB.RecordSet")
		rsList.Open(sqlList), Conn, 1, 1
		If Not(rsList.BOF And rsList.EOF) Then
			m = 0
			Redim ZM(26,1)
			Dim py
			For i = 65 To 90
				ZM(i-64,0) = Chr(i)
				ZM(i-64,1) = "0"
			Next
			Do While Not rsList.EOF
				'If m > 0 Then strList = strList & ","
				tTel = ""
				tPRZC = rsList("PRZC") : tPRZC = Replace(tPRZC, "无职称", "")
				py = UCase(Left(rsList("XMJP"), 1))
				If HR_IsNull(rsList("SJHM")) = False Then tTel = "<br>手机：" & Trim(rsList("SJHM")) & " " & Trim(rsList("DH"))
				If ZM(ASC(py)-64,1) = "0" Then
					Response.Write "<div class=""sort_letter"" id=""l_" & py & """>" & py & "</div>"
					ZM(ASC(py)-64,1) = "1"
				End If
				Response.Write "	<div class=""sort_list addInfo"">" & vbCrlf
				Response.Write "		<div class=""headimg""><em>" & UCase(Left(rsList("YGXM"), 1)) & "</em></div><div class=""num_name"" data-ygdm=""" & rsList("YGDM") & """ data-ygxm=""" & rsList("YGXM") & """>" & rsList("YGXM") & "<b> [工号：" & rsList("YGDM") & "]</b><br><b>" & rsList("KSMC") & " " & tPRZC & "</b></div>" & vbCrlf
				Response.Write "	</div>" & vbCrlf
				rsList.MoveNext
				m = m + 1
			Loop
		Else
			Response.Write "	<div class=""sort_list"">没有教师</div>" & vbCrlf
		End If
	Set rsList = Nothing
End Sub

Sub View()
	Dim tYGDM : tYGDM = HR_Clng(Request("ygdm"))
	SiteTitle = "教师详情"

	Dim tYGXM, tXMJP, tYGXB, tYGZT, tKSDM, tKSMC, tPRZC, tXZZW
	sql = "Select * From HR_Teacher Where YGDM='" & tYGDM & "'"
	Set rs = Conn.Execute(sql)
		If Not(rs.BOF And rs.EOF) Then
			tYGXM = rs("YGXM")
			tXMJP = rs("XMJP")
			tYGXB = rs("YGXB")
			tYGZT = rs("YGZT")
			tKSDM = rs("KSDM")
			tKSMC = rs("KSMC")
			tPRZC = rs("PRZC")
			tXZZW = rs("XZZW")
		End If
	Set rs = Nothing

	Dim jsonOBJ, strJson : strJson = GetWechatUserInfoQY(tYGDM)
	Dim tHeadFace, tMobile, tEmail, tQrcode
	Set jsonOBJ = parseJSON(strJson)
		If jsonOBJ.errcode = 0 Then
			tHeadFace = Trim(jsonOBJ.avatar)
			tMobile = Trim(jsonOBJ.mobile)
			tEmail = Trim(jsonOBJ.email)
			tQrcode = Trim(jsonOBJ.qr_code)
		End If
	Set jsonOBJ = Nothing
	If HR_IsNull(tHeadFace) Then tHeadFace = InstallDir & "Static/images/nopic.png"

	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.backbtn {padding:10px 5px;}" & vbCrlf
	strHtml = strHtml & "		.HeadFace {width:55px;height:55px;background:#fff center no-repeat;background-size:100% auto;border-radius: 5px;}" & vbCrlf
	strHtml = strHtml & "		.weui-photo-browser-modal {background-color:rgba(0,0,0,0.8);z-index:1000;}" & vbCrlf
	strHtml = strHtml & "		.qrcode {font-size:1.5rem;color:#0bd;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<div class=""weui-cells"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">教师姓名：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tYGXM & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">教师工号：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tYGDM & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">拼音代码：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tXMJP & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">性　别：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tYGXB & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">状　态：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tYGZT & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">科　室：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tKSMC & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">职　称：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tPRZC & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">职　务：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tXZZW & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">头　像：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><em class=""HeadFace"" style=""background-image:url(" & tHeadFace & ")""></em></div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">手　机：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tMobile & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">邮　箱：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tEmail & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">二维码：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><i class=""hr-icon qrcode"" data-qr=""" & tQrcode & """>&#xec29;</i></div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-fix backbtn""><button type=""button"" name=""back"" class=""weui-btn weui-btn_plain-default"" id=""back"">返回</button></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/js/swiper.min.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "ManageTeacher/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".HeadFace"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var img=$(this).css(""backgroundImage"").split(""("")[1].split("")"")[0];img = img.replace(""\"""","""");" & vbCrlf
	strHtml = strHtml & "		var pb1 = $.photoBrowser({ items:[{image:img, caption:""" & tYGXM & " 的头像""}] }); pb1.open();" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$("".qrcode"").on(""click"",function(){ var qrcode=$(this).data(""qr""); var pb2 = $.photoBrowser({ items:[{image:qrcode,caption:""" & tYGXM & " 的个人二维码""}] }); pb2.open(); });" & vbCrlf
	strHtml = strHtml & "	$("".backbtn"").on(""click"",function(){ location.href=""" & ParmPath & "ManageTeacher/Index.html""; });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

%>