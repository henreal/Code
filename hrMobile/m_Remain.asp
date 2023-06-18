<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "待办"

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "Passed" Call ListPass()
	Case "Retreat" Call ListRetreat()
	Case "Affirm" Call ListAffirm()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	SiteTitle = "待办事务"
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.weui-tab {height: initial;}" & vbCrlf
	strHtml = strHtml & "		.weui-navbar__item.weui-bar__item--on {background-color:#eee;}" & vbCrlf
	strHtml = strHtml & "		.weui-navbar {z-index:101;}" & vbCrlf
	strHtml = strHtml & "		.remainBox {box-sizing: border-box;}" & vbCrlf
	strHtml = strHtml & "		.hr-list-title {margin:5px;padding:0 10px;box-sizing: border-box;line-height:35px;background:#ffe596;color:#900;}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li.item {align-items:inherit; border-bottom:8px solid #e3e3e3; box-sizing:border-box; padding:10px 0}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li.item em {border-top:1px solid #eee;box-sizing: border-box;padding:5px 0}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .icon, .hr-list-ul li .more {font-size:1.5rem;padding:0 5px;color:#ef8d75}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .more a {color:#aaa}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .info {flex-grow:2;}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .info span {color:#777;font-size:0.9rem;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim tCourseList, tAffirmList, tRetreatList
	i = 0
	sqlTmp = GetRemainCourse()		'未上课
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			tCourseList = tCourseList & "<ul class=""hr-list-ul"">"
			Do While Not rsTmp.EOF
				Set rs = Conn.Execute("Select a.*,b.ClassName,b.Template From HR_Sheet_" & rsTmp("ItemID") & " a Inner Join HR_Class B on a.ItemID=b.ClassID Where a.ID=" & rsTmp("ID"))
					If Not(rs.BOF And rs.EOF) Then
						tCourseList = tCourseList & "<li class=""hr-rows item""><em class=""icon""><i class=""hr-icon"">&#xef5f;</i></em>"
						tCourseList = tCourseList & "<em class=""info"">"
						If rs("Template") = "TempTableA" Then
							tCourseList = tCourseList & "授课时间：" & FormatDate(ConvertNumDate(rsTmp("VA4")), 4) & "<br><span>" & rs("VA5") & "周 第" & rs("VA7") & "节 " & GetPeriodTime(rs("VA11"), rs("VA7"),0) & "</span>"
							tCourseList = tCourseList & "<br><span>课程信息：" & rs("VA8") & " " & rs("VA9") & "</span>"
						Else
							tAffirmList = tAffirmList & "授课时间：" & FormatDate(ConvertNumDate(rs("VA4")), 2) & "<br><span>授课内容：" & rs("VA5") & "</span><br><span>所属项目：" & rs("ClassName") & "</span>"
						End If
						tCourseList = tCourseList & "</em><em class=""more""><a href=""" & ParmPath & "Course/View.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """><i class=""hr-icon"">&#xef8d;</i></a></em></li>"
					End If
				Set rs = Nothing
				rsTmp.MoveNext
				i = i + 1
			Loop
			tCourseList = tCourseList & "</ul>"
		End If
	Set rsTmp = Nothing
	tCourseList = "<h3 class=""hr-list-title"">待授课程：本学年共有<b>" & HR_CLng(i) & "</b>节</h3>" & tCourseList

	sqlTmp = GetItemUnionQuery(" And State=0")		'未确认
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			tAffirmList = tAffirmList & "<ul class=""hr-list-ul"">"
			Do While Not rsTmp.EOF
				Set rs = Conn.Execute("Select a.*,b.ClassName,b.Template From HR_Sheet_" & rsTmp("ItemID") & " a Inner Join HR_Class B on a.ItemID=b.ClassID Where ID=" & rsTmp("ID"))
					If Not(rs.BOF And rs.EOF) Then
						tAffirmList = tAffirmList & "<li class=""hr-rows item""><em class=""icon""><i class=""hr-icon"">&#xef5f;</i></em><em class=""info""><b>" & rs("ClassName") & "</b>"
						If rs("Template") = "TempTableA" Then
							tAffirmList = tAffirmList & " " & FormatDate(ConvertNumDate(rs("VA4")), 2) & " 第" & rs("VA7") & "节"
							tAffirmList = tAffirmList & "<br><span>" & rs("VA8") & " " & rs("VA9") & "</span>"
						ElseIf rs("Template") = "TempTableC" Then
							tAffirmList = tAffirmList & " " & FormatDate(ConvertNumDate(rs("VA4")), 2) & "<br><span>" & rs("VA5") & "</span>"
						Else
							tAffirmList = tAffirmList & " " & Trim(rs("VA4")) & "<br><span>" & rs("VA5") & "</span>"
						End If
						tAffirmList = tAffirmList & "</em><em class=""more""><a href=""" & ParmPath & "Course/View.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """><i class=""hr-icon"">&#xef8d;</i></a></em></li>"
					End If
				Set rs = Nothing
				rsTmp.MoveNext
				i = i + 1
			Loop
			tAffirmList = tAffirmList & "</ul>"
			tAffirmList = "<h3 class=""hr-list-title"">您共有<b>" & i & "</b>条业绩未确认！</h3>" & tAffirmList
		End If
	Set rsTmp = Nothing

	sqlTmp = GetItemUnionQuery(" And Retreat=" & HR_True)		'退回
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			tRetreatList = tRetreatList & "<ul class=""hr-list-ul"">"
			Do While Not rsTmp.EOF
				Set rs = Conn.Execute("Select a.*,b.ClassName,b.Template From HR_Sheet_" & rsTmp("ItemID") & " a Inner Join HR_Class B on a.ItemID=b.ClassID Where ID=" & rsTmp("ID"))
					If Not(rs.BOF And rs.EOF) Then
						tRetreatList = tRetreatList & "<li class=""hr-rows item""><em class=""icon""><i class=""hr-icon"">&#xefa8;</i></em><em class=""info""><b>" & rs("ClassName") & "</b>"
						If rs("Template") = "TempTableA" Then
							tRetreatList = tRetreatList & " " & FormatDate(ConvertNumDate(rs("VA4")), 2) & " 第" & rs("VA7") & "节"
							tRetreatList = tRetreatList & "<br><span>" & rs("VA8") & " " & rs("VA9") & "</span>"
						ElseIf rs("Template") = "TempTableC" Then
							tRetreatList = tRetreatList & " " & FormatDate(ConvertNumDate(rs("VA4")), 2) & "<br><span>" & rs("VA5") & "</span>"
						Else
							tRetreatList = tRetreatList & " " & Trim(rs("VA4")) & "<br><span>" & rs("VA5") & "</span>"
						End If
						tRetreatList = tRetreatList & "</em><em class=""more""><a href=""" & ParmPath & "Course/View.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """><i class=""hr-icon"">&#xef8d;</i></a></em></li>"
					End If
				Set rs = Nothing
				rsTmp.MoveNext
				i = i + 1
			Loop
			tRetreatList = tRetreatList & "</ul>"
			tRetreatList = "<h3 class=""hr-list-title"">您共有<b>" & i & "</b>条业绩被退回！</h3>" & tRetreatList
		Else
			tRetreatList = "没有数据"
		End If
	Set rsTmp = Nothing

	Response.Write "<div class=""weui-tab"">" & vbCrlf
	Response.Write "	<div class=""weui-navbar"">" & vbCrlf
	Response.Write "		<a class=""weui-navbar__item weui-bar__item--on"" href=""#tab1"">个人待办</a>" & vbCrlf
	Response.Write "		<a class=""weui-navbar__item"" href=""#tab2"">未确认业绩</a>" & vbCrlf
	Response.Write "		<a class=""weui-navbar__item"" href=""#tab3"">退回业绩</a>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-tab__bd"">" & vbCrlf
	Response.Write "		<div id=""tab1"" class=""weui-tab__bd-item weui-tab__bd-item--active"">" & vbCrlf
	Response.Write "			<div class=""remainBox"">" & tCourseList & "</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div id=""tab2"" class=""weui-tab__bd-item"">" & vbCrlf
	Response.Write "			<div class=""remainBox"">" & tAffirmList & "</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div id=""tab3"" class=""weui-tab__bd-item"">" & vbCrlf
	Response.Write "			<div class=""remainBox"">" & tRetreatList & "</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub ListPass()
	SiteTitle = "未审核业绩"
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.weui-tab {height: initial;}" & vbCrlf
	strHtml = strHtml & "		.weui-navbar__item.weui-bar__item--on {background-color:#eee;}" & vbCrlf
	strHtml = strHtml & "		.weui-navbar {z-index:101;}" & vbCrlf
	strHtml = strHtml & "		.remainBox {box-sizing: border-box;}" & vbCrlf
	strHtml = strHtml & "		.hr-list-title {margin:5px;padding-left:3px;box-sizing: border-box;line-height:35px;background:#f9ebeb;color:#911;border:1px solid #911}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li.item {border-top:1px solid #eee;box-sizing: border-box;padding:5px 0}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .icon, .hr-list-ul li .more {font-size:1.5rem;padding:0 3px;color:#ef8d75}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .more {padding:0} .hr-list-ul li .more a {color:#aaa}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .info {flex-grow:2;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = getPageHead(1)
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	strHeadHtml = ReplaceCommonLabel(strHeadHtml)

	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim tmpList
	sqlTmp = GetItemUnionQuery(" And Passed=" & HR_False)		'未审
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			tmpList = tmpList & "<ul class=""hr-list-ul"">"
			Do While Not rsTmp.EOF
				Set rs = Conn.Execute("Select a.*,b.ClassName,b.Template From HR_Sheet_" & rsTmp("ItemID") & " a Inner Join HR_Class B on a.ItemID=b.ClassID Where ID=" & rsTmp("ID"))
					If Not(rs.BOF And rs.EOF) Then
						tmpList = tmpList & "<li class=""hr-rows item""><em class=""icon""><i class=""hr-icon"">&#xe945;</i></em><em class=""info""><b>" & rs("ClassName") & "</b>"
						If rs("Template") = "TempTableA" Then
							tmpList = tmpList & "<br>" & FormatDate(ConvertNumDate(rs("VA4")), 2) & " 第" & rs("VA7") & "节"
							tmpList = tmpList & "<br><span>" & rs("VA8") & " " & rs("VA9") & "</span>"
						ElseIf rs("Template") = "TempTableC" Then
							tmpList = tmpList & "<br>" & FormatDate(ConvertNumDate(rs("VA4")), 2) & "<br><span>" & rs("VA5") & "</span>"
						Else
							tmpList = tmpList & "<br>" & Trim(rs("VA4")) & "<br><span>" & rs("VA5") & "</span>"
						End If
						tmpList = tmpList & "</em><em class=""more""><a href=""" & ParmPath & "Course/View.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """><i class=""hr-icon"">&#xef8d;</i></a></em></li>"
					End If
				Set rs = Nothing
				rsTmp.MoveNext
				i = i + 1
			Loop
			tmpList = tmpList & "</ul>"
			tmpList = "<h3 class=""hr-list-title""><i class=""hr-icon"">&#xe947;</i>您共有<b>" & i & "</b>条业绩未审核！</h3>" & tmpList
		Else
			tmpList = "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>您没有未审课程！</h3></div>" & vbCrlf
		End If
	Set rsTmp = Nothing
	Response.Write "<div class=""remainBox"">" & tmpList & "</div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub ListRetreat()
	SiteTitle = "退回业绩"
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.weui-tab {height: initial;}" & vbCrlf
	strHtml = strHtml & "		.weui-navbar__item.weui-bar__item--on {background-color:#eee;}" & vbCrlf
	strHtml = strHtml & "		.weui-navbar {z-index:101;}" & vbCrlf
	strHtml = strHtml & "		.remainBox {box-sizing: border-box;}" & vbCrlf
	strHtml = strHtml & "		.hr-list-title {margin:5px;padding-left:3px;box-sizing: border-box;line-height:35px;background:#fffde7;color:#776f0e;border:1px solid #776f0e}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li.item {border-top:1px solid #eee;box-sizing: border-box;padding:5px 0}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .icon, .hr-list-ul li .more {font-size:1.5rem;padding:0 3px;color:#ef8d75}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .more {padding:0} .hr-list-ul li .more a {color:#aaa}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .info {flex-grow:2;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = getPageHead(1)
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	strHeadHtml = ReplaceCommonLabel(strHeadHtml)

	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim tmpList
	sqlTmp = GetItemUnionQuery(" And Retreat=" & HR_True)		'退回
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			tmpList = tmpList & "<ul class=""hr-list-ul"">"
			Do While Not rsTmp.EOF
				Set rs = Conn.Execute("Select a.*,b.ClassName,b.Template From HR_Sheet_" & rsTmp("ItemID") & " a Inner Join HR_Class B on a.ItemID=b.ClassID Where ID=" & rsTmp("ID"))
					If Not(rs.BOF And rs.EOF) Then
						tmpList = tmpList & "<li class=""hr-rows item""><em class=""icon""><i class=""hr-icon"">&#xe945;</i></em><em class=""info""><b>" & rs("ClassName") & "</b>"
						If rs("Template") = "TempTableA" Then
							tmpList = tmpList & "<br>" & FormatDate(ConvertNumDate(rs("VA4")), 2) & " 第" & rs("VA7") & "节"
							tmpList = tmpList & "<br><span>" & rs("VA8") & " " & rs("VA9") & "</span>"
						ElseIf rs("Template") = "TempTableC" Then
							tmpList = tmpList & "<br>" & FormatDate(ConvertNumDate(rs("VA4")), 2) & "<br><span>" & rs("VA5") & "</span>"
						Else
							tmpList = tmpList & "<br>" & Trim(rs("VA4")) & "<br><span>" & rs("VA5") & "</span>"
						End If
						tmpList = tmpList & "</em><em class=""more""><a href=""" & ParmPath & "Course/View.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """><i class=""hr-icon"">&#xef8d;</i></a></em></li>"
					End If
				Set rs = Nothing
				rsTmp.MoveNext
				i = i + 1
			Loop
			tmpList = tmpList & "</ul>"
			tmpList = "<h3 class=""hr-list-title""><i class=""hr-icon"">&#xe947;</i>您共有<b>" & i & "</b>条业绩被退回！</h3>" & tmpList
		Else
			tmpList = "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>您没有退回课程！</h3></div>" & vbCrlf
		End If
	Set rsTmp = Nothing
	Response.Write "<div class=""remainBox"">" & tmpList & "</div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub ListAffirm()
	SiteTitle = "未确认业绩"
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.weui-tab {height: initial;}" & vbCrlf
	strHtml = strHtml & "		.weui-navbar__item.weui-bar__item--on {background-color:#eee;}" & vbCrlf
	strHtml = strHtml & "		.weui-navbar {z-index:101;}" & vbCrlf
	strHtml = strHtml & "		.remainBox {box-sizing: border-box;}" & vbCrlf
	strHtml = strHtml & "		.hr-list-title {margin:5px;padding-left:3px;box-sizing: border-box;line-height:35px;background:#fffde7;color:#776f0e;border:1px solid #776f0e}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li.item {border-top:1px solid #eee;box-sizing: border-box;padding:5px 0}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .icon, .hr-list-ul li .more {font-size:1.5rem;padding:0 3px;color:#ef8d75}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .more {padding:0} .hr-list-ul li .more a {color:#aaa}" & vbCrlf
	strHtml = strHtml & "		.hr-list-ul li .info {flex-grow:2;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = getPageHead(1)
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	strHeadHtml = ReplaceCommonLabel(strHeadHtml)

	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim tmpList
	sqlTmp = GetItemUnionQuery(" And State=0")		'未确认
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			tmpList = tmpList & "<ul class=""hr-list-ul"">"
			Do While Not rsTmp.EOF
				Set rs = Conn.Execute("Select a.*,b.ClassName,b.Template From HR_Sheet_" & rsTmp("ItemID") & " a Inner Join HR_Class B on a.ItemID=b.ClassID Where ID=" & rsTmp("ID"))
					If Not(rs.BOF And rs.EOF) Then
						tmpList = tmpList & "<li class=""hr-rows item""><em class=""icon""><i class=""hr-icon"">&#xe945;</i></em><em class=""info""><b>" & rs("ClassName") & "</b>"
						If rs("Template") = "TempTableA" Then
							tmpList = tmpList & "<br>" & FormatDate(ConvertNumDate(rs("VA4")), 2) & " 第" & rs("VA7") & "节"
							tmpList = tmpList & "<br><span>" & rs("VA8") & " " & rs("VA9") & "</span>"
						ElseIf rs("Template") = "TempTableC" Then
							tmpList = tmpList & "<br>" & FormatDate(ConvertNumDate(rs("VA4")), 2) & "<br><span>" & rs("VA5") & "</span>"
						Else
							tmpList = tmpList & "<br>" & Trim(rs("VA4")) & "<br><span>" & rs("VA5") & "</span>"
						End If
						tmpList = tmpList & "</em><em class=""more""><a href=""" & ParmPath & "Course/View.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """><i class=""hr-icon"">&#xef8d;</i></a></em></li>"
					End If
				Set rs = Nothing
				rsTmp.MoveNext
				i = i + 1
			Loop
			tmpList = tmpList & "</ul>"
			tmpList = "<h3 class=""hr-list-title""><i class=""hr-icon"">&#xe947;</i>您共有<b>" & i & "</b>条业绩未确认！</h3>" & tmpList
		Else
			tmpList = "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>您没有未确认的课程！</h3></div>" & vbCrlf
		End If
	Set rsTmp = Nothing
	Response.Write "<div class=""remainBox"">" & tmpList & "</div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Function GetItemUnionQuery(fWhere)		'获取未确认及退回业绩
	Dim iFun, funItem, arrItem, strFun : strFun = ""
	funItem = GetItemClassID("")		'取考核项ID
	If HR_IsNull(funItem) = False Then
		arrItem = Split(FilterArrNull(funItem, ","), ",")
		For iFun = 0 To Ubound(arrItem)
			If iFun > 0 Then strFun = strFun & " union all "
			strFun = strFun & "(Select ID,ItemID,VA1,VA2 From HR_Sheet_" & arrItem(iFun) & " Where VA1=" & HR_Clng(UserYGDM) & " And scYear=" & DefYear
			If HR_IsNull(fWhere) = False Then strFun = strFun & " " & fWhere
			strFun = strFun & ")"
		Next
	End If
	GetItemUnionQuery = strFun
End Function

Function GetRemainCourse()				'取未上课程
	Dim iFun, funItem, arrItem, strFun : strFun = ""
	funItem = GetItemClassID(" And Template in('TempTableA','TempTableC') ")		'取考核项ID
	If HR_IsNull(funItem) = False Then
		arrItem = Split(FilterArrNull(funItem, ","), ",")
		For iFun = 0 To Ubound(arrItem)
			If iFun > 0 Then strFun = strFun & " union all "
			strFun = strFun & "(Select ID,ItemID,VA1,VA2,VA4 From HR_Sheet_" & arrItem(iFun) & " Where VA4>CAST((GetDate()) As SMALLDATETIME) And VA1=" & HR_Clng(UserYGDM) & " And scYear=" & DefYear
			strFun = strFun & ")"
		Next
	End If
	GetRemainCourse = strFun
End Function
%>