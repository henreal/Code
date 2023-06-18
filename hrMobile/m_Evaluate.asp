<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./m_EvaluateCEX.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "评价"

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "TeachQuality" Call TeachQuality()
	Case "EditQuality" Call EditQuality()
	Case "SaveQuality" Call SaveQuality()
	Case "getItemCourse" Call getItemCourse()
	Case "getCourse" Call getCourse()
	Case "CEX" Call CEX()
	Case "Apply" Call Apply()
	Case "ApplyModify" Call ApplyModify()
	Case "SaveApply" Call SaveApply()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		body {background-color:#fff;}" & vbCrlf
	strHtml = strHtml & "		.navExtend {height: initial;flex-grow:2;text-align:right;}" & vbCrlf
	strHtml = strHtml & "		.navExtend span {font-size:1.2rem;display:line-block;background-color:#f7ce93;padding:2px 3px;color:#035;border-radius: 2px}" & vbCrlf
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
	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf
	
	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "Evaluate/TeachQuality.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xe9dc;</i></div><div class=""weui-cell__bd weui-cell_primary viewMSG"" data-id=""1""><p>课堂教学质量评价</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "Evaluate/CEX.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xe9dc;</i></div><div class=""weui-cell__bd weui-cell_primary viewMSG"" data-id=""2""><p>mini-CEX<sup>plus</sup>记录</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var content = $("".weui-textarea"").val();" & vbCrlf
	strHtml = strHtml & "		if(content==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""内容太少"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else{" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Propose/SavePost.html"", {Content:content}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(rsStr.reMessge, function(){ $.closePopup();location.reload(); });" & vbCrlf
	strHtml = strHtml & "			},""json"");" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "		" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub TeachQuality()
	SiteTitle = "课堂教学质量评价"
	Dim rsList, strList, dataTable, tParm, tmpID
	If Ubound(arrParm) > 1 Then
		tParm = Trim(arrParm(2))
		tmpID = HR_Clng(Request("ID"))
		Select Case tParm
			Case "ViewQuality" Call ViewQuality()
			Case "SaveQuality" Call SaveQuality()
		End Select
		Exit Sub
	End If
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn {width:55px;height:55px;text-align:center;font-size:2.6rem;position: fixed;right: 1rem;bottom: 3rem;z-index: 99;}" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn i {color:#814ee2;}" & vbCrlf
	strHtml = strHtml & "		.iconTit {color:#814ee2;padding-right:15px;}" & vbCrlf
	strHtml = strHtml & "		.viewMSG h5 {color:#777;}" & vbCrlf
	strHtml = strHtml & "		.viewMSG .pass {color:#fff;background-color:#ce67b9;padding:0 5px;}" & vbCrlf
	strHtml = strHtml & "		.itemHref {align-items:flex-start;}" & vbCrlf
	strHtml = strHtml & "		.itemHref .weui-cell__ft:after {margin-top:3px;}" & vbCrlf
	strHtml = strHtml & "		.Count {padding:0 15px;} .Count b {padding:0 3px;color:#900}" & vbCrlf
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

	Dim tPassed, tCount
	Set rsList = Conn.Execute("Select * From HR_Evaluate Where ParticipantID=" & UserYGDM)
		If Not(rsList.BOF And rsList.EOF) Then
			tCount = 0
			Do While Not rsList.EOF
				dataTable = "HR_Sheet_" & rsList("ItemID")
				tPassed = ""
				If HR_CBool(rsList("Passed")) Then tPassed = "<span class=""pass"">已提交</span>"
				strList = strList & "	<a class=""weui-cell weui-cell_access itemHref"" href=""" & ParmPath & "Evaluate/TeachQuality/ViewQuality.html?ID=" & rsList("ID") & """>" & vbCrlf
				strList = strList & "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xead1;</i></div>" & vbCrlf
				strList = strList & "		<div class=""weui-cell__bd weui-cell_primary viewMSG"" data-id=""1""><h4>授课教师：" & rsList("Teacher") & "</h4><h4>课程：" & rsList("Course") & "</h4><h5>评价时间：" & FormatDate(rsList("CreateTime"),10) & "　" & tPassed & "</h5></div>" & vbCrlf
				strList = strList & "		<div class=""weui-cell__ft""></div>" & vbCrlf
				strList = strList & "	</a>" & vbCrlf
				rsList.MoveNext
				tCount = tCount + 1
			Loop
		Else
			strList = "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>您暂时还没有发表评价！</h3></div>" & vbCrlf
		End If
	Set rsList = Nothing
	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""Count""><em>您共有<b>" & HR_CLng(tCount) & "</b>次评价</em></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf
	Response.Write " " & strList
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-float-btn"">" & vbCrlf
	Response.Write "	<a href=""" & ParmPath & "Evaluate/EditQuality.html?AddNew=True"" class=""addBtn""><i class=""hr-icon"">&#xf3c0;</i></a>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var content = $("".weui-textarea"").val();" & vbCrlf
	strHtml = strHtml & "		if(content==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""内容太少"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else{" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Propose/SavePost.html"", {Content:content}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(rsStr.reMessge, function(){ $.closePopup();location.reload(); });" & vbCrlf
	strHtml = strHtml & "			},""json"");" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "		" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub EditQuality()
	Dim tmpHtml, tmpID : tmpID = HR_Clng(Request("ID"))
	Dim isModify : isModify = HR_CBool(Request("Modify"))
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tCourseID : tCourseID = HR_Clng(Request("CourseID"))

	Dim dataTable, tTeacher, tTeacherID, tCourse, tContents, tStuClass, tAddress, tTitle, tPeriod, tItemName, tClassTime
	Dim tCampus, tScore1, tScore2, tScore3, tScore4, tScore5, tScore6, tScore7, tScore8, tScore9, tScore10
	Dim tTotalScore, tMerit, tSuggest0, tSuggest1, tSuggest2, tSuggest3, tCreateTime

	dataTable = "HR_Sheet_" & tItemID
	If ChkTable(dataTable) Then
		Set rs = Conn.Execute("Select a.*,b.ClassName,b.Template From " & dataTable & " a Inner Join HR_Class b on a.ItemID=b.ClassID Where a.ID=" & tCourseID)
			If Not(rs.BOF And rs.EOF) Then
				tItemName = rs("ClassName")
				tTeacher = rs("VA2")
				tTeacherID = rs("VA1")
				tClassTime = FormatDate(ConvertNumDate(rs("VA4")), 4)
				If rs("Template") = "TempTableA" Then
					tCourse = rs("VA8")
					tContents = rs("VA9")
					tStuClass = rs("VA10")
					tAddress = rs("VA11") & " " & rs("VA12")
					tPeriod = "第" & rs("VA7") & "节 " & GetPeriodTime(rs("VA11"), rs("VA7"), 0)
				Else
					tCourse = rs("VA6")
					tContents = rs("VA7")
				End If
			End If
		Set rs = Nothing
	End If

	Set rsTmp = Conn.Execute("Select * From HR_Evaluate Where ID=" & tmpID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			dataTable = "HR_Sheet_" & rsTmp("ItemID")
			tTeacher = rsTmp("Teacher")
			tTeacherID = rsTmp("TeacherID")
			tItemID = rsTmp("ItemID")
			tCourseID = rsTmp("CourseID")
			tCourse = rsTmp("Course")
			tTitle = rsTmp("Title")
			tCampus = rsTmp("Campus")
			tScore1 = rsTmp("Score1")
			tScore2 = rsTmp("Score2")
			tScore3 = rsTmp("Score3")
			tScore4 = rsTmp("Score4")
			tScore5 = rsTmp("Score5")
			tScore6 = rsTmp("Score6")
			tScore7 = rsTmp("Score7")
			tScore8 = rsTmp("Score8")
			tScore9 = rsTmp("Score9")
			tScore10 = rsTmp("Score10")
			tTotalScore = rsTmp("TotalScore")
			tSuggest0 = rsTmp("Suggest0")
			tSuggest1 = rsTmp("Suggest1")
			tSuggest2 = rsTmp("Suggest2")
			tSuggest3 = rsTmp("Suggest3")
			If ChkTable(dataTable) Then
				Set rs = Conn.Execute("Select a.*,b.ClassName,b.Template From " & dataTable & " a Inner Join HR_Class b on a.ItemID=b.ClassID Where a.ID=" & rsTmp("CourseID"))
					If Not(rs.BOF And rs.EOF) Then
						tItemName = rs("ClassName")
						tClassTime = FormatDate(ConvertNumDate(rs("VA4")), 4)
						If rs("Template") = "TempTableA" Then
							tCourse = rs("VA8")
							tContents = rs("VA9")
							tStuClass = rs("VA10")
							tAddress = rs("VA11") & " " & rs("VA12")
							tPeriod = "第" & rs("VA7") & "节 " & GetPeriodTime(rs("VA11"), rs("VA7"), 0)
						Else
							tCourse = rs("VA6")
							tContents = rs("VA7")
						End If
					End If
				Set rs = Nothing
			End If
		End If
	Set rsTmp = Nothing
	SiteTitle = "课堂教学质量评价"
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-swap-items dl {align-items:stretch;border-bottom:1px solid #eee;padding:5px 0;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .weui-cell__ft {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .popWin i {font-size:1.2rem;position:relative;top:2px;color:#4ce}" & vbCrlf
	strHtml = strHtml & "		.weui-toast {margin-left: auto;} .weui-textarea{font-size:1rem}" & vbCrlf
	strHtml = strHtml & "		.hr-tips {border-top:1px solid #eee;padding:5px} .hr-tips .tipsIcon {color:#f30;padding-right:5px;font-size:1.5rem;}" & vbCrlf
	strHtml = strHtml & "		.title {border-top:1px solid #eee;padding:10px} .title h3 {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.remarkbar {padding:5px 10px;border-top:1px solid #eee;color:#444;background-color:#f2f2f2}" & vbCrlf
	strHtml = strHtml & "		.weui-count .weui-count__number {font-size:1.1rem;width:2rem;}" & vbCrlf

	strHtml = strHtml & "		.verifyerr {background-color:#ffc8ba}" & vbCrlf
	strHtml = strHtml & "		.popbtn {position:fixed;bottom:0px;left:0;right:0;padding:10px;z-index:10;}" & vbCrlf

	strHtml = strHtml & "		.hr-pop-bg {display:none;position:fixed;background-color:rgba(0,0,0,0.3);width:100%;height:100%;z-index:1000;top:0;}" & vbCrlf
	strHtml = strHtml & "		.hr-pop-box {position:fixed;background-color:#fff;width:0;height:100%;z-index:1001;top:0;overflow:hidden;box-sizing:border-box;}" & vbCrlf
	strHtml = strHtml & "		.hr-pop-tit {border-bottom:1px solid #eee;height:46px;line-height:46px;box-sizing:border-box;padding:0 20px;}" & vbCrlf
	strHtml = strHtml & "		.hr-pop-tit .close {border:2px solid #59bfe4;height:35px;line-height:35px;width:35px;box-sizing:border-box;text-align:center;border-radius:50%;background-color:#2fabd8;color:#fff;}" & vbCrlf

	strHtml = strHtml & "		.popbox {box-sizing:border-box;overflow-y:auto;height:100%;}" & vbCrlf
	strHtml = strHtml & "		.item-list {box-sizing:border-box;padding:15px;}" & vbCrlf
	strHtml = strHtml & "		.item-list em {border-bottom:1px solid #eee;line-height:40px;padding:0 10px}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write "<header class=""hr-rows hr-header"">" & vbCrlf
	Response.Write "	<nav class=""navBack""><em><i class=""hr-icon"">&#xf320;</i></em></nav>" & vbCrlf
	Response.Write "	<nav class=""navTitle""><span>" & SiteTitle & "</span></nav>" & vbCrlf
	Response.Write "	<nav class=""navMenu""><em><i class=""hr-icon"">&#xf32a;</i></em></nav>" & vbCrlf
	Response.Write "</header>" & vbCrlf
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课教师：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Teacher"" class=""weui-input verify"" id=""Teacher"" type=""text"" value=""" & tTeacher & """ data-key=""Teacher"" data-value=""TeacherID"" placeholder="""">" & vbCrlf
	Response.Write "			<input name=""TeacherID"" class=""weui-input"" id=""TeacherID"" type=""hidden"" value=""" & tTeacherID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft popWin"" data-id=""Teacher""><i class=""hr-icon"">&#xeeed;</i>选择</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择项目：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Item"" class=""weui-input"" id=""Item"" type=""text"" value=""" & tItemName & """>" & vbCrlf
	Response.Write "			<input name=""ItemID"" class=""weui-input"" id=""ItemID"" type=""hidden"" value=""" & tItemID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft itempop"" data-name=""Item""><i class=""hr-icon"">&#xef63;</i>选择</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择课程：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Course"" class=""weui-input verify"" id=""Course"" type=""text"" value=""" & tCourse & """>" & vbCrlf
	Response.Write "			<input name=""CourseID"" class=""weui-input"" id=""CourseID"" type=""hidden"" value=""" & tCourseID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft coursepop"" data-id=""Teacher""><i class=""hr-icon"">&#xef63;</i>选择</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课对象：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Student"" class=""weui-input verify"" id=""Student"" type=""text"" value=""" & tStuClass & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课内容：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Contents"" class=""weui-input"" id=""Contents"" type=""text"" value=""" & tContents & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课时间：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""CourseDate"" class=""weui-input verify"" id=""CourseDate"" type=""text"" value=""" & tClassTime & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">开课学院：</label></div>" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	'Response.Write "			<input name=""Campus"" class=""weui-input verify"" id=""Campus"" type=""text"" value=""" & tCampus & """ readonly>" & vbCrlf
	'Response.Write "		</div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""hr-rows hr-tips"">" & vbCrlf
	Response.Write "		<em class=""tipsIcon""><i class=""hr-icon"">&#xf06a;</i></em>" & vbCrlf
	Response.Write "		<em class=""hr-row-fill tipstxt"">评价标准：>9优/9-6良/<6欠佳</em>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""title""><h3>一、教学态度与基本技能</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、要求脱稿讲授，语言准确流畅，逻辑性强，富感染力，语速、语调适宜、抑扬顿挫。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score1"" id=""Score1"" class=""weui-count__number Score"" type=""number"" value=""" & HR_CLng(tScore1) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark1"" id=""Remark1"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、精神饱满，教态大方，仪表端正。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score2"" id=""Score2"" class=""weui-count__number Score"" type=""number"" value=""" & HR_CLng(tScore2) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark2"" id=""Remark2"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、PPT设计科学，板书工整，教案讲稿规范。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score3"" id=""Score3"" class=""weui-count__number Score"" type=""number"" value=""" & HR_CLng(tScore3) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark3"" id=""Remark3"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>二、教学设计与方法</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、运用先进教学理念、方法进行教学，三维目标明确，学情清楚，因材施教，循循善诱。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score4"" id=""Score4"" class=""weui-count__number Score"" type=""number"" value=""" & HR_CLng(tScore4) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark4"" id=""Remark4"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、教学设计科学，新课导入、知识教授、总结巩固、课外自主学习等教学环节设计合理。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score5"" id=""Score5"" class=""weui-count__number Score"" type=""number"" value=""" & HR_CLng(tScore5) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark5"" id=""Remark5"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、广泛使用多媒体、互联网等现代化教学手段进行辅助教学。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score6"" id=""Score6"" class=""weui-count__number Score"" type=""number"" value=""" & HR_CLng(tScore5) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark6"" id=""Remark6"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>三、教学内容</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">符合教学大纲（或课程标准）要求，授课内容正确，重点难点突出，深度与广度适宜，联系实际，例证恰当，适当关注学科进展。（10分） </em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score7"" id=""Score7"" class=""weui-count__number Score"" type=""number"" value=""" & HR_CLng(tScore7) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark7"" id=""Remark7"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>四、教学效果</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、课堂驾驭能力强，师生互动性、课堂纪律、学习气氛好。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score8"" id=""Score8"" class=""weui-count__number Score"" type=""number"" value=""" & HR_CLng(tScore8) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark8"" id=""Remark8"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、完成教学任务，实现教学目的，学生反馈教学效果好。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score9"" id=""Score9"" class=""weui-count__number Score"" type=""number"" value=""" & HR_CLng(tScore9) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark9"" id=""Remark9"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>五、整体评价</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">整体评价（10分） </em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score10"" id=""Score10"" class=""weui-count__number Score"" type=""number"" value=""" & HR_CLng(tScore10) & """ /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	'Response.Write "	<div class=""weui-cell"">" & vbCrlf
	'Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark10"" id=""Remark10"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>总评得分（100分）：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><input name=""TotalScore"" id=""TotalScore"" class=""weui-count__number"" type=""number"" value=""" & HR_CLng(tTotalScore) & """ readonly /></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>六、优点</h3></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Merit"" id=""Merit"" placeholder=""请输入优点"" rows=""2"">" & Trim(tMerit) & "</textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>七、问题与建议</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、意识形态</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Suggest0"" id=""Suggest0"" placeholder=""请输入内容"" rows=""3"">" & Trim(tSuggest0) & "</textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、教学</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Suggest1"" id=""Suggest1"" placeholder=""请输入内容"" rows=""3"">" & Trim(tSuggest1) & "</textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、学风</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Suggest2"" id=""Suggest2"" placeholder=""请输入内容"" rows=""3"">" & Trim(tSuggest2) & "</textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">4、硬件</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Suggest3"" id=""Suggest3"" placeholder=""请输入内容"" rows=""3"">" & Trim(tSuggest3) & "</textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "	<div class=""popbtn""><em class=""weui-btn weui-btn_primary"" id=""subPost"">保存评价</em></div>" & vbCrlf
	If isModify Then Response.Write "	<input name=""Modify"" type=""hidden"" value=""True""><input name=""ID"" type=""hidden"" value=""" & tmpID & """>" & vbCrlf
	Response.Write "</form>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-header-hide""></div>" & vbCrlf
	Response.Write "<div id=""full"" class=""hr-popup"">" & vbCrlf
	Response.Write "	<iframe src=""about:bank"" name=""listFrame"" id=""listFrame"" title=""ListFrame"" width=""100%"" height=""100%"" frameborder=""0""></iframe>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/hrui-touch.js?v=1.0.0""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ location.href=""" & ParmPath & "Evaluate/TeachQuality.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf

	strHtml = strHtml & "	$("".popWin"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		$(""#full"").show();var obj=$(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "		$(""#listFrame"").attr(""src"",""" & ParmPath & "Directories/SelectTeacher.html?Type=3&reObjTxt="" + $(""#""+obj).data(""key"") + ""&reObjValue="" +  $(""#""+obj).data(""value""));" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	var str1 = ""<div class='hr-pop-bg'></div><div class='hr-pop-box'><div class='hr-rows hr-pop-tit'><em class='tit'>请选择</em><em class='close'><i class='hr-icon'>&#xe960;</i></em></div><div class='popbox'><p class='hr-pop-load'><i class='hr-icon'>&#xefe3;</i></p></div></div>"";" & vbCrlf	'提前增加选择层
	strHtml = strHtml & "	$(""body"").append(str1);" & vbCrlf
	strHtml = strHtml & "	var arrItem =[" & GetSelectOptionItem() & "];" & vbCrlf		'业绩项目数据

	strHtml = strHtml & "	$("".itempop"").on(""click"",function(){" & vbCrlf			'选择项目
	strHtml = strHtml & "		var el1=$(this).data(""name""), popw = $("".hr-pop-box"").width(), str1="""";" & vbCrlf
	strHtml = strHtml & "		if(!$(""#Teacher"").val()){ $.toast(""请先选择授课教师！"", ""forbidden""); return false;}" & vbCrlf
	strHtml = strHtml & "		$("".hr-pop-bg"").fadeIn();$("".hr-pop-box"").animate({width:'60%'});" & vbCrlf
	strHtml = strHtml & "		var tid = $(""#Item"").data(""values""), teacher=$(""#TeacherID"").val();" & vbCrlf
	strHtml = strHtml & "		str1+=""<div class='item-list'>"";" & vbCrlf
	strHtml = strHtml & "		for(var i=0;i<arrItem.length;i++){" & vbCrlf
	strHtml = strHtml & "			str1+=""<em data-itemid='"" + arrItem[i].value + ""'>"" + arrItem[i].title + ""</em>"";" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "		str1+=""</div>"";" & vbCrlf
	strHtml = strHtml & "		$("".popbox"").html(str1);" & vbCrlf


	'strHtml = strHtml & "		$.get(""" & ParmPath & "Evaluate/getItemCourse.html"",{ItemID:tid, TeacherID:teacher}, function(strForm){" & vbCrlf
	'strHtml = strHtml & "			$("".popbox"").html(strForm);" & vbCrlf
	'strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		$("".hr-pop-tit .close"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "			$("".hr-pop-bg"").fadeOut();$("".hr-pop-box"").animate({width:'0'});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$("".item-list em"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var itemid = $(this).data(""itemid""), item = $(this).text();" & vbCrlf
	strHtml = strHtml & "			console.log(item);" & vbCrlf
	strHtml = strHtml & "			$(""#ItemID"").val(itemid); $(""#Item"").val(item);" & vbCrlf



	'strHtml = strHtml & "			var cid = $(""#Course"").data(""values""),itemid = $(""#ItemID"").val();" & vbCrlf
	'strHtml = strHtml & "			if(cid==0){ $.toast(""请先选择考核项目！"",  ""forbidden""); };" & vbCrlf
	'strHtml = strHtml & "			$(""#CourseID"").val(cid); console.log(cid);" & vbCrlf
	'strHtml = strHtml & "			$.getJSON(""" & ParmPath & "Evaluate/getCourse.html"",{ID:cid, ItemID:itemid}, function(redata){" & vbCrlf
	'strHtml = strHtml & "				console.log(redata);$(""#Student"").val(redata.Student);" & vbCrlf
	'strHtml = strHtml & "				$(""#Contents"").val(redata.Contents); $(""#CourseDate"").val(redata.CourseDate);" & vbCrlf
	'strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			$("".hr-pop-bg"").fadeOut();$("".hr-pop-box"").animate({width:'0'});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$("".coursepop"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var el1=$(this).data(""name""), popw = $("".hr-pop-box"").width(), str1="""";" & vbCrlf
	strHtml = strHtml & "		var tid = $(""#ItemID"").val(), teacher=$(""#TeacherID"").val();" & vbCrlf
	strHtml = strHtml & "		if(!$(""#ItemID"").val()){ $.toast(""请先选择考核项目！"", ""forbidden""); return false;}" & vbCrlf
	strHtml = strHtml & "		if(!$(""#Teacher"").val()){ $.toast(""请先输入授课教师！"", ""forbidden""); return false;}" & vbCrlf
	strHtml = strHtml & "		$("".hr-pop-bg"").fadeIn();$("".hr-pop-box"").animate({width:'75%'});" & vbCrlf
	strHtml = strHtml & "		$.get(""" & ParmPath & "Evaluate/getItemCourse.html"",{ItemID:tid,TeacherID:teacher}, function(strForm){" & vbCrlf		'取授课教师本项目课程
	strHtml = strHtml & "			var reData = eval(""("" + strForm + "")"").items;" & vbCrlf
	strHtml = strHtml & "			str1+=""<div class='item-list'>"";" & vbCrlf
	strHtml = strHtml & "			for(var i=0;i<reData.length;i++){" & vbCrlf
	strHtml = strHtml & "				str1+=""<em data-id='"" + reData[i].value + ""'>"" + reData[i].title + ""</em>"";" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "			str1+=""</div>"";" & vbCrlf
	strHtml = strHtml & "			$("".popbox"").html(str1);" & vbCrlf
	strHtml = strHtml & "			$("".item-list em"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "				var courseid = $(this).data(""id""), course = $(this).text();" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "Evaluate/getCourse.html"",{ID:courseid, ItemID:tid}, function(redata){" & vbCrlf
	strHtml = strHtml & "					$(""#Course"").val(redata.Course); $(""#CourseID"").val(courseid);" & vbCrlf
	strHtml = strHtml & "					$(""#Student"").val(redata.Student);" & vbCrlf
	strHtml = strHtml & "					$(""#Contents"").val(redata.Contents); $(""#CourseDate"").val(redata.CourseDate);" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "				$("".hr-pop-bg"").fadeOut();$("".hr-pop-box"").animate({width:'0'});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	$(""#Item1"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择"",items:arrItem," & vbCrlf
	strHtml = strHtml & "		onOpen:function(){" & vbCrlf
	strHtml = strHtml & "			if(!$(""#Teacher"").val()){ $.toast(""请先选择授课教师！"", ""forbidden"");}" & vbCrlf
	strHtml = strHtml & "		}," & vbCrlf
	strHtml = strHtml & "		onClose:function(){" & vbCrlf
	strHtml = strHtml & "			var tid = $(""#Item"").data(""values""), teacher=$(""#TeacherID"").val();console.log(tid);" & vbCrlf
	strHtml = strHtml & "			$(""#ItemID"").val(tid);" & vbCrlf
	strHtml = strHtml & "			$.get(""" & ParmPath & "Evaluate/getItemCourse.html"",{ItemID:tid,TeacherID:teacher}, function(strForm){" & vbCrlf
	strHtml = strHtml & "				var reData = eval(""("" + strForm + "")"");" & vbCrlf
	strHtml = strHtml & "				$(""#Course"").select(""update"", reData);" & vbCrlf
	strHtml = strHtml & "				$(""#Course"").val("""");" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#CourseDate"").calendar({dateFormat: 'yyyy年mm月dd日'});" & vbCrlf
	strHtml = strHtml & "	var arrCampus =[" & GetCampusArrData("", 0) & "];" & vbCrlf
	strHtml = strHtml & "	$(""#Campus"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择院区"",items:arrCampus," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf


	strHtml = strHtml & "	var maxNum = 10, minNum = 1;" & vbCrlf
	strHtml = strHtml & "	$('.weui-count__decrease').click(function (e) {" & vbCrlf
	strHtml = strHtml & "		var $input = $(e.currentTarget).parent().find('.weui-count__number');" & vbCrlf
	strHtml = strHtml & "		var number = parseInt($input.val() || ""0"") - 1" & vbCrlf
	strHtml = strHtml & "		if (number < minNum) number = minNum;" & vbCrlf
	strHtml = strHtml & "		$input.val(number);" & vbCrlf
	strHtml = strHtml & "		CountTotalScore(number);" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf
	strHtml = strHtml & "	$('.weui-count__increase').click(function (e) {" & vbCrlf
	strHtml = strHtml & "		var $input = $(e.currentTarget).parent().find('.weui-count__number');" & vbCrlf
	strHtml = strHtml & "		var number = parseInt($input.val() || ""0"") + 1" & vbCrlf
	strHtml = strHtml & "		if (number > maxNum) number = maxNum;" & vbCrlf
	strHtml = strHtml & "		$input.val(number);" & vbCrlf
	strHtml = strHtml & "		CountTotalScore(number);" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf

	strHtml = strHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf
	'strHtml = strHtml & "		var TeacherID = $(""#TeacherID"").val(), ItemID = $(""#ItemID"").val(), CourseID = parseInt($(""#CourseID"").val()), Campus = $(""#Campus"").val();" & vbCrlf
	'strHtml = strHtml & "		if(TeacherID==""""){" & vbCrlf
	'strHtml = strHtml & "			$.toast(""请选择授课教师"", ""cancel"", function(){ return false; });" & vbCrlf
	'strHtml = strHtml & "		}else if(ItemID==""""){" & vbCrlf
	'strHtml = strHtml & "			$.toast(""请选择项目"", ""cancel"", function(){ return false; });" & vbCrlf
	'strHtml = strHtml & "		}else if(CourseID==0){" & vbCrlf
	'strHtml = strHtml & "			$.toast(""请选择课程"", ""cancel"", function(){ return false; });" & vbCrlf
	'strHtml = strHtml & "		}else{" & vbCrlf
	'
	strHtml = strHtml & "			var err1 = false;" & vbCrlf
	strHtml = strHtml & "			$("".verify"").each(function(){" & vbCrlf
	strHtml = strHtml & "				var val1 = $(this).val();" & vbCrlf
	strHtml = strHtml & "				if(!val1){" & vbCrlf
	strHtml = strHtml & "					$(this).addClass(""verifyerr""); $(this).focus();" & vbCrlf
	strHtml = strHtml & "					$.toast( $(this).parent().prev().text() + '不能为空' );" & vbCrlf
	strHtml = strHtml & "					err1=true;" & vbCrlf
	strHtml = strHtml & "				}else{$(this).removeClass(""verifyerr"");}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf

	strHtml = strHtml & "			$("".Score"").each(function(){" & vbCrlf			'判断分值是否为0
	strHtml = strHtml & "				var score = parseInt($(this).val());" & vbCrlf
	strHtml = strHtml & "				if(score==0){" & vbCrlf
	strHtml = strHtml & "					$(this).addClass(""verifyerr""); $(this).focus();" & vbCrlf
	strHtml = strHtml & "					$.toast( '您还有未打分的选项' );" & vbCrlf
	strHtml = strHtml & "					err1=true;" & vbCrlf
	strHtml = strHtml & "				}else{$(this).removeClass(""verifyerr"");}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			if(!err1){" & vbCrlf
	strHtml = strHtml & "				$.post(""" & ParmPath & "Evaluate/SaveQuality.html"", $(""#EditForm"").serialize(), function(rsStr){" & vbCrlf

	strHtml = strHtml & "					$.toast(rsStr.reMessge, function(){ $.closePopup();location.href=""" & ParmPath & "Evaluate/TeachQuality.html""; });" & vbCrlf
	strHtml = strHtml & "				},""json"");" & vbCrlf
	strHtml = strHtml & "			};" & vbCrlf
	
	strHtml = strHtml & "		" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	function CountTotalScore(score){" & vbCrlf
	strHtml = strHtml & "		var total=0;" & vbCrlf
	strHtml = strHtml & "		$("".Score"").each(function(){" & vbCrlf
	strHtml = strHtml & "			total = total + parseInt($(this).val());" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#TotalScore"").val(total);" & vbCrlf
	strHtml = strHtml & "	}" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub SaveQuality()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim isModify : isModify = False

	ErrMsg = ""
	Dim rsSave
	If HR_IsNull(Request("Teacher")) Then ErrMsg = "您没有输入授课教师"
	If HR_IsNull(Request("Course")) Then ErrMsg = "您没有输入课程"
	If HR_IsNull(Request("Student")) Then ErrMsg = "请输入授课对象"
	If HR_IsNull(Request("CourseDate")) Then ErrMsg = "请选择授课程时间"
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """}" : Exit Sub

	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open("Select * From HR_Evaluate Where ID=" & tmpID ), Conn, 1, 3
		If rsSave.BOF And rsSave.EOF Then
			rsSave.AddNew
			rsSave("ID") = GetNewID("HR_Evaluate", "ID")
			rsSave("Participant") = UserYGXM
			rsSave("ParticipantID") = UserYGDM
			rsSave("CreateTime") = Now()
		End If
		rsSave("ItemID") = HR_Clng(Request("ItemID"))
		rsSave("CourseID") = HR_Clng(Request("CourseID"))
		rsSave("Course") = Trim(Request("Course"))
		rsSave("Title") = "课堂教学质量评价"
		rsSave("TeacherID") = HR_Clng(Request("TeacherID"))
		rsSave("Teacher") = Trim(Request("Teacher"))
		rsSave("Campus") = Trim(Request("Campus"))
		rsSave("StuClass") = Trim(Request("Student"))
		rsSave("Contents") = Trim(Request("Contents"))
		rsSave("ClassTime") = Trim(Request("CourseDate"))
		rsSave("Score1") = HR_Clng(Request("Score1"))
		rsSave("Remark1") = Trim(Request("Remark1"))
		rsSave("Score2") = HR_Clng(Request("Score2"))
		rsSave("Remark2") = Trim(Request("Remark2"))
		rsSave("Score3") = HR_Clng(Request("Score3"))
		rsSave("Remark3") = Trim(Request("Remark3"))
		rsSave("Score4") = HR_Clng(Request("Score4"))
		rsSave("Remark4") = Trim(Request("Remark4"))
		rsSave("Score5") = HR_Clng(Request("Score5"))
		rsSave("Remark5") = Trim(Request("Remark5"))
		rsSave("Score6") = HR_Clng(Request("Score6"))
		rsSave("Remark6") = Trim(Request("Remark6"))
		rsSave("Score7") = HR_Clng(Request("Score7"))
		rsSave("Remark7") = Trim(Request("Remark7"))
		rsSave("Score8") = HR_Clng(Request("Score8"))
		rsSave("Remark8") = Trim(Request("Remark8"))
		rsSave("Score9") = HR_Clng(Request("Score9"))
		rsSave("Remark9") = Trim(Request("Remark9"))
		rsSave("Score10") = HR_Clng(Request("Score10"))
		rsSave("Remark10") = Trim(Request("Remark10"))
		rsSave("TotalScore") = HR_Clng(Request("TotalScore"))
		rsSave("Merit") = Trim(Request("Merit"))
		rsSave("Suggest0") = Trim(Request("Suggest0"))
		rsSave("Suggest1") = Trim(Request("Suggest1"))
		rsSave("Suggest2") = Trim(Request("Suggest2"))
		rsSave("Suggest3") = Trim(Request("Suggest3"))
		rsSave.Update
	Set rsSave = Nothing
	ErrMsg = "评价已提交成功！"
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub

Sub ViewQuality()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim isModify : isModify = False

	SiteTitle = "课堂教学质量评价"
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.weui-toast {margin-left: auto;} .weui-textarea{font-size:1rem}" & vbCrlf
	strHtml = strHtml & "		.hr-tips {border-top:1px solid #eee;padding:5px} .hr-tips .tipsIcon {color:#f30;padding-right:5px;font-size:1.5rem;}" & vbCrlf
	strHtml = strHtml & "		.title {border-top:1px solid #eee;padding:10px} .title h3 {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.remarkbar {padding:5px 10px;border-top:1px solid #eee;color:#444;background-color:#f2f2f2}" & vbCrlf
	strHtml = strHtml & "		.popbtn {position:fixed;bottom:0;padding:10px;width:100%;box-sizing: border-box;z-index:12} .popbtn em {width:45%;}" & vbCrlf
	strHtml = strHtml & "		.popbtn em.pass {width:auto;}" & vbCrlf
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

	Dim dataTable, tItemName, tCourse, tStuClass, tContents, tAddress, tClassTime, tPeriod
	Set rsTmp = Conn.Execute("Select * From HR_Evaluate Where ID=" & tmpID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			dataTable = "HR_Sheet_" & rsTmp("ItemID")
			tCourse = Trim(rsTmp("Course"))
			tContents = Trim(rsTmp("Contents"))
			tStuClass = Trim(rsTmp("StuClass"))
			tClassTime = FormatDate(Trim(rsTmp("ClassTime")), 4)
			'tAddress = Trim(rsTmp("TeachAdd"))
			If ChkTable(dataTable) Then
				Set rs = Conn.Execute("Select a.*,b.ClassName,b.Template From " & dataTable & " a Inner Join HR_Class b on a.ItemID=b.ClassID Where a.ID=" & rsTmp("CourseID"))
					If Not(rs.BOF And rs.EOF) Then
						tItemName = rs("ClassName")
						tClassTime = FormatDate(ConvertNumDate(rs("VA4")), 4)
						If rs("Template") = "TempTableA" Then
							tCourse = rs("VA8")
							tContents = rs("VA9")
							tStuClass = rs("VA10")
							tAddress = rs("VA11") & " " & rs("VA12")
							tPeriod = "第" & rs("VA7") & "节 " & GetPeriodTime(rs("VA11"), rs("VA7"), 0)
						Else
							tCourse = rs("VA6")
							tContents = rs("VA7")
						End If
					End If
				Set rs = Nothing
			End If
			Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课教师：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Teacher") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">项目：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & tItemName & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">课程：</label></div><div class=""weui-cell__bd"">" & tCourse & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			'Response.Write "	<div class=""weui-cell"">" & vbCrlf
			'Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">学院：</label></div><div class=""weui-cell__bd"">" & rsTmp("Campus") & "</div>" & vbCrlf
			'Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">课程内容：</label></div><div class=""weui-cell__bd"">" & tContents & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课对象：</label></div><div class=""weui-cell__bd"">" & tStuClass & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课时间：</label></div><div class=""weui-cell__bd"">" & tClassTime & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">节次：</label></div><div class=""weui-cell__bd"">" & tPeriod & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			'Response.Write "	<div class=""weui-cell"">" & vbCrlf
			'Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课地点：</label></div><div class=""weui-cell__bd"">" & tAddress & "</div>" & vbCrlf
			'Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""hr-rows hr-tips"">" & vbCrlf
			Response.Write "		<em class=""tipsIcon""><i class=""hr-icon"">&#xf06a;</i></em>" & vbCrlf
			Response.Write "		<em class=""hr-row-fill tipstxt"">评价标准：>9优/9-6良/<6欠佳</em>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""title""><h3>教学态度与基本技能</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、要求脱稿讲授，语言准确流畅，逻辑性强，富感染力，语速、语调适宜、抑扬顿挫。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score1") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark1") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、精神饱满，教态大方，仪表端正。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score2") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark2") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、PPT设计科学，板书工整，教案讲稿规范。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score3") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark3") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>教学设计与方法</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、运用先进教学理念、方法进行教学，三维目标明确，学情清楚，因材施教，循循善诱。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score4") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark4") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、教学设计科学，新课导入、知识教授、总结巩固、课外自主学习等教学环节设计合理。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score5") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark5") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、广泛使用多媒体、互联网等现代化教学手段进行辅助教学。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score6") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark6") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>教学内容</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">符合教学大纲（或课程标准）要求，授课内容正确，重点难点突出，深度与广度适宜，联系实际，例证恰当，适当关注学科进展。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score7") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark7") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>教学效果</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、课堂驾驭能力强，师生互动性、课堂纪律、学习气氛好。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score8") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark8") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、完成教学任务，实现教学目的，学生反馈教学效果好。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score9") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark9") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>整体评价</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">整体评价。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score10") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark10") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">总评得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("TotalScore") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>优点</h3></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Merit") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>问题与建议</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、意识形态</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Suggest0") & "</div></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、教学</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Suggest1") & "</div></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、学风</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Suggest2") & "</div></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">4、硬件</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Suggest3") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">听课人：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Participant") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">评价时间：" & FormatDate(rsTmp("CreateTime"), 1) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			If HR_CBool(rsTmp("Passed")) Then
				Response.Write "	<div class=""popbtn"">" & vbCrlf
				Response.Write "		<em class=""pass""><a href=""" & ParmPath & "Evaluate/ApplyModify.html?ID=" & rsTmp("ID") & """ class=""weui-btn weui-btn_primary"" id=""subApply"">申请修改</a></em>" & vbCrlf
				Response.Write "	</div>" & vbCrlf
			Else
				Response.Write "	<div class=""hr-rows popbtn""><em><a href=""" & ParmPath & "Evaluate/EditQuality.html?ID=" & rsTmp("ID") & "&Modify=True"" class=""weui-btn weui-btn_primary"" id=""subPost"">修改评价</a></em>" & vbCrlf
				Response.Write "		<em><a href=""" & ParmPath & "Evaluate/Apply.html?ID=" & rsTmp("ID") & """ class=""weui-btn weui-btn_primary"" id=""subApply"">提交评价</a></em>" & vbCrlf
				Response.Write "	</div>" & vbCrlf
			End If
			Response.Write "</div>" & vbCrlf
			Response.Write "<div class=""hr-header-hide""></div>" & vbCrlf
		End If
	Set rsTmp = Nothing
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf

	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub getItemCourse()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tTeacherID : tTeacherID = HR_Clng(Request("TeacherID"))
	Dim tCourseID : tCourseID = HR_Clng(Request("ID"))
	Dim tTableName : tTableName = "HR_Sheet_" & tItemID
	Dim strTmp
	strTmp = strTmp & "{""items"":["
	If ChkTable(tTableName) Then
		sql = "Select a.*,b.ClassName,b.Template From " & tTableName & " a Inner Join HR_Class b on a.ItemID=b.ClassID Where a.ID>0 And a.scYear=" & DefYear
		If tTeacherID>0 Then sql = sql & " And VA1=" & tTeacherID
		If tCourseID>0 Then sql = sql & " And ID=" & tCourseID

		Set rs = Conn.Execute(sql)
			If Not(rs.BOF And rs.EOF) Then
				i = 0
				Do While Not rs.EOF
					If i>0 Then strTmp = strTmp & ","
					strTmp = strTmp & "{""title"":"""
					If rs("Template") = "TempTableA" Then
						strTmp = strTmp & " " & rs("VA8") & "_" & rs("VA7") & "节"
					Else
						strTmp = strTmp & " " & rs("VA6")
					End If
					strTmp = strTmp & " " & FormatDate(ConvertNumDate(rs("VA4")), 2) & """,""value"":""" & rs("ID") & """,""Data1"":""" & rs("VA4") & """}"
					rs.MoveNext
					i = i + 1
				Loop
			Else
				strTmp = strTmp & "{""title"":""该教师在本项目中没有课程"",""value"":""0""}"
			End If
		Set rs = Nothing
	End If
	strTmp = strTmp & "]}"
	Response.Write strTmp
End Sub
Sub getCourse()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tCourseID : tCourseID = HR_Clng(Request("ID"))
	Dim tTableName : tTableName = "HR_Sheet_" & tItemID
	Dim strTmp
	If ChkTable(tTableName) Then
		sql = "Select a.*,b.ClassName,b.Template From " & tTableName & " a Inner Join HR_Class b on a.ItemID=b.ClassID Where a.ID>0 And a.scYear=" & DefYear
		sql = sql & " And a.ID=" & tCourseID
		Set rs = Conn.Execute(sql)
			If Not(rs.BOF And rs.EOF) Then
				strTmp = """err"":false,""errcode"":0,""errmsg"":"""",""CourseDate"":""" & FormatDate(ConvertNumDate(rs("VA4")), 4) & """"
				If rs("Template") = "TempTableA" Then
					strTmp = strTmp & ",""Course"":""" & rs("VA8") & """,""Student"":""" & FilterHtmlToText(rs("VA10")) & """,""Period"":""第" & rs("VA7") & "节"""
					strTmp = strTmp & ",""Contents"":""" & FilterHtmlToText(rs("VA9")) & """"
				Else
					strTmp = strTmp & ",""Course"":""" & rs("VA6")& """"
				End If
				strTmp = strTmp & ",""CourseID"":" & HR_CLng(rs("ID")) & ",""ItemID"":" & HR_CLng(rs("ItemID")) & ""
			Else
				strTmp = """err"":true,""errcode"":500,""errmsg"":""该教师在本项目中没有课程"""
			End If
		Set rs = Nothing
	End If
	Response.Write "{" & strTmp & "}"
End Sub

Sub Apply()
	Dim rsSave, tmpID : tmpID = HR_Clng(Request("ID"))
	ErrMsg = ""
	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open("Select * From HR_Evaluate Where ID=" & tmpID ), Conn, 1, 3
		If rsSave.BOF And rsSave.EOF Then
			ErrMsg = "课堂教学质量评价没找到，可能已经删除！"
		Else
			rsSave("Passed") = HR_True
			rsSave.Update
			ErrMsg = "课堂教学质量评价已经提交成功！"
		End If
	Set rsSave = Nothing

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background-color:#f1f1f1;} .error{margin:0;width:100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-err-box {position:fixed;left:20px;top:15%;right:20px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-err-box .error {width:100%;padding:20px;box-sizing:border-box;border:1px solid #ccc;background-color:rgba(255,255,255,0.8);}" & vbCrlf
	tmpHtml = tmpHtml & "		.error .errorInfo {flex-grow:2;padding: 0 10px;font-size:1rem;} .error .errorInfo h2 {font-size:1.3rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.error .errorIcon i {font-size:4rem;color:#080;}" & vbCrlf
	tmpHtml = tmpHtml & "		.reindex {padding-top:20px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	strHtml = strHtml & getHeadNav(0)
	strHtml = strHtml & "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	strHtml = strHtml & "<div class=""hr-err-box"">" & vbCrlf
	strHtml = strHtml & "	<div class=""hr-rows error"">" & vbCrlf
	strHtml = strHtml & "		<div class=""errorIcon""><i class=""hr-icon"">&#xe813;</i></div>" & vbCrlf
	strHtml = strHtml & "		<div class=""errorInfo"">" & vbCrlf
	strHtml = strHtml & "			<h2>操作成功！</h2>" & vbCrlf
	strHtml = strHtml & "			<p>" & ErrMsg & "</p>" & vbCrlf
	strHtml = strHtml & "		</div>" & vbCrlf
	strHtml = strHtml & "	</div>" & vbCrlf
	strHtml = strHtml & "	<div class=""reindex""><a class=""weui-btn weui-btn_primary"" href=""" & ParmPath & "Evaluate/TeachQuality.html"">返回</a></div>" & vbCrlf
	strHtml = strHtml & "</div>" & vbCrlf
	strHtml = strHtml & getPageFoot(1)

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub ApplyModify()
	SiteTitle = "评价修改申请"
	Dim rsSave, tReason, tmpID : tmpID = HR_Clng(Request("ID"))
	ErrMsg = ""
	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open("Select * From HR_Evaluate Where ID=" & tmpID ), Conn, 1, 1
		If rsSave.BOF And rsSave.EOF Then
			ErrMsg = "课堂教学质量评价没找到，可能已经删除！"
			Response.Write GetErrBody(0) : Exit Sub
		End If
	Set rsSave = Nothing

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background-color:#fff;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-err-box {padding:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.reindex {padding-top:20px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.remarkbar {padding:5px 10px;border-bottom:1px solid #ddd;color:#444;background-color:#fff}" & vbCrlf
	tmpHtml = tmpHtml & "		.reasonbox {border:1px solid #eee;margin-top:15px;} .reasonbox textarea {border:0;padding:10px;font-size:1.2rem;width:100%;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	strHtml = strHtml & "<header class=""hr-rows hr-header"">" & vbCrlf
	strHtml = strHtml & "	<nav class=""navBack""><em><i class=""hr-icon"">&#xf320;</i></em></nav>" & vbCrlf
	strHtml = strHtml & "	<nav class=""navTitle""><span>" & SiteTitle & "</span></nav>" & vbCrlf
	strHtml = strHtml & "	<nav class=""navMenu""><em><i class=""hr-icon"">&#xf32a;</i></em></nav>" & vbCrlf
	strHtml = strHtml & "</header>" & vbCrlf
	strHtml = strHtml & "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	strHtml = strHtml & "<div class=""hr-err-box"">" & vbCrlf

	strHtml = strHtml & "	<div class=""remarkbar""><em class=""remarktxt"">修改理由</em></div>" & vbCrlf
	strHtml = strHtml & "	<div class=""reasonbar"">" & vbCrlf
	strHtml = strHtml & "		<div class=""reasonbox""><textarea name=""Reason"" id=""Reason"" placeholder=""请输入内容"" rows=""5"">" & tReason & "</textarea></div>" & vbCrlf
	strHtml = strHtml & "	</div>" & vbCrlf

	strHtml = strHtml & "	<div class=""reindex""><a class=""weui-btn weui-btn_primary"" id=""subpost"" href=""javascript:;"">提交申请</a></div>" & vbCrlf
	strHtml = strHtml & "</div>" & vbCrlf
	strHtml = strHtml & getPageFoot(1)

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ location.href=""" & ParmPath & "Evaluate/TeachQuality.html"";  });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#subpost"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var reason = $(""#Reason"").val();" & vbCrlf
	tmpHtml = tmpHtml & "		if(!reason){$.toast(""请输入修改理由"", ""cancel"", function(){ return false; });};" & vbCrlf
	tmpHtml = tmpHtml & "		$.post(""" & ParmPath & "Evaluate/SaveApply.html"",{ID:" & tmpID & ",Reason:reason}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "			if(strForm.err){" & vbCrlf
	tmpHtml = tmpHtml & "				$.toast(strForm.errmsg, ""cancel"");" & vbCrlf
	tmpHtml = tmpHtml & "			}else{" & vbCrlf
	tmpHtml = tmpHtml & "				$.toast(strForm.errmsg, function(){ location.href=""" & ParmPath & "Evaluate/TeachQuality.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		},""json"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub SaveApply()
	Dim rsSave, tReason, tmpID : tmpID = HR_Clng(Request("ID"))
	tReason = Trim(Request("Reason"))
	ErrMsg = ""
	If HR_IsNull(tReason) Then ErrMsg = "您还没有输入修改理由"
	Set rsSave = Conn.Execute("Select * From HR_Evaluate Where ParticipantID=" & UserYGDM & " And ID=" & tmpID )
		If rsSave.BOF And rsSave.EOF Then
			ErrMsg = "课堂教学质量评价没找到，可能已经删除！"
		End If
	Set rsSave = Nothing
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""err"":true,""errcode"":500,""errmsg"":""" & ErrMsg & """,""icon"":2}" : Exit Sub

	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open("Select * From HR_Apply Where Module='Evaluate' And YGDM=" & UserYGDM & " And RelateID=" & tmpID ), Conn, 1, 3
		If rsSave.BOF And rsSave.EOF Then
			rsSave.AddNew
			rsSave("ID") = GetNewID("HR_Apply", "ID")
			rsSave("YGDM") = UserYGDM
			rsSave("Module") = "Evaluate"
			rsSave("ItemID") = 0
			rsSave("RelateID") = tmpID
			rsSave("CreateTime") = Now()
		End If
		rsSave("Reason") = tReason
		rsSave.Update
	Set rsSave = Nothing
	Response.Write "{""err"":false,""errcode"":0,""errmsg"":""您的申请已经提交成功！"",""icon"":1}"
End Sub

Function GetCampusArrData(fCampus, fType)		'取校院区数据
	Dim strFun, iFun, fArrCampus : fArrCampus = Split(XmlText("Common", "Campus", ""), "|")
	For iFun = 0 To Ubound(fArrCampus)
		If iFun > 0 Then strFun = strFun & ","
		strFun = strFun & """" & fArrCampus(iFun) & """"
	Next
	GetCampusArrData = strFun
End Function

Function GetSelectOptionItem()				'取考核项目下拉
	Dim iFun, funItem, rsFun, sqlFun
	sqlFun = "Select ClassID, ClassName From HR_Class Where ModuleID=1001 And Child=0 And Template in('TempTableA','TempTableC')"
	sqlFun = sqlFun & " Order By RootID, OrderID"
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then funItem = funItem & ","
				funItem = funItem & "{title:""" & rsFun("ClassName") & """,value:""" & rsFun("ClassID") & """}"
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	GetSelectOptionItem = funItem
End Function
%>