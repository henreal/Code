<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./m_ManageCommon.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl : SiteTitle = "调换课程管理"
Dim arrProcess : arrProcess = Split("待审,审核中,通过审核,未批准", ",")

Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "Pass" Call PassBody()
	Case "View" Call View()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim Count1, tAssort		'汇总
	tAssort = HR_CLng(Request("Assort"))
	Set rsTmp = Conn.Execute("Select count(0) From HR_Swap where newCourseID>0")
		Count1 = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing
	If tAssort = 2 Then
		Set rsTmp = Conn.Execute("Select count(0) From HR_Swap where newCourseID=0")
			Count1 = HR_Clng(rsTmp(0))
		Set rsTmp = Nothing
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background:#fff;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-tabs {display:flex;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-tabs a {width:50vw;text-align:center;line-height: 45px; box-sizing: border-box;border-right:1px solid #eee;background:#ddd;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-tabs a:last-child {border-right:0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-tabs a.this {border-top:1px solid #f90;background:transparent;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-tips_wrap {border:1px solid #f90;border-radius:5px; margin:10px 5px; padding:5px; color:#999}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	tmpHtml = tmpHtml & "		.title {font-size:1rem;font-weight: bold;} .tips {font-size:0.8rem;color:#999;}" & vbCrlf
	tmpHtml = tmpHtml & "		.cell_href {border-bottom:8px solid #e3e3e3;padding:10px 5px 10px 10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-cell_item {display:flex; color:#999; padding:5px 0 0 0;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", tmpHtml)
	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ FastClick.attach(document.body); });" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-tabs"">" & vbCrlf
	If tAssort = 2 Then
		Response.Write "	<a href=""" & ParmPath & "ManageSwap/Index.html?Assort=1"">换课申请</a>" & vbCrlf
		Response.Write "	<a class=""this"" href=""" & ParmPath & "ManageSwap/Index.html?Assort=2"">代课申请</a>" & vbCrlf
	Else
		Response.Write "	<a class=""this"" href=""" & ParmPath & "ManageSwap/Index.html?Assort=1"">换课申请</a>" & vbCrlf
		Response.Write "	<a href=""" & ParmPath & "ManageSwap/Index.html?Assort=2"">代课申请</a>" & vbCrlf
	End If
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-rows hr-tips_wrap"">" & vbCrlf
	Response.Write "	<div class=""hr-row_left"">申请数：</div>" & vbCrlf
	Response.Write "	<div class=""hr-row_num"">" & Count1 & " 人次</div>"  & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-cells"">" & vbCrlf
	Dim tTitle, tSheetName, tTemplate, tCourseDate
	sqlTmp = "Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.YGDM) As Proposer,(Select Template From HR_Class Where ClassID=a.ItemID) As Template,(Select ClassName From HR_Class Where ClassID=a.ItemID) As ItemName From HR_Swap a Where a.ID>0"
	If tAssort = 2 Then
		sqlTmp = sqlTmp & " And a.newCourseID=0"
	Else
		sqlTmp = sqlTmp & " And a.newCourseID>0"
	End If
	sqlTmp = sqlTmp & " Order By a.ApplyTime DESC"
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				tSheetName = "HR_Sheet_" & rsTmp("ItemID")		'数据表名
				tTemplate = Trim(rsTmp("Template"))		'数据模型
				If ChkTable(tSheetName) Then
					Set rs = Conn.Execute("Select * From " & tSheetName & " Where ID=" & rsTmp("CourseID"))
						If Not(rs.BOF And rs.EOF) Then
							If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Then
								tCourseDate = FormatDate(ConvertNumDate(rs("VA4")), 2)
							Else
								tCourseDate = Trim(rs("VA4"))
							End If
						End If
					Set rs = Nothing
					tTitle = "【代课】" & Trim(rsTmp("ItemName")) & "" & tCourseDate & "的课程"
					If HR_CLng(rsTmp("newCourseID")) > 0 Then tTitle = "【换课】" & Trim(rsTmp("ItemName")) & "" & tCourseDate & "的课程"
					Response.Write "	<a class=""hr-rows hr-cell cell_href"" href=""" & ParmPath & "ManageSwap/Pass.html?ID=" & rsTmp("ID") & """>" & vbCrlf
					Response.Write "		<div class=""hr-wrap"" data-id=""" & rsTmp("ID") & """><h3 class=""title"">" & tTitle & "</h3>" & vbCrlf
					Response.Write "			<dl class=""hr-cell_item""><dt>申请课程：</dt><dd>" & FormatDate(rsTmp("VA4"), 4) & " 第" & rsTmp("VA7") & "节</dd></dl>" & vbCrlf
					Response.Write "			<dl class=""hr-cell_item""><dt>申请人：</dt><dd>" & rsTmp("Proposer") & " " & FormatDate(rsTmp("ApplyTime"), 10) & "</dd></dl>" & vbCrlf
					Response.Write "		</div>" & vbCrlf
					Response.Write "		<div class=""hr-cell_more""><i class=""hr-icon"">&#xf321;</i></div>" & vbCrlf
					Response.Write "	</a>" & vbCrlf
				Else
					Response.Write "	<p class=""hr-list-error"">数据表不存在！</p>" & vbCrlf
				End If
				rsTmp.MoveNext
			Loop
		Else
			Response.Write "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>暂时没有调换课申请！</h3></div>" & vbCrlf
		End If
	Set rsTmp = Nothing
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Manage/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

Sub PassBody()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	SiteTitle = "调换课详情"

	Dim tProposer, tReason, tReplacer, tPasser, tPassTime, tPasser1, tPassTime1, tPasser2, tPassTime2, tProcess, tExplain
	Dim tApplyTime, tCourse, tStuClass, tPlace, tCourseDate, tCourseTime, tPeriod
	Dim tSheetName, tTemplate, tItemName, tCourseID, tTeachDate
	sqlTmp = "Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.YGDM) As Proposer"
	sqlTmp = sqlTmp & ",(Select YGXM From HR_Teacher Where YGDM=a.Replacer) As Replacer"
	sqlTmp = sqlTmp & ",(Select YGXM From HR_Teacher Where YGDM=a.Passer) As PasserName"
	sqlTmp = sqlTmp & ",(Select YGXM From HR_Teacher Where YGDM=a.Passer1) As PasserName1"
	sqlTmp = sqlTmp & ",(Select YGXM From HR_Teacher Where YGDM=a.Passer2) As PasserName2"
	sqlTmp = sqlTmp & ",(Select Template From HR_Class Where ClassID=a.ItemID) As Template,(Select ClassName From HR_Class Where ClassID=a.ItemID) As ItemName From HR_Swap a Where a.ID=" & tmpID
	Set rs = Conn.Execute(sqlTmp)
		If Not(rs.BOF And rs.EOF) Then
			tSheetName = "HR_Sheet_" & rs("ItemID")		'数据表名
			tTemplate = rs("Template")
			tItemName = rs("ItemName")
			tProposer = Trim(rs("Proposer"))	'申请人
			tReason = Trim(rs("Reason"))	'申请理由
			tReplacer = Trim(rs("Replacer"))	'替换教师
			tPasser = Trim(rs("PasserName"))	'教务主任
			tPasser1 = Trim(rs("PasserName1"))	'教学处审核人
			tPasser2 = Trim(rs("PasserName2"))	'教辅审核人
			tPassTime = FormatDate(rs("PassTime"), 10)	'教务主任审核时间
			tPassTime1 = FormatDate(rs("PassTime1"), 10)	'教务主任审核时间
			tPassTime2 = FormatDate(rs("PassTime2"), 10)	'教务主任审核时间
			tExplain = Trim(rs("Explain"))	'审核说明
			tApplyTime = FormatDate(rs("ApplyTime"), 1)	'申请提交时间
			tCourseID = HR_Clng(rs("CourseID"))
		End If
	Set rs = Nothing
	If ChkTable(tSheetName) Then
		sql = "Select a.* From " & tSheetName & " a Where a.ID=" & tCourseID
		Set rs = Conn.Execute(sql)
			If Not(rs.BOF And rs.EOF) Then
				If tTemplate = "TempTableA" Then
					tCourse = rs("VA8")
					tStuClass = rs("VA10")
					tPlace = rs("VA11") & " " & rs("VA12")
					tTeachDate = FormatDate(ConvertNumDate(rs("VA4")), 4) & " 第" & Trim(rs("VA7")) & "节"
				Else
					tCourse = rs("VA6")
				End If
			End If
		Set rs = Nothing
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dl {align-items:stretch;border-bottom:1px solid #eee;padding:5px 0;font-size:1.2rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dt {width:30%;text-align:right;color:#999}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dd {flex-grow:2;width:70%;box-sizing: border-box;padding-right:3px}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", tmpHtml)
	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ FastClick.attach(document.body); });" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-swap-box"">" & vbCrlf
	Response.Write "	<div class=""hr-swap-items"">" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>申请人：</dt><dd>" & tProposer & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>申请时间：</dt><dd>" & tApplyTime & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>原　因：</dt><dd>" & tReason & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>替课教师：</dt><dd>" & tReplacer & "</dd></dl>" & vbCrlf
	Response.Write "		<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>项目名称：</dt><dd>" & tItemName & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>授课时间：</dt><dd>" & tTeachDate & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>课程名称：</dt><dd>" & tCourse & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>授课对象：</dt><dd>" & tStuClass & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>授课地点：</dt><dd>" & tPlace & "</dd></dl>" & vbCrlf
	Response.Write "		<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>教研主任：</dt><dd>" & tPasser & "<br>"
	If HR_IsNull(tPassTime) Then
		Response.Write "[未审]" & vbCrlf
	Else
		Response.Write "[已审]" & tPassTime & vbCrlf
	End If
	Response.Write "			</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>教学处：</dt><dd>" & tPasser1 & "<br>"
	If HR_IsNull(tPassTime1) Then
		Response.Write "[未审]" & vbCrlf
	Else
		Response.Write "[已审]" & tPassTime1 & vbCrlf
	End If
	Response.Write "			</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>教辅：</dt><dd>" & tPasser2 & "<br>"
	If HR_IsNull(tPassTime2) Then
		Response.Write "[未审]" & vbCrlf
	Else
		Response.Write "[已审]" & tPassTime2 & vbCrlf
	End If
	Response.Write "			</dd></dl>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	'Response.Write "	<div class=""hr-rows hr-editbtn"">" & vbCrlf
	'Response.Write "		<em><i class=""hr-icon"">&#xea3f;</i></em>" & vbCrlf
	'Response.Write "		<em><button type=""button"" name=""passed"" class=""passed"" id=""passed"" data-id=""" & tmpID & """><i class=""hr-icon"">&#xf00c;</i>通过</button></em>" & vbCrlf
	'Response.Write "		<em><button type=""button"" name=""fail"" class=""fail"" id=""fail"" data-id=""" & tmpID & """><i class=""hr-icon"">&#xe14b;</i>拒绝</button></em>" & vbCrlf
	'Response.Write "		<em><button type=""button"" name=""delete"" class=""delete"" id=""delete"" data-id=""" & tmpID & """><i class=""hr-icon"">&#xe872;</i>删除</button></em>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-fix backbtn""><button type=""button"" name=""back"" class=""weui-btn weui-btn_plain-default"" id=""back"">返回</button></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/js/swiper.min.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "ManageSwap/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".passed"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		$.toast(""审核通过！"");" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$("".fail"").on(""click"",function(){ $.toast(""拒绝申请！""); });" & vbCrlf
	strHtml = strHtml & "	$(""#back"").on(""click"",function(){ location.href=""" & ParmPath & "ManageSwap/Index.html""; });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

%>