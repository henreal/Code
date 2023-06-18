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
Dim scriptCtrl : SiteTitle = "对换课申请"
If ChkWechatTokenQY() = False Then Call GetWechatTokenQY()		'提前检查企业微信Token是否过期
If ChkTokenBobao() = False Then Call GetTokenBobao()			'检查信息播报Token

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "Details","AlternDetails" Call Details()
	Case "Edit" Call Edit()
	Case "SavePost" Call SavePost()
	Case "Delete" Call Delete()

	Case "Alternate" Call Alternate()
	Case "Agree" Call Agree()
	Case "getItemCourse" Call getItemCourse()
	Case "getCourse" Call getCourse()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background:#fff;}" & vbCrlf
	tmpHtml = tmpHtml & "		.listBar {padding:10px 5px;border-bottom:1px solid #eee;}" & vbCrlf
	tmpHtml = tmpHtml & "		.listBar .icon {padding-right:5px;font-size:26px;color:#3491c6;}" & vbCrlf
	tmpHtml = tmpHtml & "		.processbar {padding:5px 5px; border-bottom:10px solid #eee;}" & vbCrlf	
	tmpHtml = tmpHtml & "		.weui-cell {padding:5px 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-float-btn {width:55px;height:55px;text-align:center;font-size:2.6rem;position: fixed;right: 1rem;bottom: 3rem;z-index: 99;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-float-btn i {color:#2196f3;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-list-title {margin:5px;padding:0 10px;box-sizing: border-box;line-height:35px;background:#ffe596;color:#900;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-list-ul .iconTit {color:#2196f3;padding-right:5px;font-size:2rem;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim tCourseList, tStep, tPeriodTime, tReplacer
	sqlTmp = "Select a.*,(Select ClassName From HR_CLass Where ClassID=a.ItemID) As ItemName,(Select YGXM From HR_Teacher Where YGDM=a.Replacer) As ReplacerXM From HR_Swap a"
	sqlTmp = sqlTmp & " Where a.newItemID>0 And a.newCourseID>0 And a.YGDM=" & UserYGDM		'取课程SQL（注意代课和换课分离）
	sqlTmp = sqlTmp & " Order By a.ApplyTime DESC"
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			tCourseList = tCourseList & "<div class=""hr-list-ul"">"
			Do While Not rsTmp.EOF
				tStep = PassProcess(0, 0)
				If HR_CLng(rsTmp("Process")) = 1 Then tStep = "换课教师" & PassProcess(1, HR_CLng(rsTmp("ReplacePass")))
				If HR_CLng(rsTmp("Process")) = 2 Then tStep = "教研主任" & PassProcess(2, HR_CLng(rsTmp("PasserPass")))
				If HR_CLng(rsTmp("Process")) = 3 Then tStep = "教学处" & PassProcess(3, HR_CLng(rsTmp("Passer1Pass")))
				If HR_CLng(rsTmp("Process")) = 4 Then tStep = "教辅" & PassProcess(4, HR_CLng(rsTmp("Passer2Pass")))

				tReplacer = Trim(rsTmp("ReplacerXM"))
				tCourseList = tCourseList & "	<a class=""hr-rows hr-stretch listBar"" href=""" & ParmPath & "SwapCourse/Details.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """>" & vbCrlf
				tCourseList = tCourseList & "		<div class=""icon""><i class=""hr-icon"">&#xf33c;</i></div>" & vbCrlf
				tCourseList = tCourseList & "		<div class=""hr-grow listinfo"" data-id=""" & rsTmp("CourseID") & """>" & vbCrlf
				tCourseList = tCourseList & "			<h3>我的课程：" & FormatDate(Trim(rsTmp("VA4")), 4) & "</h3><em>" & GetPeriodTime(rsTmp("VA11"), rsTmp("VA7"),0) & " 第" & rsTmp("VA7") & "节</em>" & vbCrlf
				tCourseList = tCourseList & "			<em>" & rsTmp("VA8") & " " & rsTmp("VA11") & "</em>" & vbCrlf
				tCourseList = tCourseList & "			<em>换课教师：<span>" & tReplacer & "</span></em>" & vbCrlf
				tCourseList = tCourseList & "			<em>换课课程：<span>" & FormatDate(Trim(rsTmp("VA4")), 4) & "第" & rsTmp("VA7") & "节</span></em>" & vbCrlf
				tCourseList = tCourseList & "		</div>" & vbCrlf
				tCourseList = tCourseList & "	</a>" & vbCrlf
				tCourseList = tCourseList & "	<a class=""hr-rows processbar"" href=""" & ParmPath & "SwapCourse/Details.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """>" & vbCrlf
				tCourseList = tCourseList & "		<em>审核：" & tStep & "</em><tt class=""more""><i class=""hr-icon"">&#xf321;</i></tt>" & vbCrlf
				tCourseList = tCourseList & "	</a>" & vbCrlf
				rsTmp.MoveNext
				i = i + 1
			Loop
			tCourseList = tCourseList & "</div>"
			tCourseList = "<h3 class=""hr-list-title""><i class=""hr-icon"">&#xe972;</i>您共有<b>" & i & "</b>条换课申请！</h3>" & tCourseList
		Else
			tCourseList = tCourseList & "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef61;</i></h2><h3>您暂时还没有申请对换课程！</h3></div>" & vbCrlf
		End If
	Set rsTmp = Nothing

	Response.Write "<div class=""weui-cells"">" & tCourseList & "</div>" & vbCrlf
	Response.Write "<div class=""hr-float-btn"">" & vbCrlf
	Response.Write "	<a href=""" & ParmPath & "SwapCourse/Edit.html?AddNew=True"" class=""addBtn""><i class=""hr-icon"">&#xf3c0;</i></a>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").html(""<i class='hr-icon'>&#xf0ca;</i>"");" & vbCrlf		
	tmpHtml = tmpHtml & "	$("".navBack em"").html(""<i class='hr-icon'>&#xf320;</i>"");" & vbCrlf	
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf

	tmpHtml = tmpHtml & "	var laynav=""<li><a href=\""" & ParmPath & "SwapCourse/Index.html\""><i class=\""hr-icon hr-icon-top\"">&#xe853;</i>我申请的换课</a></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "	laynav+=""<li><a href=\""" & ParmPath & "SwapCourse/Alternate.html\""><i class=\""hr-icon hr-icon-top\"">&#xf2dd;</i>我的换课记录</a></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "	$("".nctouch-nav-menu ul"").html(laynav);" & vbCrlf				'更新右上角导航

	tmpHtml = tmpHtml & "</script>" & vbCrlf
	
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub Details()
	SiteTitle = "换课申请详情"
	If Action = "AlternDetails" Then SiteTitle = "换课记录详情"
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim SwapPass : SwapPass = HR_CLng(GetTypeName("HR_User", "SwapPass", "YGDM", UserYGDM))		'判断教学处或教辅审核

	Dim tVA2, tVA4, tCourse, tItem, tReason, tReplacerID, tReplacer, tPasser, tPasserID, tPassTime, tPasserPass
	Dim tCourseDate, tPeriod, tPeriodTime, tPlace, tStuClass, NotModi, tProcess, strProcess
	Dim tCourse1, tCourseDate1, tPeriod1, tPeriodTime1, tPlace1, tStuClass1
	Dim tApplyID, tApplyer, tApplyerKS, tApplyerZW, tApplyerZC, tApplyTime
	Dim tReplacerKS, tReplacerZW, tReplacerZC, tReplacerTime, tReplacePass
	Dim tPasser1, tPasser2, tPassTime1, tPassTime2, tPasser1Pass, tPasser2Pass
	NotModi = False
	sqlTmp = "Select * From HR_Swap Where ItemID=" & tItemID & " And ID=" & tmpID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tReason = Trim(rsTmp("Reason"))
			tProcess = HR_Clng(rsTmp("Process"))
			tApplyID = HR_Clng(rsTmp("YGDM"))
			tApplyer = strGetTypeName("HR_Teacher", "YGXM", "YGDM", tApplyID)		'申请老师
			tApplyerKS = strGetTypeName("HR_Teacher", "KSMC", "YGDM", tApplyID)
			tApplyerZW = strGetTypeName("HR_Teacher", "XZZW", "YGDM", tApplyID)
			tApplyerZC = strGetTypeName("HR_Teacher", "PRZC", "YGDM", tApplyID)
			tApplyTime = FormatDate(rsTmp("ApplyTime"), 10)

			tReplacerID = HR_Clng(rsTmp("Replacer"))
			tReplacerTime = FormatDate(rsTmp("ReplacerTime"), 10)
			tReplacer = strGetTypeName("HR_Teacher", "YGXM", "YGDM", tReplacerID)		'替课老师
			tReplacePass = HR_Clng(rsTmp("ReplacePass"))								'替课老师确认状态
			tReplacerKS = strGetTypeName("HR_Teacher", "KSMC", "YGDM", tReplacerID)
			tReplacerZW = strGetTypeName("HR_Teacher", "XZZW", "YGDM", tReplacerID)
			tReplacerZC = strGetTypeName("HR_Teacher", "PRZC", "YGDM", tReplacerID)

			tPasserID = HR_Clng(rsTmp("Passer"))
			tPassTime = FormatDate(rsTmp("PassTime"), 10)								'教研主任审核时间
			tPasser = strGetTypeName("HR_Teacher", "YGXM", "YGDM", tPasserID)			'教研室主任
			tPasserPass = HR_Clng(rsTmp("PasserPass"))									'教研主任审核状态

			tPasser1 = HR_Clng(rsTmp("Passer1"))
			tPassTime1 = FormatDate(rsTmp("PassTime1"), 10)								'教学处审核时间
			tPasser1Pass = HR_Clng(rsTmp("Passer1Pass"))								'教学处审核状态

			tPasser2 = HR_Clng(rsTmp("Passer2"))
			tPassTime2 = FormatDate(rsTmp("PassTime2"), 10)								'教辅审核时间
			tPasser2Pass = HR_Clng(rsTmp("Passer2Pass"))								'教辅审核状态

			tReason = Trim(rsTmp("Reason"))
			tCourseDate = FormatDate(Trim(rsTmp("VA4")), 4)
			tCourse = Trim(rsTmp("VA8"))
			tPeriod = " 第" & Trim(rsTmp("VA7")) & "节"
			tPeriodTime = "" & GetPeriodTime(rsTmp("VA11"), rsTmp("VA7"), 1) & ""
			tPlace = Trim(rsTmp("VA11")) & "" & Trim(rsTmp("VA12"))
			tStuClass = Trim(rsTmp("VA10"))

			tCourseDate1 = FormatDate(Trim(rsTmp("newVA4")), 4)
			tCourse1 = Trim(rsTmp("newVA8"))
			tPeriod1 = " 第" & Trim(rsTmp("newVA7")) & "节"
			tPeriodTime1 = "" & GetPeriodTime(rsTmp("newVA11"), rsTmp("newVA7"), 1) & ""
			tPlace1 = Trim(rsTmp("newVA11")) & "" & Trim(rsTmp("newVA12"))
			tStuClass1 = Trim(rsTmp("newVA10"))

		Else
			Response.Redirect ParmPath & "Swap/Index.html"
		End If
	Set rsTmp = Nothing

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background-color:#fff}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dl {align-items:stretch;border-bottom:1px solid #eee;padding:5px 0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dt {width:30%;text-align:right;color:#999}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dd {flex-grow:2;width:70%;box-sizing: border-box;padding-right:3px}" & vbCrlf
	tmpHtml = tmpHtml & "		.revoke-tips {text-align:center;color:#fff;background:#f30;font-size:1.4rem;position:fixed;bottom:0;width:100%;line-height:50px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-gap-10 {height:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-editbtn em.pass {flex-grow:4;width:auto;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-swap-box"">" & vbCrlf
	Response.Write "	<div class=""hr-swap-items"">" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>申请人：</dt><dd>" & tApplyer & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>申请时间：</dt><dd>" & tApplyTime & "</dd></dl>" & vbCrlf
	If Action = "AlternDetails" Then
		Response.Write "		<dl class=""hr-rows""><dt>科室：</dt><dd>" & tApplyerKS & "</dd></dl>" & vbCrlf
		Response.Write "		<dl class=""hr-rows""><dt>职务：</dt><dd>" & tApplyerZW & "</dd></dl>" & vbCrlf
		Response.Write "		<dl class=""hr-rows""><dt>职称：</dt><dd>" & tApplyerZC & "</dd></dl>" & vbCrlf
	End If
	Response.Write "		<div class=""hr-gap-20 hr-gap-10 hr-fix""></div>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>申请课程：</dt><dd>" & tCourseDate & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>　</dt><dd>" & tPeriod & " " & tPeriodTime & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>课程名称：</dt><dd>" & tCourse & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>授课对象：</dt><dd>" & tStuClass & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>授课地点：</dt><dd>" & tPlace & "</dd></dl>" & vbCrlf


	Response.Write "		<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>换课教师：</dt><dd>" & tReplacer & " " & PassProcess(1, tReplacePass) & "</dd></dl>" & vbCrlf
	If HR_IsNull(tReplacerTime) = False Then Response.Write "		<dl class=""hr-rows""><dt>确认时间：</dt><dd>" & tReplacerTime & "</dd></dl>" & vbCrlf
	If Action = "Details" Then
		Response.Write "		<dl class=""hr-rows""><dt>科室：</dt><dd>" & tReplacerKS & "</dd></dl>" & vbCrlf
		Response.Write "		<dl class=""hr-rows""><dt>职务：</dt><dd>" & tReplacerZW & "</dd></dl>" & vbCrlf
		Response.Write "		<dl class=""hr-rows""><dt>职称：</dt><dd>" & tReplacerZC & "</dd></dl>" & vbCrlf
	End if
	Response.Write "		<div class=""hr-gap-20 hr-gap-10 hr-fix""></div>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>授课时间：</dt><dd>" & tCourseDate1 & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>　</dt><dd>" & tPeriod & " " & tPeriodTime1 & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>课程名称：</dt><dd>" & tCourse1 & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>授课对象：</dt><dd>" & tStuClass1 & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>授课地点：</dt><dd>" & tPlace1 & "</dd></dl>" & vbCrlf
	Response.Write "		<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>原　因：</dt><dd>" & Replace(tReason, chr(10), "<br>") & "</dd></dl>" & vbCrlf
	Response.Write "		<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>教研室主任：</dt><dd>" & tPasser & " " & PassProcess(2, tPasserPass) & "</dd></dl>" & vbCrlf
	If HR_IsNull(tPassTime) = False Then Response.Write "		<dl class=""hr-rows""><dt>审核时间：</dt><dd>" & tPassTime & "</dd></dl>" & vbCrlf

	Response.Write "		<dl class=""hr-rows""><dt>教学处：</dt><dd>" & PassProcess(3, tPasser1Pass) & "</dd></dl>" & vbCrlf
	If HR_IsNull(tPassTime1) = False Then Response.Write "		<dl class=""hr-rows""><dt>审核时间：</dt><dd>" & tPassTime1 & "</dd></dl>" & vbCrlf

	Response.Write "		<dl class=""hr-rows""><dt>教辅：</dt><dd>" & PassProcess(4, tPasser2Pass) & "</dd></dl>" & vbCrlf
	If HR_IsNull(tPassTime2) = False Then Response.Write "		<dl class=""hr-rows""><dt>审核时间：</dt><dd>" & tPassTime2 & "</dd></dl>" & vbCrlf
	Response.Write "		<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	
	Response.Write "	<div class=""hr-shrink-x20""></div>" & vbCrlf
	If tProcess = 5 Then
		Response.Write "	<div class=""revoke-tips""><em>已撤销</em></div>" & vbCrlf
	ElseIf tProcess < 4 And UserYGDM = tApplyID Then
		Response.Write "	<div class=""hr-rows hr-editbtn"">" & vbCrlf
		Response.Write "		<em><i class=""hr-icon"">&#xea3f;</i></em>" & vbCrlf
		If tProcess = 0 Then
			Response.Write "		<em><button type=""button"" name=""edit"" class=""edit"" id=""EditForm"" data-id=""" & tmpID & """>修改</button></em>" & vbCrlf
			Response.Write "		<em><button type=""button"" name=""delete"" class=""delete"" id=""Delete"" data-id=""" & tmpID & """>删除</button></em>" & vbCrlf
		End If
		Response.Write "		<em><button type=""button"" name=""revoke"" class=""revoke"" id=""Revoke"" data-id=""" & tmpID & """>撤销</button></em>" & vbCrlf
		Response.Write "	</div>" & vbCrlf
	ElseIf tProcess=0 And UserYGDM = tReplacerID Then		'换课老师
		Response.Write "	<div class=""hr-rows hr-editbtn"">" & vbCrlf
		Response.Write "		<em><i class=""hr-icon"">&#xea3f;</i></em>" & vbCrlf
		Response.Write "		<em><button type=""button"" name=""agree"" class=""agree"" data-id=""" & tmpID & """ data-pass=""1"">确认</button></em>" & vbCrlf
		Response.Write "		<em><button type=""button"" name=""agree"" class=""agree"" data-id=""" & tmpID & """ data-pass=""2"">拒绝</button></em>" & vbCrlf
		Response.Write "	</div>" & vbCrlf
	ElseIf UserYGDM = tPasserID Or SwapPass>0 Then		'教研主任审核或教学处及教辅
		Response.Write "	<div class=""hr-rows hr-editbtn"">" & vbCrlf
		Response.Write "		<em><i class=""hr-icon"">&#xea3f;</i></em>" & vbCrlf
		Response.Write "		<em class=""pass""><a href=""" & ParmPath & "SwapPass/EditPass.html?ID=" & tmpID & """ name=""edit"" class=""edit"" data-id=""" & tmpID & """>进入审核</a></em>" & vbCrlf
		Response.Write "	</div>" & vbCrlf
	End If
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").html(""<i class='hr-icon'>&#xf0ca;</i>"");" & vbCrlf		
	tmpHtml = tmpHtml & "	$("".navBack em"").html(""<i class='hr-icon'>&#xf320;</i>"");" & vbCrlf	

	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "SwapCourse/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	var laynav=""<li><a href=\""" & ParmPath & "SwapCourse/Index.html\""><i class=\""hr-icon hr-icon-top\"">&#xe853;</i>我申请的换课</a></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "	laynav+=""<li><a href=\""" & ParmPath & "SwapCourse/Alternate.html\""><i class=\""hr-icon hr-icon-top\"">&#xf2dd;</i>我的换课记录</a></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "	$("".nctouch-nav-menu ul"").html(laynav);" & vbCrlf				'更新右上角导航

	tmpHtml = tmpHtml & "	$(""#EditForm"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var swapid=$(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		location.href=""" & ParmPath & "SwapCourse/Edit.html?Modify=True&ID=""+ swapid;" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#Delete"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var swapid=$(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		$.confirm(""您确定要删除本次申请？"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "SwapCourse/Delete.html"",{ID:swapid},function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "				$.toast(reData.reMessge, function(){location.href=""" & ParmPath & "SwapCourse/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$("".agree"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var swapid=$(this).data(""id""), swappass=$(this).data(""pass"");" & vbCrlf
	tmpHtml = tmpHtml & "		var passtxt=""您确定同意与" & tApplyer & "老师对换课程吗？"";" & vbCrlf
	tmpHtml = tmpHtml & "		if(!swappass){ passtxt=""您确定拒绝与" & tApplyer & "老师换课吗？""; };" & vbCrlf
	tmpHtml = tmpHtml & "		$.confirm(passtxt,function(){" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "SwapCourse/Agree.html"",{ID:swapid, Passed:swappass},function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "				$.toast(reData.errmsg, function(){" & vbCrlf
	'tmpHtml = tmpHtml & "					if(!reData.err){ location.href=""" & ParmPath & "SwapCourse/ListAgree.html""; }" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub Edit()
	SiteTitle = "对换课申请"
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim IsModify : IsModify = False

	Dim tTableName, tTemplate, tItemName, tCourse, tCourseID, tReplacer, tReplacerID, tReason
	Dim tPasser, tPasserID, tPasser1, tPasserID1, tPasser2, tPasserID2
	Dim tStudent, tContents, tCourseDate
	Dim tVA3, tVA5, tVA6, tVA7, tVA8, tVA11, tVA12
	Dim tSwapItem, tSwapItemID, tSwapCourse, tSwapCourseID

	If tmpID > 0 Then
		sqlTmp = "Select a.* From HR_SwapCourse a Where a.ID=" & tmpID
		Set rsTmp = Conn.Execute(sqlTmp)
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				IsModify = True
				SiteTitle = "修改代课申请"
				tItemID = rsTmp("ItemID")
				tItemName = GetTypeName("HR_Class", "ClassName", "ClassID", tItemID)
				tTemplate = GetTypeName("HR_Class", "Template", "ClassID", tItemID)
				tCourseID = rsTmp("CourseID")
				tTableName = "HR_Sheet_" & tItemID
				tCourse = FormatDate(Trim(rsTmp("newVA4")), 2)
				oldCourseDate = FormatDate(rsTmp("VA4"), 2)
				tCourseDate = FormatDate(rsTmp("newVA4"), 2)
				tCourse = tCourse & " " & rsTmp("newVA8") & "_第" & rsTmp("newVA7") & "节"
				oldContents = rsTmp("VA9")
				tContents = rsTmp("newVA9")
				oldStudent = rsTmp("VA10")
				tStudent = rsTmp("newVA10")
				oldVA3 = rsTmp("VA3")
				tVA3 = rsTmp("newVA3")
				tVA5 = rsTmp("newVA5")
				tVA6 = rsTmp("newVA6")
				oldVA7 = rsTmp("VA7")
				tVA7 = rsTmp("newVA7")
				tVA8 = rsTmp("newVA8")
				oldVA11 = rsTmp("VA11")
				tVA11 = rsTmp("newVA11")
				tVA12 = rsTmp("newVA12")
				tReplacerID = HR_Clng(rsTmp("Replacer"))
				tReplacer = rsTmp("ReplacerName")
				tPasserID = HR_Clng(rsTmp("Passer"))
				tPasser = Trim(rsTmp("PasserName"))
				tReason = Trim(rsTmp("Reason"))
			End If
		Set rsTmp = Nothing
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dl {align-items:stretch;border-bottom:1px solid #eee;padding:5px 0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cells_form .weui-cell__ft {font-size:1.1rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cells_form .popWin i {font-size:1.2rem;position:relative;top:2px;color:#4ce}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-toast {margin-left: auto;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cell__hd .weui-label {color:#999;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cells__title {border-top:10px solid #eee;font-size:1.1rem;padding-top:8px}" & vbCrlf

	tmpHtml = tmpHtml & "		.old-box h3 {padding:10px; border-bottom:1px solid #4fb74e;}" & vbCrlf
	tmpHtml = tmpHtml & "		.old-box ul {padding:10px; display:flex; flex-direction:column;}" & vbCrlf
	tmpHtml = tmpHtml & "		.old-box li {padding:10px; border-bottom:1px solid #ddd; display:flex;}" & vbCrlf
	tmpHtml = tmpHtml & "		.old-box li tt {width:5.2rem;color:#999;flex-shrink:0;font-size:0.9rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.old-box li em {font-size:1rem;}" & vbCrlf

	tmpHtml = tmpHtml & "		.weui-cell_switch {border-bottom:1px solid #f17be2;color:#b563ab;}" & vbCrlf
	tmpHtml = tmpHtml & "		.modi-body {display:none;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	
	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "<div class=""weui-cells__title"">第一步：选择课程</div>" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">申请人：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Applyer"" class=""weui-input"" id=""Applyer"" type=""text"" value=""" & UserYGXM & """ readonly>" & vbCrlf
	Response.Write "			<input name=""ApplyID"" class=""weui-input"" id=""ApplyID"" type=""hidden"" value=""" & UserYGDM & """ data-values=""" & UserYGDM & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择项目：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""CourseItem"" class=""weui-input opt1"" id=""CourseItem"" type=""text"" value=""" & tItemName & """ data-values=""" & tItemID & """ readonly>" & vbCrlf
	Response.Write "			<input name=""CourseItemID"" id=""CourseItemID"" type=""hidden"" value=""" & tItemID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><i class=""hr-icon"">&#xf321;</i></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择课程：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Course"" class=""weui-input opt2"" id=""Course"" type=""text"" value=""" & tCourse & """ data-values=""" & tCourseID & """>" & vbCrlf
	Response.Write "			<input name=""CourseID"" id=""CourseID"" type=""hidden"" value=""" & tCourseID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><i class=""hr-icon"">&#xf321;</i></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20 hr-gapbg""></div>" & vbCrlf
	Response.Write "	<div class=""old-box myCourse""><h3>我的课程信息：</h3><ul>" & vbCrlf
	If IsModify Then
		Response.Write "	<li><tt>授课日期：</tt><em>" & oldCourseDate & "</em></li>" & vbCrlf
		Response.Write "	<li><tt>节次：</tt><em>" & oldVA7 & "</em></li>" & vbCrlf
		Response.Write "	<li><tt>学时：</tt><em>" & oldVA3 & "</em></li>" & vbCrlf
		Response.Write "	<li><tt>授课对象：</tt><em>" & oldStudent & "</em></li>" & vbCrlf
		Response.Write "	<li><tt>授课内容：</tt><em>" & oldContents & "</em></li>" & vbCrlf
		Response.Write "	<li><tt>校(院)区：</tt><em>" & oldVA11 & "</em></li>" & vbCrlf
		Response.Write "	" & vbCrlf
	End If
	Response.Write "	</ul></div>" & vbCrlf
	Response.Write "	<div class=""weui-cells__title"">第二步：选择换课教师</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">换课教师：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Swapper"" class=""weui-input"" id=""ygxm"" type=""text"" value=""" & tReplacer & """ data-key=""ygxm"" data-value=""ygdm"" placeholder="""""
	Response.Write ">" & vbCrlf
	Response.Write "			<input name=""SwapperID"" class=""weui-input"" id=""ygdm"" type=""hidden"" value=""" & tReplacerID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft popWin"" data-id=""ygxm""><i class=""hr-icon"">&#xeeed;</i>选择</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择项目：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""SwapItem"" class=""weui-input opt1"" id=""SwapItem"" type=""text"" value=""" & tSwapItem & """ data-values=""" & tSwapItemID & """ readonly>" & vbCrlf
	Response.Write "			<input name=""SwapItemID"" id=""SwapItemID"" type=""hidden"" value=""" & tSwapItemID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><i class=""hr-icon"">&#xf321;</i></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择课程：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""SwapCourse"" class=""weui-input opt2"" id=""SwapCourse"" type=""text"" value=""" & tSwapCourse & """ data-values=""" & tSwapCourseID & """>" & vbCrlf
	Response.Write "			<input name=""SwapCourseID"" id=""SwapCourseID"" type=""hidden"" value=""" & tSwapCourseID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><i class=""hr-icon"">&#xf321;</i></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20 hr-gapbg""></div>" & vbCrlf
	Response.Write "	<div class=""old-box swapCourse""><h3>对换课程信息：</h3><ul>" & vbCrlf
	If IsModify Then
		Response.Write "	<li><tt>授课日期：</tt><em>" & oldCourseDate & "</em></li>" & vbCrlf
		Response.Write "	<li><tt>节次：</tt><em>" & oldVA7 & "</em></li>" & vbCrlf
		Response.Write "	<li><tt>学时：</tt><em>" & oldVA3 & "</em></li>" & vbCrlf
		Response.Write "	<li><tt>授课对象：</tt><em>" & oldStudent & "</em></li>" & vbCrlf
		Response.Write "	<li><tt>授课内容：</tt><em>" & oldContents & "</em></li>" & vbCrlf
		Response.Write "	<li><tt>校(院)区：</tt><em>" & oldVA11 & "</em></li>" & vbCrlf
		Response.Write "	" & vbCrlf
	End If
	Response.Write "	</ul></div>" & vbCrlf

	Response.Write "	<div class=""weui-cells__title"">第三步：选择教研主任</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">教研主任：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Passer"" class=""weui-input"" id=""Passer"" type=""text"" value=""" & tPasser & """ data-key=""Passer"" data-value=""PasserID"" placeholder="""">" & vbCrlf
	Response.Write "			<input name=""PasserID"" class=""weui-input"" id=""PasserID"" type=""hidden"" value=""" & tPasserID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft popWin"" data-id=""Passer""><i class=""hr-icon"">&#xeeed;</i>选择审核人</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""weui-cells__title"">换课理由：</div>" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Reason"" id=""Reason"" placeholder=""请输入换课理由"" rows=""5"">" & tReason & "</textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	If IsModify Then Response.Write "<input name=""SwapID"" type=""hidden"" value=""" & tmpID & """>" & vbCrlf
	Response.Write "<div class=""weui-btn-area""><em class=""weui-btn weui-btn_primary"" id=""subPost"">提交申请</em></div>" & vbCrlf
	Response.Write "</form>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div id=""full"" class=""hr-popup"">" & vbCrlf
	Response.Write "	<iframe src=""about:bank"" name=""listFrame"" id=""listFrame"" title=""ListFrame"" width=""100%"" height=""100%"" frameborder=""0""></iframe>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").html(""<i class='hr-icon'>&#xf0ca;</i>"");" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack em"").html(""<i class='hr-icon'>&#xf320;</i>"");" & vbCrlf	
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf

	tmpHtml = tmpHtml & "	var laynav=""<li><a href=\""" & ParmPath & "SwapCourse/Index.html\""><i class=\""hr-icon hr-icon-top\"">&#xe853;</i>我申请的换课</a></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "	laynav+=""<li><a href=\""" & ParmPath & "SwapCourse/Alternate.html\""><i class=\""hr-icon hr-icon-top\"">&#xf2dd;</i>我的换课记录</a></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "	$("".nctouch-nav-menu ul"").html(laynav);" & vbCrlf				'更新右上角导航

	tmpHtml = tmpHtml & "	$("".popWin"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#full"").show();var obj=$(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#listFrame"").attr(""src"",""" & ParmPath & "Directories/SelectTeacher.html?Type=3&reObjTxt="" + $(""#""+obj).data(""key"") + ""&reObjValue="" +  $(""#""+obj).data(""value""));" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	$(""#CourseDate"").on(""change"",function(){" & vbCrlf		'自动计算星期
	tmpHtml = tmpHtml & "		var today = new Array('日','一','二','三','四','五','六'), day = new Date($(this).val());" & vbCrlf
	tmpHtml = tmpHtml & "		var week = today[day.getDay()];$(""#VA6"").val(week);" & vbCrlf
	tmpHtml = tmpHtml & "		console.log($(this).val());" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	var sPeriod=[], ePeriod=[];" & vbCrlf
	tmpHtml = tmpHtml & "	for(var k=1; k<20; k++){ sPeriod.push(k); ePeriod.push(k+1); }" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#VA7"").picker({" & vbCrlf
	tmpHtml = tmpHtml & "		title: ""请选择节次"",cols:[" & vbCrlf
	tmpHtml = tmpHtml & "			{textAlign:'center',values:sPeriod}," & vbCrlf
	tmpHtml = tmpHtml & "			{textAlign:'left',values:ePeriod}," & vbCrlf
	tmpHtml = tmpHtml & "		]," & vbCrlf
	tmpHtml = tmpHtml & "		onClose:function(e){" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#VA7"").val(e.value[0] +""-""+ e.value[1]);" & vbCrlf
	tmpHtml = tmpHtml & "		}" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#switch"").on(""change"",function(e){" & vbCrlf
	tmpHtml = tmpHtml & "		var chked = $(this).is("":checked"");" & vbCrlf
	tmpHtml = tmpHtml & "		if(chked){ $("".modi-body"").slideToggle(); }" & vbCrlf
	tmpHtml = tmpHtml & "		else{ $("".modi-body"").slideToggle(); }" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf		'提交保存
	tmpHtml = tmpHtml & "		var reason = $(""#Reason"").val(), swapperid=$(""#ygdm"").val(), itemid=$(""#CourseItem"").data(""values""), courseid=$(""#Course"").data(""values""), passer=$(""#PasserID"").val();" & vbCrlf
	tmpHtml = tmpHtml & "		var newItemID = $(""#SwapItem"").data(""values""), newCourseID= $(""#SwapCourse"").data(""values"");" & vbCrlf

	'tmpHtml = tmpHtml & "		$(""#ItemID"").val(parseInt(itemid)); $(""#CourseID"").val(parseInt(courseid));" & vbCrlf
	'tmpHtml = tmpHtml & "		$(""#ItemID"").val(parseInt(itemid)); $(""#CourseID"").val(parseInt(courseid));" & vbCrlf

	tmpHtml = tmpHtml & "		if(!reason){$.toast(""请输入换课理由！"", ""text"");return false;};" & vbCrlf
	tmpHtml = tmpHtml & "		if(!courseid){ $.toast(""请选择课程！"",""text""); return false;}" & vbCrlf
	tmpHtml = tmpHtml & "		if(!newCourseID){ $.toast(""请选择换课老师的课程！"",""text""); return false;}" & vbCrlf
	tmpHtml = tmpHtml & "		if(parseInt(swapperid)==0){$.toast(""请选择换课老师！"", ""text""); return false;};" & vbCrlf
	tmpHtml = tmpHtml & "		if(passer==""""){$.toast(""请选择教研室主任！"", ""text""); return false;};" & vbCrlf
	tmpHtml = tmpHtml & "		$.showLoading();" & vbCrlf	
	tmpHtml = tmpHtml & "		$.post(""" & ParmPath & "SwapCourse/SavePost.html"",$(""#EditForm"").serialize(), function(res){" & vbCrlf
	tmpHtml = tmpHtml & "			console.log(res);$.hideLoading();" & vbCrlf
	tmpHtml = tmpHtml & "			if(res.err){ $.toptip(res.errmsg, 'error'); }" & vbCrlf
	tmpHtml = tmpHtml & "			else{ $.toast(res.errmsg, function(){ location.href=""" & ParmPath & "SwapCourse/Index.html""; }); }" & vbCrlf
	tmpHtml = tmpHtml & "		},""json"");" & vbCrlf
	tmpHtml = tmpHtml & "		return false;" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	var arrItem =[" & GetSelectOptionItem() & "], chkItem = false;" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#CourseItem"").select({" & vbCrlf			'申请人选择项目
	tmpHtml = tmpHtml & "		title: ""请选择项目"", items:arrItem," & vbCrlf
	tmpHtml = tmpHtml & "		beforeClose:function(data1){ $.showLoading(); }," & vbCrlf			'选择时
	tmpHtml = tmpHtml & "		onClose:function(elBack){" & vbCrlf
	tmpHtml = tmpHtml & "			var tid = elBack.data.values; $(""#CourseItemID"").val(tid);" & vbCrlf
	tmpHtml = tmpHtml & "			$.get(""" & ParmPath & "SwapCourse/getItemCourse.html"",{Item:tid, YGDM:" & UserYGDM & "}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "				var reData = eval(""("" + strForm + "")"");" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#Course"").select(""update"", reData); $.hideLoading();" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#Course"").val(""""); $(""#Course"").data(""values"",""0"");" & vbCrlf
	tmpHtml = tmpHtml & "				$("".myCourse ul"").html("""");" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		}" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	$(""#Course"").select({" & vbCrlf			'申请人选择项目
	tmpHtml = tmpHtml & "		title:""选择课程"",items:[{title:""暂无课程"",value:""""}]," & vbCrlf
	tmpHtml = tmpHtml & "		onOpen:function(){" & vbCrlf			'打开时回调
	tmpHtml = tmpHtml & "			var itemid1 = parseInt($(""#CourseItem"").data(""values""));" & vbCrlf
	tmpHtml = tmpHtml & "			if(itemid1==0){chkItem=true;}" & vbCrlf
	tmpHtml = tmpHtml & "			console.log(itemid1);" & vbCrlf
	tmpHtml = tmpHtml & "		}," & vbCrlf
	tmpHtml = tmpHtml & "		onClose:function(elBack){" & vbCrlf
	tmpHtml = tmpHtml & "			var str1="""",cid = elBack.data.values, itemid = $(""#CourseItem"").data(""values"");" & vbCrlf
	tmpHtml = tmpHtml & "			if(itemid==0){$.toptip(""请选择项目"", ""error""); return false;};" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#CourseID"").val(cid); console.log(cid);" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "SwapCourse/getCourse.html"",{ID:cid, ItemID:itemid}, function(redata){" & vbCrlf
	tmpHtml = tmpHtml & "				str1 = ""<li><tt>授课日期：</tt><em>""+ redata.CourseDate +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>节次：</tt><em>""+ redata.Period +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>学时：</tt><em>""+ redata.VA3 +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>授课对象：</tt><em>""+ redata.Student +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>授课内容：</tt><em>""+ redata.Contents +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>校(院)区：</tt><em>""+ redata.VA11 +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				$("".myCourse ul"").html(str1);" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		}" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	$(""#ygdm"").on(""change"",function(){" & vbCrlf			'选择被换课人后更新课程信息
	tmpHtml = tmpHtml & "		console.log(""asdfasdf"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	$(""#SwapItem"").select({" & vbCrlf			'被换课人项目
	tmpHtml = tmpHtml & "		title: ""请选择换课课程项目"",items:arrItem," & vbCrlf
	tmpHtml = tmpHtml & "		beforeClose:function(data1){ $.showLoading(); }," & vbCrlf			'选择时
	tmpHtml = tmpHtml & "		onClose:function(elBack){" & vbCrlf
	tmpHtml = tmpHtml & "			var swapperCode = $(""#ygdm"").val(); $(""#SwapItemID"").val(elBack.data.values);" & vbCrlf
	tmpHtml = tmpHtml & "			$.get(""" & ParmPath & "SwapCourse/getItemCourse.html"",{Item:elBack.data.values, YGDM:swapperCode}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "				var reData = eval(""("" + strForm + "")"");" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#SwapCourse"").select(""update"", reData); $.hideLoading();" & vbCrlf	'数据读取完后关闭提示
	tmpHtml = tmpHtml & "				$(""#SwapCourse"").val("""");" & vbCrlf
	tmpHtml = tmpHtml & "				$("".swapCourse ul"").html("""");" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		}" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#SwapCourse"").select({" & vbCrlf		'换课课程
	tmpHtml = tmpHtml & "		title:""选择换课课程"",items:[{title:""暂无课程"",value:""""}]," & vbCrlf
	tmpHtml = tmpHtml & "		onOpen:function(data1){" & vbCrlf			'打开时回调
	'tmpHtml = tmpHtml & "			console.log(e.config.items); return false;" & vbCrlf
	tmpHtml = tmpHtml & "		}," & vbCrlf
	tmpHtml = tmpHtml & "		onClose:function(elBack){" & vbCrlf
	tmpHtml = tmpHtml & "			var str1="""",cid = elBack.data.values, itemid = $(""#SwapItem"").data(""values"");" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#SwapCourseID"").val(cid);" & vbCrlf
	tmpHtml = tmpHtml & "			if(cid==0){return false;};" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#SwapCourseID"").val(cid);" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "SwapCourse/getCourse.html"",{ID:cid, ItemID:itemid}, function(redata){" & vbCrlf
	tmpHtml = tmpHtml & "				str1 = ""<li><tt>授课日期：</tt><em>""+ redata.CourseDate +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>节次：</tt><em>""+ redata.Period +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>学时：</tt><em>""+ redata.VA3 +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>授课对象：</tt><em>""+ redata.Student +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>授课内容：</tt><em>""+ redata.Contents +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<li><tt>校(院)区：</tt><em>""+ redata.VA11 +""</em></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "				$("".swapCourse ul"").html(str1);" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		}" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub getItemCourse()
	Dim tItemID : tItemID = HR_Clng(Request("Item"))
	Dim tYGDM : tYGDM = HR_Clng(Request("YGDM"))
	Dim tTableName : tTableName = "HR_Sheet_" & tItemID
	Dim strTmp
	strTmp = strTmp & "{""items"":["
	If ChkTable(tTableName) Then
		sql = "Select a.*,b.ClassName,b.Template From " & tTableName & " a Inner Join HR_Class b on a.ItemID=b.ClassID Where VA1=" & HR_Clng(tYGDM)
		sql = sql & " And scYear=" & DefYear & " Order By VA4 DESC"
		Set rs = Conn.Execute(sql)
			If Not(rs.BOF And rs.EOF) Then
				i = 0
				Do While Not rs.EOF
					If i>0 Then strTmp = strTmp & ","
					strTmp = strTmp & "{title:""" & FormatDate(ConvertNumDate(rs("VA4")), 2)
					If rs("Template") = "TempTableA" Then
						strTmp = strTmp & " " & rs("VA8") & "_" & rs("VA7") & "节"
					Else
						strTmp = strTmp & " " & rs("VA6")
					End If
					strTmp = strTmp & """,value:""" & rs("ID") & """}"
					rs.MoveNext
					i = i + 1
				Loop
			Else
				strTmp = strTmp & "{title:""暂无课程"",value:""0""}"
			End If
		Set rs = Nothing
	Else
		strTmp = strTmp & "{title:""数据表不存在"",value:""0""}"
	End If
	strTmp = strTmp & "]}"
	Response.Write strTmp
End Sub

Sub SavePost()
	Dim tmpID : tmpID = HR_Clng(Request("SwapID"))
	Dim tApplyID : tApplyID = HR_Clng(Request("ApplyID"))
	Dim tCourseItemID : tCourseItemID = HR_Clng(Request("CourseItemID"))
	Dim tCourseID : tCourseID = HR_Clng(Request("CourseID"))
	Dim tSwapperID : tSwapperID = HR_Clng(Request("SwapperID"))
	Dim tSwapItemID : tSwapItemID = HR_Clng(Request("SwapItemID"))
	Dim tSwapCourseID : tSwapCourseID = HR_Clng(Request("SwapCourseID"))
	Dim tPasserID : tPasserID = HR_Clng(Request("PasserID"))
	Dim isModify : isModify = HR_CBool(Request("Modify"))
	Dim tTableName, tReason : ErrMsg = ""

	If tApplyID = 0 Then ErrMsg = "您未登陆"
	If tCourseItemID = 0 Then ErrMsg = "请选择项目"
	If tCourseID = 0 Then ErrMsg = "您还没有选择课程"
	If tSwapperID = 0 Then ErrMsg = "请选择换课老师"
	If tSwapItemID = 0 Then ErrMsg = "换课项目没有选择"
	If tSwapCourseID = 0 Then ErrMsg = "请选择换课课程"
	If tPasserID = 0 Then ErrMsg = "请选择教研主任"
	tReason = Trim(Request("Reason"))
	If HR_IsNull(tReason) Then ErrMsg = "换课理由没有输入"

	If HR_IsNull(ErrMsg) = False Then Response.Write "{""err"":true,""errcodde"":500,""errmsg"":""" & ErrMsg & """,""icon"":2}" : Exit Sub
	'Response.Write "{""err"":false,""errcodde"":0,""errmsg"":""暂未保存"",""icon"":1}"
	'Exit Sub
	Dim rsSave, sqlSave
	sqlSave = "Select * From HR_Swap Where ID=" & tmpID
	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open sqlSave, Conn, 1, 3
		If rsSave.BOF And rsSave.EOF Then
			rsSave.AddNew
			tmpID = GetNewID("HR_Swap", "ID")
			rsSave("ID") = tmpID
			rsSave("YGDM") = UserYGDM
		Else
			isModify = True
		End If
		rsSave("ItemID") = tCourseItemID
		rsSave("CourseID") = tCourseID
		rsSave("newItemID") = tSwapItemID
		rsSave("newCourseID") = tSwapCourseID
		'保存原数据
		tTableName = "HR_Sheet_" & tCourseItemID
		If ChkTable(tTableName) Then
			Set rsTmp = Conn.Execute("Select * From " & tTableName & " Where ID=" & tCourseID)
				If Not(rsTmp.BOF And rsTmp.EOF) Then
					rsSave("VA3") = HR_CDbl(rsTmp("VA3"))
					rsSave("VA4") = FormatDate(ConvertNumDate(rsTmp("VA4")), 4)
					rsSave("VA5") = Trim(rsTmp("VA5"))
					rsSave("VA6") = Trim(rsTmp("VA6"))
					rsSave("VA7") = Trim(rsTmp("VA7"))
					rsSave("VA8") = Trim(rsTmp("VA8"))
					rsSave("VA9") = Trim(rsTmp("VA9"))
					rsSave("VA10") = Trim(rsTmp("VA10"))
					rsSave("VA11") = Trim(rsTmp("VA11"))
					rsSave("VA12") = Trim(rsTmp("VA12"))
				End If
			Set rsTmp = Nothing
		End If
		'保存对换课数据
		tTableName = "HR_Sheet_" & tSwapItemID
		If ChkTable(tTableName) Then
			Set rsTmp = Conn.Execute("Select * From " & tTableName & " Where ID=" & tSwapCourseID)
				If Not(rsTmp.BOF And rsTmp.EOF) Then
					rsSave("newVA3") = HR_CDbl(rsTmp("VA3"))
					rsSave("newVA4") = FormatDate(ConvertNumDate(rsTmp("VA4")), 4)
					rsSave("newVA5") = Trim(rsTmp("VA5"))
					rsSave("newVA6") = Trim(rsTmp("VA6"))
					rsSave("newVA7") = Trim(rsTmp("VA7"))
					rsSave("newVA8") = Trim(rsTmp("VA8"))
					rsSave("newVA9") = Trim(rsTmp("VA9"))
					rsSave("newVA10") = Trim(rsTmp("VA10"))
					rsSave("newVA11") = Trim(rsTmp("VA11"))
					rsSave("newVA12") = Trim(rsTmp("VA12"))
				End If
			Set rsTmp = Nothing
		End If
		rsSave("Reason") = Trim(Request("Reason"))
		rsSave("Replacer") = tSwapperID
		rsSave("Passer") = tPasserID
		rsSave("ApplyTime") = Now()
		rsSave("Process") = 0		'审核步骤
		rsSave.Update
	Set rsSave = Nothing
	ErrMsg = "申请提交成功！"
	'发送提醒
	Dim url1 : url1 = SiteUrl & "/Touch/SwapCourse/Details.html?ItemID=" & tCourseItemID & "&ID=" & tmpID

	ErrMsg = UserYGXM & "老师：您申请与" & Trim(Request("Swapper")) & "老师对换课已提交成功，等待" & Trim(Request("Swapper")) & "老师确认后由教研主任" & Trim(Request("Passer")) & "审核。"
	ErrMsg = ErrMsg & "<br>我的课程：" & Trim(Request("Course")) & "<br>对换课程：" & Trim(Request("SwapCourse")) & "<br>理由：" & tReason & "<br>申请时间：" & FormatDate(Now(), 10)
	Call SentWechatMSG_QYCard(UserYGDM, UserYGXM & "：您已申请与" & Trim(Request("Swapper")) & "老师对换课程", url1, ErrMsg)		'发送给自己

	ErrMsg = UserYGXM & "申请与您对换课程，等待您确认后由教研主任" & Trim(Request("Passer")) & "审核。"
	ErrMsg = ErrMsg & "<br>他的课程：" & Trim(Request("Course")) & "<br>我的课程：" & Trim(Request("SwapCourse")) & "<br>理由：" & tReason & "<br>申请时间：" & FormatDate(Now(), 10)
	Call SentWechatMSG_QYCard(tSwapperID, UserYGXM & " 申请与您对换课程，需要您确认", url1, ErrMsg)	'发送给换课老师
	'Call SentWechatMSG_QYCard(tPasserID, UserYGXM & " 申请调换课程，需要您审核", url1, UserYGXM & "老师申请调换课程。<br>理由：" & tReason & "<br>申请时间：" & FormatDate(Now(), 1))		'发送给教研主任
	Response.Write "{""err"":false,""errcode"":0,""icon"":1,""errmsg"":""对换课申请提交成功！""}"
End Sub

Sub Delete()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	tmpID = FilterArrNull(tmpID, ",")

	Conn.Execute("Delete From HR_Swap Where YGDM=" & UserYGDM & " And ID in(" & tmpID & ")")
	ErrMsg = "删除成功！"
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub

Function GetSelectOptionItem()				'取考核项目下拉
	Dim iFun, funItem, rsFun, sqlFun
	sqlFun = "Select ClassID, ClassName From HR_Class Where ModuleID=1001 And Child=0 And Template='TempTableA'"
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
					strTmp = strTmp & ",""Course"":""" & Trim(rs("VA8")) & """,""Student"":""" & FilterHtmlToText(rs("VA10")) & """,""Period"":""" & rs("VA7") & """,""VA3"":" & HR_CDbl(rs("VA3")) & ",""VA5"":""" & Trim(rs("VA5")) & """"
					strTmp = strTmp & ",""VA6"":""" & Trim(rs("VA6")) & """,""Contents"":""" & Trim(rs("VA9")) & """,""VA11"":""" & FilterHtmlToText(rs("VA11")) & """,""VA12"":""" & Trim(rs("VA12")) & """"
				Else
					strTmp = strTmp & ",""Course"":""" & rs("VA6")& """"
				End If
				strTmp = strTmp & ",""CourseID"":" & HR_CLng(rs("ID")) & ",""ItemID"":" & HR_CLng(rs("ItemID")) & ""
			Else
				strTmp = """err"":true,""errcode"":500,""errmsg"":""该教师在本项目中没有课程"""
			End If
		Set rs = Nothing
	Else
		strTmp = """err"":true,""errcode"":500,""errmsg"":""数据表不存在"""
	End If
	Response.Write "{" & strTmp & "}"
End Sub

Sub Alternate()			'换课记录
	SiteTitle = "我的换课记录"
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cell {padding:5px 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-list-title {padding:0 10px; box-sizing:border-box; line-height:35px;background:#fffbf1;color:#f60;border:1px solid f5bdaf;border-left:0;border-right:0}" & vbCrlf
	'tmpHtml = tmpHtml & "		.hr-list-cell {padding:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-list-cell .hr-rows {border-bottom:5px solid #eee;padding:8px 3px}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-list-cell .hr-item {padding:0 2px;line-height:1.5}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-list-cell .iconTit {color:#2196f3;font-size:1.3rem;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim tCourseList, tProcess, tPass, tPeriodTime, tApplyer
	sqlTmp = "Select a.*,(Select ClassName From HR_CLass Where ClassID=a.ItemID) As ItemName,(Select YGXM From HR_Teacher Where YGDM=a.YGDM) As Applyer From HR_Swap a Where a.newItemID>0 And a.newCourseID>0 And a.Replacer=" & UserYGDM		'取课程SQL
	sqlTmp = sqlTmp & " Order By a.ApplyTime DESC"
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			tCourseList = tCourseList & "<div class=""hr-list-cell"">"
			Do While Not rsTmp.EOF
				tProcess = HR_CLng(rsTmp("Process"))
				If tProcess = 1 Then
					tPass = HR_CLng(rsTmp("ReplacePass"))
				ElseIf tProcess = 2 Then
					tPass = HR_CLng(rsTmp("PasserPass"))		'教研主任
				ElseIf tProcess = 3 Then
					tPass = HR_CLng(rsTmp("Passer1Pass"))		'教学处审核
				ElseIf tProcess = 4 Then
					tPass = HR_CLng(rsTmp("Passer2Pass"))		'教辅审核
				End If

				tApplyer = Trim(rsTmp("Applyer"))
				tCourseList = tCourseList & "	<a class=""hr-rows hr-stretch"" href=""" & ParmPath & "SwapCourse/AlternDetails.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """>" & vbCrlf
				tCourseList = tCourseList & "		<div class=""hr-item iconTit""><i class=""hr-icon"">&#xe91c;</i></div>" & vbCrlf
				tCourseList = tCourseList & "		<div class=""hr-item hr-grow viewMSG"" data-id=""" & rsTmp("CourseID") & """>" & vbCrlf
				tCourseList = tCourseList & "			<p>申请教师：<span>" & tApplyer & "</span></p>" & vbCrlf
				tCourseList = tCourseList & "			<p>他的课程：" & FormatDate(rsTmp("VA4"), 4) & " 第" & rsTmp("VA7") & "节</p>" & vbCrlf
				tCourseList = tCourseList & "			<p>课程名称：<span>" & rsTmp("VA8") & " " & rsTmp("VA11") & "</span></p>" & vbCrlf
				tCourseList = tCourseList & "			<p>我的课程：" & FormatDate(rsTmp("newVA4"), 4) & " 第" & rsTmp("newVA7") & "节</p>" & vbCrlf
				tCourseList = tCourseList & "			<p>课程名称：<span>" & rsTmp("newVA8") & " " & rsTmp("newVA11") & "</span></p>" & vbCrlf				
				tCourseList = tCourseList & "		</div>" & vbCrlf
				tCourseList = tCourseList & "		<div class=""hr-item hr-fixed stepbar"">" & PassProcess(tProcess, tPass) & "</div>" & vbCrlf
				tCourseList = tCourseList & "		<div class=""hr-item more""><i class=""hr-icon"">&#xef8d;</i></div>" & vbCrlf
				tCourseList = tCourseList & "	</a>" & vbCrlf
				rsTmp.MoveNext
				i = i + 1
			Loop
			tCourseList = tCourseList & "</div>"
			tCourseList = "<h3 class=""hr-list-title""><i class=""hr-icon"">&#xe972;</i>您共有<b>" & i & "</b>条换课记录！</h3>" & tCourseList
		Else
			tCourseList = tCourseList & "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef61;</i></h2><h3>暂时没有老师与您交换课程！</h3></div>" & vbCrlf
		End If
	Set rsTmp = Nothing

	Response.Write "<div class=""weui-cells"">" & tCourseList & "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").html(""<i class='hr-icon'>&#xf0ca;</i>"");" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack em"").html(""<i class='hr-icon'>&#xf320;</i>"");" & vbCrlf	
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	var laynav=""<li><a href=\""" & ParmPath & "SwapCourse/Index.html\""><i class=\""hr-icon hr-icon-top\"">&#xe853;</i>我申请的换课</a></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "	laynav+=""<li><a href=\""" & ParmPath & "SwapCourse/Alternate.html\""><i class=\""hr-icon hr-icon-top\"">&#xf2dd;</i>我的换课记录</a></li>"";" & vbCrlf
	tmpHtml = tmpHtml & "	$("".nctouch-nav-menu ul"").html(laynav);" & vbCrlf				'更新右上角导航

	tmpHtml = tmpHtml & "</script>" & vbCrlf
	
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub Agree()			'换课老师确认
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tPassed : tPassed = HR_Clng(Request("Passed"))
	Dim tApplyer, tPasser : ErrMsg = ""
	sqlTmp = "Select * From HR_Swap Where ID=" & tmpID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			If HR_CLng(rsTmp("Replacer")) <> UserYGDM Then
				strTmp = "{""err"":true,""errcode"":500,""icon"":2,""errmsg"":""该课程是别的老师的！""}"
			Else
				Dim TeachDate, url1 : url1 = SiteUrl & "/Touch/SwapCourse/Details.html?ItemID=" & rsTmp("ItemID") & "&ID=" & tmpID
				tApplyer = strGetTypeName("HR_Teacher", "YGXM", "YGDM", HR_CLng(rsTmp("YGDM")))
				tPasser = strGetTypeName("HR_Teacher", "YGXM", "YGDM", HR_CLng(rsTmp("Passer")))
				TeachDate = FormatDate(rsTmp("VA4"), 2)
				If tPassed=1 Then
					ErrMsg = UserYGXM & "老师同意与您对换课程，操作时间：" & Now()
					Conn.Execute("Update HR_Swap Set ReplacerTime='" & Now() & "',ReplacePass=1,Process=1 Where ID=" & tmpID)
					strTmp = "{""err"":false,""errcode"":0,""icon"":1,""errmsg"":""您已经同意了对换课程！""}"
					Call SentWechatMSG_QYCard(UserYGDM, UserYGXM & "：您已同意与" & tApplyer & "老师对换课程", url1, UserYGXM & "老师：您已同意与" & tApplyer & "老师对换课程，等待教研主任" & tPasser & "审核。<br>授课时间：" & TeachDate & "<br>节次：第" & rsTmp("VA7") & "节<br>课程名称：" & rsTmp("VA8") & "<br>操作时间：" & FormatDate(Now(), 10))	'发送给自己
					Call SentWechatMSG_QYCard(HR_CLng(rsTmp("YGDM")), UserYGXM & "已同意与您交换课程", url1, tApplyer & "老师：" & UserYGXM  & "已同意与您对换课程，等待教研主任" & tPasser & "审核。<br>授课时间：" & TeachDate & "<br>节次：第" & rsTmp("VA7") & "节<br>课程名称：" & rsTmp("VA8") & "<br>操作时间：" & FormatDate(Now(), 10))	'发送给申请人
					ErrMsg = tApplyer & " 申请与" & UserYGXM  & " 对换课程，" & UserYGXM  & "已同意换课，请您审核。<br>申请课程：" & TeachDate & " 第" & rsTmp("VA7") & "节<br>对换课程：" & FormatDate(rsTmp("newVA4"), 2) & "第" & rsTmp("newVA7") & "节<br>申请时间：" & FormatDate(rsTmp("ApplyTime"), 10)
					ErrMsg = ErrMsg & "消息发送时间：" & FormatDate(Now(), 10)
					Call SentWechatMSG_QYCard(HR_CLng(rsTmp("Passer")), tApplyer & "申请与" & UserYGXM & "交换授课，请您审核", url1, ErrMsg)	'发送给教研主任
				ElseIf tPassed=2 Then
					ErrMsg = UserYGXM & "老师拒绝与您对换课程，操作时间：" & Now()
					Conn.Execute("Update HR_Swap Set ReplacerTime='" & Now() & "',ReplacePass=2,Process=1,Explain='" & ErrMsg & "' Where ID=" & tmpID)
					strTmp = "{""err"":false,""errcode"":0,""icon"":1,""errmsg"":""您已经拒绝了换课申请！""}"
					Call SentWechatMSG_QYCard(UserYGDM, UserYGXM & "：您已拒绝与" & tApplyer & "老师对换课程", url1, UserYGXM & "老师：您已拒绝与" & tApplyer & "老师对换课程。<br>授课时间：" & TeachDate & "<br>节次：第" & rsTmp("VA7") & "节<br>课程名称：" & rsTmp("VA8") & "<br>操作时间：" & FormatDate(Now(), 10))	'发送给自己
					Call SentWechatMSG_QYCard(HR_CLng(rsTmp("YGDM")), UserYGXM & "拒绝与您对换" & TeachDate & "的课程", url1, tApplyer & "老师：" & UserYGXM  & "已拒绝与您对换" & TeachDate & "的课程。<br>节次：第" & rsTmp("VA7") & "节<br>课程名称：" & rsTmp("VA8") & "<br>操作时间：" & FormatDate(Now(), 10))	'发送给申请人
				End If
			End If
		End If
	Set rsTmp = Nothing
	Response.Write strTmp
End Sub
%>