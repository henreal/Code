<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./m_ManageCommon.asp"-->
<!--#include file="../hrBase/incKPI.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl : SiteTitle = "对换课申请审核"
If ChkWechatTokenQY() = False Then Call GetWechatTokenQY()		'提前检查企业微信Token是否过期
If ChkTokenBobao() = False Then Call GetTokenBobao()			'检查信息播报Token

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "EditPass" Call EditPass()
	Case "SendPass" Call SendPass()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background:#fff;}" & vbCrlf
	tmpHtml = tmpHtml & "		.total {padding:3px 10px;} .total dl {font-size:1.2rem;} .total dd {font-size:1.5rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.list-title {padding:5px;border-bottom:1px solid #eee;} .list-title em {font-size:1.1rem;} .list-title em i {color:#4caf50;}" & vbCrlf
	tmpHtml = tmpHtml & "		.list-items {display:flex;align-items:center;padding:3px 5px;font-size:1.1rem;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.list-items dt {width:5.5rem; color:#999; text-align:right; font-size:16px; flex-shrink:0;} .list-items dd {flex-wrap:wrap; font-size:16px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.list-pass {padding:6px 3px; border-top:1px solid #ddd;}" & vbCrlf
	tmpHtml = tmpHtml & "		.list-pass .passbar {display:flex;align-items:center;}" & vbCrlf
	tmpHtml = tmpHtml & "		.list-pass .passbar em {padding-right:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.ReasonTxt {padding:5px 10px;color:#999;border-bottom:1px solid #eee;}" & vbCrlf

	tmpHtml = tmpHtml & "		.passed {color:#f30;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-dialog {max-width:initial ;} .weui-dialog__hd {padding:5px 8px;border-bottom:1px solid #eee;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-dialog__bd {padding:5px 8px;text-align:left;min-height:5rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.TipsTxt {font-size:1.2rem;color:#000} .Reason {font-size:1rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.passed {color:#f30;}" & vbCrlf
	
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim tCourseList, tStep, tPeriodTime, tReplacer, tItemName
	'取学年时间段
	Dim startTime, endTime
	startTime = DefYear-1 & "-07-01 00:00:00"
	endTime = DefYear & "-06-30 23:59:59"
	Dim SwapPass : SwapPass = HR_CLng(GetTypeName("HR_User", "SwapPass", "YGDM", UserYGDM))		'教学处或教辅审核
	sqlTmp = "Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.YGDM) As Proposer From HR_Swap a Where (a.ApplyTime Between '" & startTime & "' And '" & endTime & "') And a.newItemID>0 And a.newCourseID>0 And a.Passer=" & UserYGDM & " Order By ApplyTime DESC"		'取SQL(教研主任)
	If SwapPass = 1 Then
		sqlTmp = "Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.YGDM) As Proposer From HR_Swap a Where (a.ApplyTime Between '" & startTime & "' And '" & endTime & "') And a.newItemID>0 And a.newCourseID>0 And Process<3 Order By a.ApplyTime DESC"		'取SQL(教学处)
	ElseIf SwapPass = 2 Then
		sqlTmp = "Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.YGDM) As Proposer From HR_Swap a Where (a.ApplyTime Between '" & startTime & "' And '" & endTime & "') And a.newItemID>0 And a.newCourseID>0 Order By a.ApplyTime DESC"		'取SQL(教辅)
	End If
	Dim tProcess
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			tCourseList = tCourseList & "<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
			tCourseList = tCourseList & "<div class=""hr-list-ul"">" & vbCrlf
			Do While Not rsTmp.EOF
				tStep = "待审"
				tProcess = HR_CLng(rsTmp("Process"))
				If HR_IsNull(rsTmp("PassTime")) = False Then
					tStep = "审核中"
				ElseIf HR_IsNull(rsTmp("PassTime2")) = False Then
					tStep = "审核通过"
				End If

				tReplacer = strGetTypeName("HR_Teacher", "YGXM", "YGDM", HR_Clng(rsTmp("Replacer")))
				tCourseList = tCourseList & "<div class=""hr-list-wrap"">" & vbCrlf
				tCourseList = tCourseList & "	<div class=""hr-rows list-title"" data-id=""" & rsTmp("ID") & """ data-href=""" & ParmPath & "SwapCourse/Details.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """>" & vbCrlf
				tCourseList = tCourseList & "		<em><i class=""hr-icon"">&#xe853;</i>申请人：" & Trim(rsTmp("Proposer")) & "</em>" & vbCrlf
				tCourseList = tCourseList & "		<em class=""hr-grow"" data-id=""" & rsTmp("ID") & """><i class=""hr-icon"">&#xe8b5;</i>" & FormatDate(rsTmp("ApplyTime"), 10) & "</em>" & vbCrlf
				tCourseList = tCourseList & "		<em><i class=""hr-icon"">&#xf321;</i></em>" & vbCrlf
				tCourseList = tCourseList & "	</div>" & vbCrlf
				tCourseList = tCourseList & "	<div class=""ReasonTxt"">" & rsTmp("Reason") & "</div>" & vbCrlf
				'tCourseList = tCourseList & "	<dl class=""list-items""><dt>考核项目：</dt><dd>" & tItemName & "</dd></dl>" & vbCrlf
				tCourseList = tCourseList & "	<dl class=""list-items""><dt>申请课程：</dt><dd>" & FormatDate(rsTmp("VA4"), 4) & " 第" & rsTmp("VA7") & "节</dd></dl>" & vbCrlf
				tCourseList = tCourseList & "	<dl class=""list-items""><dt>课程名称：</dt><dd>" & rsTmp("VA8") & " " & rsTmp("VA11") & "</dd></dl>" & vbCrlf
				tCourseList = tCourseList & "	<dl class=""list-items""><dt>换课教师：</dt><dd>" & tReplacer & PassProcess(1, HR_CLng(rsTmp("ReplacePass"))) & "</dd></dl>" & vbCrlf
				tCourseList = tCourseList & "	<dl class=""list-items""><dt>对换课程：</dt><dd>" & FormatDate(rsTmp("newVA4"), 4) & " 第" & rsTmp("newVA7") & "节</dd></dl>" & vbCrlf
				tCourseList = tCourseList & "	<dl class=""list-items""><dt>课程名称：</dt><dd>" & rsTmp("newVA8") & " " & rsTmp("newVA11") & "</dd></dl>" & vbCrlf
	

				tCourseList = tCourseList & "	<div class=""hr-rows list-pass"">" & vbCrlf
				tCourseList = tCourseList & "		<div>状态：</div>" & vbCrlf
				tCourseList = tCourseList & "		<div class=""hr-grow passbar"">" & vbCrlf
				If tProcess = 5 Then
					tCourseList = tCourseList & "			<em>" & PassProcess(5, 0) & "</em>" & vbCrlf
				Else
					tCourseList = tCourseList & "			<em>教研主任" & PassProcess(2, rsTmp("PasserPass")) & "</em>" & vbCrlf
					tCourseList = tCourseList & "			<em>教学处" & PassProcess(3, rsTmp("Passer1Pass")) & "</em>" & vbCrlf
					tCourseList = tCourseList & "			<em>教辅" & PassProcess(4, rsTmp("Passer2Pass")) & "</em>" & vbCrlf
				End If
				tCourseList = tCourseList & "		</div>" & vbCrlf

				Dim passHref : passHref = False			'判断流程
				If HR_CLng(rsTmp("Passer")) = UserYGDM And tProcess = 1 Then		'判断教研主任
					If HR_CLng(rsTmp("PasserPass")) = 0 And tProcess = 1 Then passHref = True
				ElseIf SwapPass=1 Then							'判断是否为教学处
					If HR_CLng(rsTmp("Passer1Pass")) = 0 And tProcess = 2 Then passHref = True
				ElseIf SwapPass=2 Then							'判断是否为教辅
					If HR_CLng(rsTmp("Passer2Pass")) = 0 And tProcess = 3 And HR_CLng(rsTmp("Passer2")) = UserYGDM Then passHref = True
				End If
				If HR_CLng(rsTmp("ReplacePass")) = 2 Or HR_CLng(rsTmp("PasserPass")) = 2 Or HR_CLng(rsTmp("Passer1Pass")) = 2 Or tProcess=5 Then passHref = False		'当任一审核者拒绝时，即终止流程
				'tCourseList = tCourseList & "<em>" & SwapPass & "</em>"
				If passHref Then tCourseList = tCourseList & "		<div class=""passbtn"" data-id=""" & rsTmp("ID") & """><i class=""hr-icon"">&#xf321;</i></div>" & vbCrlf
				tCourseList = tCourseList & "	</div>" & vbCrlf
				tCourseList = tCourseList & "	<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
				tCourseList = tCourseList & "</div>" & vbCrlf
				rsTmp.MoveNext
				i = i + 1
			Loop
			tCourseList = tCourseList & "</div>" & vbCrlf
		End If
	Set rsTmp = Nothing

	Response.Write "<div class=""hr-fix total"">" & vbCrlf
	Response.Write "	<dl class=""hr-rows""><dt>代课申请：</dt><dd>" & HR_Clng(i) & "条</dd></dl>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-fix"">" & tCourseList & "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".list-title"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		location.href = $(this).data(""href"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$("".passbtn"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var swapid=$(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		location.href=""" & ParmPath & "SwapPass/EditPass.html?ID=""+swapid;" & vbCrlf
	'tmpHtml = tmpHtml & "		$.modal({title: ""审核操作"",text:""<p class='TipsTxt'>请您选择是否同意该教师的调换课申请？</p><p class='Reason'>原因："" + $(""#ReasonTxt""+ swapid).html() + ""</p>"",buttons:[{text:""同意"", onClick: function(){" & vbCrlf
	'tmpHtml = tmpHtml & "				$.getJSON(""" & ParmPath & "SwapPass/SendPass.html"",{ID:swapid, Passed:""True""},function(reData){$.toast(reData.reMessge)});" & vbCrlf
	'tmpHtml = tmpHtml & "			} },{text:""拒绝"", onClick: function(){" & vbCrlf
	'tmpHtml = tmpHtml & "				$.getJSON(""" & ParmPath & "SwapPass/SendPass.html"",{ID:swapid, Passed:""False""},function(reData){" & vbCrlf
	'tmpHtml = tmpHtml & "					$.toast(reData.reMessge, ""forbidden"",function(){ location.reload(); });" & vbCrlf
	'tmpHtml = tmpHtml & "				});" & vbCrlf
	'tmpHtml = tmpHtml & "			} },{text:""关闭""}]" & vbCrlf
	'tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub SendPass()
	Dim rsSave, tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tPassed : tPassed = HR_Clng(Request("Passed"))		'1通过2拒绝0未审
	Dim tAssistant : tAssistant = HR_Clng(Request("AssistantCode"))		'指定教辅工号
	Dim SwapPass : SwapPass = HR_CLng(GetTypeName("HR_User", "SwapPass", "YGDM", UserYGDM))		'教学处或教辅审核权
	Dim tReplacer, tVA4

	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open("Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.YGDM) As Proposer From HR_Swap a Where a.newItemID>0 And a.newCourseID>0 And a.ID=" & tmpID), Conn, 1, 3
		If Not(rsSave.BOF And rsSave.EOF) Then
			tReplacer = strGetTypeName("HR_Teacher", "YGXM", "YGDM", HR_Clng(rsSave("Replacer")))
			tVA4 = Trim(rsSave("VA4"))
			If SwapPass = 1 And HR_CLng(rsSave("Process"))=2 Then			'教学处审核（必须教研主任先审核）
				If tPassed=1 Then
					rsSave("Passer1") = UserYGDM
					rsSave("PassTime1") = Now()
					rsSave("Passer1Pass") = 1
					rsSave("Passer2") = tAssistant
					rsSave("Process") = 3
					ErrMsg = "教学处已审核！"
					Response.Write "{""err"":false,""icon"":1,""errcode"":0,""errmsg"":""您已同意了代课申请！""}"
				ElseIf tPassed=2 Then
					rsSave("Passer1") = UserYGDM
					rsSave("PassTime1") = Now()
					rsSave("Process") = 3
					rsSave("Passer1Pass") = 2
					ErrMsg = "教学处拒绝申请！"
					Response.Write "{""err"":false,""icon"":1,""errcode"":0,""errmsg"":""您已拒绝了代课申请！""}"
				End If
				rsSave.Update
				If tAssistant > 0 Then		'发送消息给教辅
					ErrMsg = "" & UserYGXM & "老师结转给您" & Trim(rsSave("Proposer")) & "老师的换课申请。<br>申请人：" & Trim(rsSave("Proposer")) & "<br>授课时间：" & FormatDate(tVA4, 4) & "<br>课程名称：" & rsSave("VA8") & "<br>代课教师：" & tReplacer & "<br>发送时间：" & FormatDate(Now(), 10)
					Call SentWechatMSG_QYCard(tAssistant,"" & UserYGXM & "老师结转给您代课申请，请查阅！", SiteUrl & "/Touch/SwapCourse/Details.html?ItemID=" & rsSave("ItemID") & "&ID=" & tmpID, ErrMsg)
				End If
			ElseIf rsSave("Passer2") = UserYGDM And HR_CLng(rsSave("Process"))=3 Then		'教辅审核后直接替换课程数据
				If tPassed=1 Then
					rsSave("PassTime2") = Now()
					rsSave("Process") = 4
					rsSave("Passer2Pass") = 1
					rsSave.Update

					'更新课程数据及业绩汇总
					Dim tItemID, tCourseID, tCourseID1, tSheetName, tSheetName1
					tItemID = rsSave("ItemID")
					tCourseID = rsSave("CourseID")
					tCourseID1 = rsSave("newCourseID")
					tSheetName = "HR_Sheet_" & tItemID		'数据表名
					tSheetName1 = "HR_Sheet_" & rsSave("newItemID")
					If IsDate(tVA4) Then
						tVA4 = ConvertDateToNum(tVA4) + 2
					End If
					If ChkTable(tSheetName) And ChkTable(tSheetName1) Then
						Set rs = Server.CreateObject("ADODB.RecordSet")
							rs.Open("Select * From " & tSheetName & " Where ID=" & tCourseID), Conn, 1, 3
							If Not(rs.BOF And rs.EOF) Then
								rs("VA1") = HR_CLng(rsSave("Replacer"))
								rs("VA2") = tReplacer
								rs.Update
							End If
						Set rs = Nothing
						Set rs = Server.CreateObject("ADODB.RecordSet")
							rs.Open("Select * From " & tSheetName1 & " Where ID=" & tCourseID1), Conn, 1, 3
							If Not(rs.BOF And rs.EOF) Then
								rs("VA1") = HR_CLng(rsSave("YGDM"))
								rs("VA2") = Trim(rsSave("Proposer"))
								rs.Update
							End If
						Set rs = Nothing

						Call ChkTeacherKPI(rsSave("YGDM"))			'更新原老师KPI
						Call UpdateTeacherKPI(tItemID, rsSave("YGDM"), "")
						Call UpdateTeacherTotalKPI(rsSave("YGDM"))

						Call ChkTeacherKPI(rsSave("Replacer"))		'更新替换老师KPI
						Call UpdateTeacherKPI(tItemID, rsSave("Replacer"), "")
						Call UpdateTeacherTotalKPI(rsSave("Replacer"))
					End If
					ErrMsg = rsSave("Proposer") & "老师：您的换课申请已经审核完成！<br>" & FormatDate(rsSave("VA4"), 4) & " 第" & rsSave("VA7") & "节的课程由" & tReplacer & "老师授课。点击查看详情。"
					ErrMsg = ErrMsg & "<br>申请课程：" & FormatDate(rsSave("VA4"), 4) & " 第" & rsSave("VA7") & "节<br>课程内容：" & Trim(rsSave("VA8")) & ""
					ErrMsg = ErrMsg & "<br>换课老师：" & tReplacer & "<br>交换课程：" & FormatDate(rsSave("newVA4"), 4) & " 第" & rsSave("newVA7") & "节<br>课程内容：" & Trim(rsSave("newVA8")) & "<br>发送时间：" & FormatDate(Now(), 10)
					Call SentWechatMSG_QYCard(rsSave("YGDM"), "您的代课申请已经审核完成！", SiteUrl & "/Touch/SwapCourse/Details.html?ItemID=" & rsSave("ItemID") & "&ID=" & tmpID, ErrMsg)		'发送给申请人

					ErrMsg = tReplacer & "老师：" & rsSave("Proposer") & "与您换课的申请已经完成审核！<br>" & FormatDate(rsSave("VA4"), 4) & " 第" & rsSave("VA7") & "节的课程由您授课，而" & FormatDate(rsSave("newVA4"), 4) & " 第" & rsSave("newVA7") & "节的课程由" & rsSave("Proposer") & "老师上课。点击查看详情。"
					ErrMsg = ErrMsg & "<br>申请人：" & rsSave("Proposer") & "<br>申请课程：" & FormatDate(rsSave("VA4"), 4) & " 第" & rsSave("VA7") & "节<br>课程内容：" & Trim(rsSave("VA8"))
					ErrMsg = ErrMsg & "<br>交换课程：" & FormatDate(rsSave("newVA4"), 4) & " 第" & rsSave("newVA7") & "节<br>课程内容：" & Trim(rsSave("newVA8")) & "<br>发送时间：" & FormatDate(Now(), 10)
					Call SentWechatMSG_QYCard(rsSave("YGDM"), "您的换课申请已经审核完成！", SiteUrl & "/Touch/SwapCourse/Details.html?ItemID=" & rsSave("ItemID") & "&ID=" & tmpID, ErrMsg)		'发送给换课人
					Response.Write "{""err"":false,""icon"":1,""errcode"":0,""errmsg"":""您已同意了换课申请！""}"
				ElseIf tPassed=2 Then		'教辅拒绝
					rsSave("Process") = 4
					rsSave("Passer2Pass") = 2
					rsSave("PassTime2") = Now()
					rsSave.Update
					ErrMsg = "" & rsSave("Proposer") & "老师：您" & FormatDate(tVA4, 4) & "的换课申请已被拒绝！<br>课程名称：" & rsSave("VA8") & "节次：第" & rsTmp("VA7") & "节<br>换课教师：" & tReplacer & "。<br>发送时间：" & FormatDate(Now(), 10)
					Call SentWechatMSG_QYCard(rsSave("YGDM"), "您的换课申请已被拒绝！", SiteUrl & "/Touch/SwapCourse/Details.html?ItemID=" & rsSave("ItemID") & "&ID=" & tmpID, ErrMsg)
					Response.Write "{""err"":false,""icon"":1,""errcode"":0,""errmsg"":""您已拒绝了换课申请！""}"
				End If
			ElseIf HR_CLng(rsSave("Passer"))=UserYGDM And HR_CLng(rsSave("Process"))=1 Then				'教研主任审核
				If tPassed=1 Then
					rsSave("PassTime") = Now()
					rsSave("PasserPass") = 1
					rsSave("Process") = 2
					rsSave("Explain") = UserYGXM & " 同意了" & rsSave("Proposer") & "的换课申请！"
					rsSave.Update
					ErrMsg = rsSave("Proposer") & "老师：您的换课申请教研主任" & UserYGXM & "已审核！<br>申请课程：" & FormatDate(tVA4, 4) & " 第" & rsSave("VA7") & "节<br>课程名称：" & rsSave("VA8") & "<br>换课教师：" & tReplacer & "<br>对换课程：" & FormatDate(rsSave("newVA4"), 4) & "第" & rsSave("newVA7") & "节<br>发送时间：" & FormatDate(Now(), 10)
					Call SentWechatMSG_QYCard(rsSave("YGDM"), UserYGXM & " 同意了您的代课申请！", SiteUrl & "/Touch/SwapCourse/Details.html?ItemID=" & rsSave("ItemID") & "&ID=" & tmpID, ErrMsg)
					Response.Write "{""err"":false,""icon"":1,""errcode"":0,""errmsg"":""您已同意了代课申请！""}"
				ElseIf tPassed=2 Then
					rsSave("Process") = 3
					rsSave("PassTime") = Now()
					rsSave("PasserPass") = 2
					rsSave("Explain") = UserYGXM & " 拒绝了" & rsSave("Proposer") & "的换课申请！"
					rsSave.Update
					ErrMsg = rsSave("Proposer") & "老师：教研主任" & UserYGXM & "拒绝了您的换课申请！<br>申请课程：" & FormatDate(tVA4, 4) & " 第" & rsSave("VA7") & "节<br>课程名称：" & rsSave("VA8") & "<br>换课教师：" & tReplacer & "<br>对换课程：" & FormatDate(rsSave("newVA4"), 4) & "第" & rsSave("newVA7") & "节<br>发送时间：" & FormatDate(Now(), 10)
					Call SentWechatMSG_QYCard(rsSave("YGDM"), UserYGXM & " 拒绝了您的调换课申请！", SiteUrl & "/Touch/SwapCourse/Details.html?ItemID=" & rsSave("ItemID") & "&ID=" & tmpID, ErrMsg)
					Response.Write "{""err"":false,""icon"":1,""errcode"":0,""errmsg"":""您已拒绝了换课申请！""}"
				End If
			End If
			
		Else
			ErrMsg = "换课申请不存在！ID:" & tmpID
			Response.Write "{""err"":true,""icon"":2,""errcode"":500,""errmsg"":""" & ErrMsg & """}"
		End If
	Set rsSave = Nothing
End Sub

Sub EditPass()
	Dim tmpID : tmpID = HR_CLng(Request("ID"))
	Dim tItemID
	Dim SwapPass : SwapPass = HR_CLng(GetTypeName("HR_User", "SwapPass", "YGDM", UserYGDM))		'教学处或教辅审核权限

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background:#fff}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-gap-10 {height:10px}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf

	sqlTmp = "Select a.*,b.YGXM,b.KSMC,b.PRZC,b.XZZW,(Select YGXM From HR_Teacher Where YGDM=a.YGDM) As Proposer"
	sqlTmp = sqlTmp & ",(Select KSMC From HR_Teacher Where YGDM=a.YGDM) As ProposerKS,(Select PRZC From HR_Teacher Where YGDM=a.YGDM) As ProposerZC,(Select XZZW From HR_Teacher Where YGDM=a.YGDM) As ProposerZW"
	sqlTmp = sqlTmp & ",(Select ClassName From HR_Class Where ClassID=a.ItemID) As ItemName"
	sqlTmp = sqlTmp & " From HR_Swap a Left Join HR_Teacher b On a.Replacer=b.YGDM Where a.ID=" & tmpID
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemID = HR_CLng(rsTmp("ItemID"))
			Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">申请人：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("Proposer")) & " [" & Trim(rsTmp("YGDM")) & "]</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">申请理由：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("Reason")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""hr-gap-20 hr-gap-10 hr-gapbg""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">科室：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("ProposerKS")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">职务：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("ProposerZW")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">职称：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("ProposerZC")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""hr-gap-20 hr-gap-10 hr-gapbg""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">项目名称：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("ItemName")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">申请课程：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & FormatDate(rsTmp("VA4"), 4) & " 第" & Trim(rsTmp("VA7")) & "节</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">学时：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & HR_CDbl(rsTmp("VA3")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">星期：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("VA6")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">周次：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("VA5")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">课程名称：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("VA8")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课内容：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("VA9")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课对象：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("VA10")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课教室：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("VA11")) & "" & Trim(rsTmp("VA12")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20 hr-gapbg""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">换课老师：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("YGXM")) & "[" & Trim(rsTmp("YGDM")) & "]</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">科室：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("KSMC")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">职务：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("XZZW")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">职称：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("PRZC")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""hr-gap-20 hr-gap-10 hr-gapbg""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">对换课程：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & FormatDate(rsTmp("newVA4"), 4) & " 第" & Trim(rsTmp("newVA7")) & "节</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">学时：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & HR_CDbl(rsTmp("newVA3")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">星期：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("newVA6")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">周次：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("newVA5")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">课程名称：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("newVA8")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课内容：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("newVA9")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课对象：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("newVA10")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">校(院)区：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
			Response.Write "			<input name=""VA11"" class=""weui-input"" id=""VA11"" type=""text"" value=""" & Trim(rsTmp("newVA11")) & """>" & vbCrlf
			Response.Write "		</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课教室：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & Trim(rsTmp("newVA11")) & " " & Trim(rsTmp("newVA12")) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<input type=""hidden"" name=""id"" value=""" & tmpID & """>" & vbCrlf
			'若为教学处，选择教辅
			If SwapPass = 1 Then
				Response.Write "	<div class=""hr-gap-20 hr-gapbg""></div>" & vbCrlf
				Response.Write "	<div class=""weui-cell"">" & vbCrlf
				Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">结转教辅：</label></div>" & vbCrlf
				Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
				Response.Write "			<input name=""Assistant"" class=""weui-input"" id=""Assistant"" type=""text"" value="""">" & vbCrlf
				Response.Write "		</div>" & vbCrlf
				Response.Write "	</div>" & vbCrlf
			End If

			Response.Write "</div>" & vbCrlf
			Response.Write "<div class=""hr-shrink-x20""></div>" & vbCrlf
			Response.Write "<div class=""hr-shrink-x10""></div>" & vbCrlf
			Response.Write "<div class=""hr-rows hr-editbtn"">" & vbCrlf
			Response.Write "	<em><i class=""hr-icon"">&#xea3f;</i></em>" & vbCrlf
			If UserYGDM = HR_CLng(rsTmp("Passer")) And HR_CLng(rsTmp("PasserPass"))=0 Then
				Response.Write "	<em><button type=""button"" name=""pass"" class=""passbtn"" value=""1"" data-id=""" & tmpID & """>同意</button></em>" & vbCrlf
				Response.Write "	<em><button type=""button"" name=""edit"" class=""passbtn"" value=""2"" data-id=""" & tmpID & """>拒绝</button></em>" & vbCrlf
			ElseIf SwapPass = 1 And HR_CLng(rsTmp("Passer1Pass"))=0 Then
				Response.Write "	<em><button type=""button"" name=""pass"" class=""passbtn"" value=""1"" data-id=""" & tmpID & """>同意</button></em>" & vbCrlf
				Response.Write "	<em><button type=""button"" name=""edit"" class=""passbtn"" value=""2"" data-id=""" & tmpID & """>拒绝</button></em>" & vbCrlf
			ElseIf SwapPass = 2 And HR_CLng(rsTmp("Passer2Pass"))=0 And HR_CLng(rsTmp("Passer2"))=UserYGDM Then
				Response.Write "	<em><button type=""button"" name=""pass"" class=""passbtn"" value=""1"" data-id=""" & tmpID & """>同意</button></em>" & vbCrlf
				Response.Write "	<em><button type=""button"" name=""edit"" class=""passbtn"" value=""2"" data-id=""" & tmpID & """>拒绝</button></em>" & vbCrlf
			End If
			Response.Write "	<em><button type=""button"" name=""retreat"" class=""retreat"" id=""Retreat"" data-id=""" & tmpID & """>返回</button></em>" & vbCrlf
			Response.Write "</div>" & vbCrlf
		Else
			Response.Write "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef61;</i></h2><h3>换课申请被删了！</h3></div>" & vbCrlf
		End If
	Set rsTmp = Nothing
	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "SwapPass/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#CourseDate"").calendar({dateFormat: 'yyyy-mm-dd'});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#VA8"").select({" & vbCrlf
	tmpHtml = tmpHtml & "		title: ""请选择"",items:[" & GetCourseSelect("VA8", "") & "]" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#VA10"").select({" & vbCrlf
	tmpHtml = tmpHtml & "		title: ""请选择"",items:[" & getFieldSelect(tItemID, "VA10", "") & "]" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#VA11"").select({" & vbCrlf
	tmpHtml = tmpHtml & "		title: ""请选择"",items:[" & GetCampusSelect("VA11", "") & "]" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#VA12"").select({" & vbCrlf
	tmpHtml = tmpHtml & "		title: ""请选择"",items:[" & GetClassRoomSelect("VA12", "") & "]" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	If SwapPass = 1 Then
		Dim tAssistant
		Set rsTmp = Conn.Execute("Select a.* From HR_User a Where a.SwapPass=2")
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				i = 0
				Do While Not rsTmp.EOF
					If i>0 Then tAssistant = tAssistant & ","
					tAssistant = tAssistant & "{""title"":""" & rsTmp("YGXM") & """,""value"":""" & rsTmp("YGDM") & """}"
					rsTmp.MoveNext
					i = i + 1
				Loop
			End If
		Set rsTmp = Nothing
		tmpHtml = tmpHtml & "	$(""#Assistant"").select({" & vbCrlf
		tmpHtml = tmpHtml & "		title: ""请选择教辅"",items:[" & tAssistant & "]" & vbCrlf
		tmpHtml = tmpHtml & "	});" & vbCrlf
	End If
	tmpHtml = tmpHtml & "	$("".passbtn"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var passed = $(this).val();" & vbCrlf
	tmpHtml = tmpHtml & "		var pass1=" & SwapPass & ", assistant = $(""#Assistant"").data(""values"");" & vbCrlf			'教辅工号
	tmpHtml = tmpHtml & "		if(passed==1&&pass1==1){" & vbCrlf			'若为教学处时
	tmpHtml = tmpHtml & "			if(!assistant){ $.toast(""请选择教辅！"",""forbidden""); return false; };" & vbCrlf
	tmpHtml = tmpHtml & "		}" & vbCrlf

	tmpHtml = tmpHtml & "		$.getJSON(""" & ParmPath & "SwapPass/SendPass.html?Passed="" + passed + ""&AssistantCode="" + assistant, $(""#EditForm"").serialize(), function(res){" & vbCrlf
	tmpHtml = tmpHtml & "			$.toast(res.errmsg, function(){" & vbCrlf
	tmpHtml = tmpHtml & "				if(!res.err){ location.href=""" & ParmPath & "SwapPass/Index.html""; }" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#Retreat"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		location.href=""" & ParmPath & "SwapPass/Index.html"";" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
%>