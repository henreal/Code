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
Dim scriptCtrl : SiteTitle = "修改申请"

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "ShowList" Call ShowList()
	Case "View" Call View()
	Case "AgreeModify" Call AgreeModify()
	Case "RefuseModify" Call RefuseModify()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.viewPanel li.hr-gap-20 {background-color:#eee;} .viewPanel li {line-height:inherit;padding:5px 0;border:0}" & vbCrlf
	tmpHtml = tmpHtml & "		.viewPanel li b {width:29%;text-align:right;color:#999}" & vbCrlf
	tmpHtml = tmpHtml & "		.viewPanel li em {width:70%;color:#000}" & vbCrlf

	tmpHtml = tmpHtml & "		.btnBar {box-sizing:border-box;padding:10px 0;} .btnBar em {width:50%;box-sizing:border-box;padding:0 5px}" & vbCrlf
	tmpHtml = tmpHtml & "		.view {align-items: inherit;}" & vbCrlf
	tmpHtml = tmpHtml & "		.viewbox, .popbtn {box-sizing:border-box;background-color:#fff;} .popbtn {padding:10px 5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.closebtn {width:2rem;text-align:center;height:2rem;line-height:2rem;position:fixed;bottom:3rem;right:1rem;background-color:rgba(3,169,244,0.5);color:#fff;z-index:100;display:none;border-radius: 5px;font-size:1.3rem}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", tmpHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ FastClick.attach(document.body); });" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"" id=""ListBox""></div>" & vbCrlf
	Response.Write "<div id=""fullView"" class=""weui-popup__container"">" & vbCrlf
	Response.Write "	<div class=""weui-popup__overlay""></div>" & vbCrlf
	Response.Write "	<div class=""weui-popup__modal""><div class=""viewbox""></div><div class=""Coursebox""></div><div class=""popbtn""><button class=""weui-btn weui-btn_primary close-popup"">关闭</button></div></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf	
	Response.Write "<div class=""hr-shrink-x20""></div>" & vbCrlf
	Response.Write "<div class=""closebtn close-popup"" title=""关闭弹窗""><i class=""hr-icon"">&#xee30;</i></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Manage/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	ReportItem("""");" & vbCrlf

	strHtml = strHtml & "	function ReportItem(fygdm){" & vbCrlf		'异步加载列表
	strHtml = strHtml & "		$.get(""" & ParmPath & "ManageModify/ShowList.html"",{ygdm:fygdm}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "			$(""#ListBox"").html(rsStr);" & vbCrlf
	strHtml = strHtml & "			$("".view"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "				$.get(""" & ParmPath & "ManageModify/View.html"",{ID:$(this).data(""applyid"")}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "					$("".viewbox"").html(rsStr);" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "				$.get(""" & ParmPath & "Course/ShowCourse.html"",{ItemID:$(this).data(""item""),ID:$(this).data(""id"")}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "					$("".Coursebox"").html(rsStr);$(""#fullView"").popup();$("".closebtn"").show();" & vbCrlf
	strHtml = strHtml & "					$("".close-popup"").on(""click"",function(){$("".closebtn"").hide();});" & vbCrlf		'关闭弹窗

	strHtml = strHtml & "					$("".btnManage"").on(""click"",function(){" & vbCrlf			'执行同意/拒绝
	strHtml = strHtml & "						var evname = $(this).attr(""name"");" & vbCrlf
	strHtml = strHtml & "						if(evname==""Agree""){" & vbCrlf							'同意
	strHtml = strHtml & "							$.getJSON(""" & ParmPath & "ManageModify/AgreeModify.html"",{ID:$(this).data(""id"")}, function(rsData){" & vbCrlf
	strHtml = strHtml & "								$.toast(rsData.errmsg,function(){ location.reload(); });" & vbCrlf
	strHtml = strHtml & "							});" & vbCrlf
	strHtml = strHtml & "						}else if(evname==""Refuse""){" & vbCrlf
	strHtml = strHtml & "							$.getJSON(""" & ParmPath & "ManageModify/RefuseModify.html"",{ID:$(this).data(""id"")}, function(rsData){" & vbCrlf
	strHtml = strHtml & "								$.toast(rsData.errmsg,function(){ location.reload(); });" & vbCrlf
	strHtml = strHtml & "							});" & vbCrlf
	strHtml = strHtml & "						}" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	}" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub ShowList()
	sqlTmp = "Select Top 300 a.*,b.YGXM As Sender,b.KSMC,b.PRZC"
	sqlTmp = sqlTmp & ",(Select ClassName From HR_Class Where ClassID=a.ItemID) As ItemName"
	sqlTmp = sqlTmp & ",(Select Template From HR_Class Where ClassID=a.ItemID) As Template"
	sqlTmp = sqlTmp & " From HR_Message a Left Join HR_Teacher b On a.SenderID=b.YGDM Where a.MsgType=1 And a.SenderID>0 And a.ItemID>0 Order By SendTime DESC"
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				Response.Write "	<a class=""weui-cell weui-cell_access view"" data-item=""" & rsTmp("ItemID") & """ data-id=""" & rsTmp("CourseID") & """ data-applyid=""" & rsTmp("ID") & """ href=""#"">" & vbCrlf
				Response.Write "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xef2f;</i></div><div class=""weui-cell__bd weui-cell_primary"">" & vbCrlf
				Response.Write "			<p>申请人：" & rsTmp("Sender") & " [" & rsTmp("SenderID") & "] " & rsTmp("KSMC") & "</p><p>理　由：" & rsTmp("Message") & "</p><p>项　目：" & rsTmp("ItemName") & "</p>" & vbCrlf
				Response.Write "			<p>时间：" & FormatDate(rsTmp("SendTime"),10) & "</p>" & vbCrlf
				Response.Write "		</div>" & vbCrlf
				Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
				Response.Write "	</a>" & vbCrlf
				rsTmp.MoveNext
			Loop
		Else
			Response.Write "	<a class=""weui-cell weui-cell_access"" href=""javascript:;"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>暂时没有记录</p></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
			Response.Write "	</a>" & vbCrlf
		End If
	Set rsTmp = Nothing
End Sub

Sub View()
	Dim tmpID : tmpID = HR_CLng(Request("ID"))
	Dim tmpMSG
	sqlTmp = "Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.SenderID) As Sender From HR_Message a Where ID=" & tmpID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tmpMSG = tmpMSG & "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
			tmpMSG = tmpMSG & "<ul class=""viewPanel"">" & vbCrlf
			tmpMSG = tmpMSG & "	<li class=""info listItem""><b>申请人：</b><em>" & rsTmp("Sender") & "</em></li>" & vbCrlf
			tmpMSG = tmpMSG & "	<li class=""info listItem""><b>申请时间：</b><em>" & FormatDate(rsTmp("SendTime"), 10) & "</em></li>" & vbCrlf
			tmpMSG = tmpMSG & "	<li class=""info listItem""><b>修改原因：</b><em>" & Trim(rsTmp("Message")) & "</em></li>" & vbCrlf
			tmpMSG = tmpMSG & "</ul>" & vbCrlf
			If HR_CBool(rsTmp("isRead")) = False Then
				tmpMSG = tmpMSG & "<div class=""hr-rows btnBar"">" & vbCrlf
				tmpMSG = tmpMSG & "	<em><button type=""button"" name=""Agree"" class=""weui-btn weui-btn_primary btnManage"" data-id=""" & tmpID & """>同意</button></em>" & vbCrlf
				tmpMSG = tmpMSG & "	<em><button type=""button"" name=""Refuse"" class=""weui-btn weui-btn_warn btnManage"" data-id=""" & tmpID & """>拒绝</button></em>" & vbCrlf
				tmpMSG = tmpMSG & "</div>" & vbCrlf
			End If
		End If
	Set rsTmp = Nothing
	Response.Write tmpMSG
	If ChkWechatTokenQY() = False Then Call GetWechatTokenQY()		'提前检查微信TokenQY
End Sub

Sub AgreeModify()
	Dim tmpID : tmpID = HR_CLng(Request("ID"))
	sqlTmp = "Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.SenderID) As Sender From HR_Message a Where ID=" & tmpID
	Set rsTmp = Conn.Execute(sqlTmp)
		Dim tSheetName
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tSheetName = "HR_Sheet_" & rsTmp("ItemID")
			If ChkTable(tSheetName) Then
				Conn.Execute("Update HR_Message Set isRead=" & HR_True & ",ReadTime=getDate() Where ID=" & tmpID )
				Conn.Execute("Update " & tSheetName & " Set Passed=" & HR_False & " Where ID=" & rsTmp("CourseID"))
				ErrMsg = "您的课程修改申请已经通过"
				ErrMsg = SentWechatMSG_QYCard(rsTmp("SenderID"), rsTmp("Sender") & "老师，您修改课程的申请已经通过！", SiteUrl & "/Touch/Course/View.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("CourseID"), "请点击详情进入课程修改。<br>发送时间：" & FormatDate(Now(), 1))
				Response.Write "{""err"":false,""errcode"":0,""errmsg"":""审核通过"",""icon"":1}" : Exit Sub
			End If
		End If
	Set rsTmp = Nothing
	Response.Write "{""err"":true,""errcode"":500,""errmsg"":""The data emptied"",""icon"":2}" : Exit Sub
End Sub

Sub RefuseModify()
	Dim tmpID : tmpID = HR_CLng(Request("ID"))
	sqlTmp = "Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.SenderID) As Sender From HR_Message a Where ID=" & tmpID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Conn.Execute("Update HR_Message Set isRead=" & HR_True & ",ReadTime=getDate() Where ID=" & tmpID )
			ErrMsg = "您的课程修改申请未能通过"
			ErrMsg = SentWechatMSG_QYCard(rsTmp("SenderID"), rsTmp("Sender") & "老师，您修改课程的申请被拒绝！", SiteUrl & "/Touch/Course/View.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("CourseID"), "您修改课程的申请被拒绝，若有问题请联系管理员。<br>请点击详情查看课程。<br>发送时间：" & FormatDate(Now(), 1))
			Response.Write "{""err"":false,""errcode"":0,""errmsg"":""申请已拒绝"",""icon"":1}" : Exit Sub
		End If
	Set rsTmp = Nothing
	Response.Write "{""err"":true,""errcode"":500,""errmsg"":""拒绝修改"",""icon"":2}" : Exit Sub
End Sub
%>