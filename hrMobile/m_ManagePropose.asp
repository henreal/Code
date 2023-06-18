<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./m_ManageCommon.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "意见建议管理"

Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "Delete" Call Delete()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	strHtml = "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		body {background-color:#f2f2f2;} .hr-header {background-color:#061;}" & vbCrlf
	strHtml = strHtml & "		.hr-panel-tips .sender {color:#0bd;width:5rem} .hr-panel-tips .del {color:#f30;width:2rem}" & vbCrlf
	strHtml = strHtml & "		.hr-panel-text {min-height:5rem;font-size:1.1rem;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf

	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ FastClick.attach(document.body); });" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim rsList, tList, tContent
	Set rsList = Conn.Execute("Select Top 200 a.*,(Select YGXM From HR_Teacher Where YGDM=a.YGDM) As YGXM From HR_Propose a Order By a.CreateTime DESC")
		If Not(rsList.BOF And rsList.EOF) Then
			Do While Not rsList.EOF
				tContent = Trim(rsList("Content")) : tContent = Replace(tContent, Chr(10), "<br>")
				tList = tList & "	<div class=""hr-panel-bar"">" & vbCrlf
				tList = tList & "		<dl class=""hr-rows hr-item-top hr-panel-item""><dt><i class=""hr-icon"">&#xf1d9;</i></dt>" & vbCrlf
				tList = tList & "			<dd class=""hr-row-fill hr-panel-text"">" & tContent & "</dd>" & vbCrlf
				'tList = tList & "			<dd class=""hr-panel-more""><i class=""hr-icon"">&#xf105;</i></dd>" & vbCrlf
				tList = tList & "		</dl>" & vbCrlf
				tList = tList & "		<ol class=""hr-rows hr-panel-tips""><li class=""hr-row-fill""><i class=""hr-icon"">&#xeedb;</i>" & FormatDate(rsList("CreateTime"), 1) & "</li>"
				tList = tList & "<li class=""sender"" data-id=""" & rsList("ID") & """><i class=""hr-icon"">&#xef2e;</i>" & rsList("YGXM") & "</li>"
				tList = tList & "<li class=""del"" data-id=""" & rsList("ID") & """><i class=""hr-icon"">&#xec9d;</i></li>"
				tList = tList & "</ol>" & vbCrlf
				tList = tList & "	</div>" & vbCrlf
				rsList.MoveNext
			Loop
		Else
			tList = tList & "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>还没有教师发表意见！</h3></div>" & vbCrlf
		End If
	Set rsList = Nothing
	Response.Write "<div class=""hr-panel-box"">" & vbCrlf
	Response.Write tList
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-shrink-x10""></div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Manage/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".del"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var tID = $(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "		$.confirm(""您确定要删除吗？"", function() {" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "ManagePropose/Delete.html"", {ID:tID}, function(reStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(reStr.reMessge, function(){location.reload(); });" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

Sub Delete()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Conn.Execute("Delete From HR_Propose Where ID in(" & tmpID & ")")
	ErrMsg = "删除成功！"
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub
%>