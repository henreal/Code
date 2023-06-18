<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "意见建议"

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "SavePost" Call SavePost()
	Case "Delete" Call Delete()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	strHtml = "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		body {background-color:#f2f2f2;}" & vbCrlf
	strHtml = strHtml & "		.navExtend {height: initial;flex-grow:2;text-align:right;}" & vbCrlf
	strHtml = strHtml & "		.navExtend span {font-size:1.2rem;display:line-block;background-color:#f7ce93;padding:2px 3px;color:#035;border-radius: 2px}" & vbCrlf

	strHtml = strHtml & "		.hr-panel-tips li:first-child i {color:#0af;} .hr-panel-tips .del {color:#f30;}" & vbCrlf
	strHtml = strHtml & "		.hr-panel-tips li.del {padding-right:20px;}" & vbCrlf
	strHtml = strHtml & "		.toolbar, .toolbar .title {font-size:20px;}" & vbCrlf
	strHtml = strHtml & "		.hr-header {z-index:8;} .weui-toast {margin-left: auto;}" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn {width:55px;height:55px;text-align:center;font-size:2.6rem;position: fixed;right: 1rem;bottom: 3rem;z-index: 9;}" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn i {color:#814ee2;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf

	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ FastClick.attach(document.body); });" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim tList, rsList, tContent
	Set rsList = Conn.Execute("Select * From HR_Propose Where YGDM=" & UserYGDM & " Order By CreateTime DESC")
		If Not(rsList.BOF And rsList.EOF) Then
			Do While Not rsList.EOF
				tContent = Trim(rsList("Content")) : tContent = Replace(tContent, Chr(10), "<br>")
				tList = tList & "	<div class=""hr-panel-bar"">" & vbCrlf
				tList = tList & "		<dl class=""hr-rows hr-item-top hr-panel-item""><dt><i class=""hr-icon"">&#xf1d9;</i></dt>" & vbCrlf
				tList = tList & "			<dd class=""hr-row-fill hr-panel-text"">" & tContent & "</dd>" & vbCrlf
				tList = tList & "			<dd class=""hr-panel-more""><i class=""hr-icon"">&#xf105;</i></dd>" & vbCrlf
				tList = tList & "		</dl>" & vbCrlf
				tList = tList & "		<ol class=""hr-rows hr-panel-tips""><li><i class=""hr-icon"">&#xeedb;</i>" & FormatDate(rsList("CreateTime"), 15) & "</li>"
				tList = tList & "<li class=""del"" data-id=""" & rsList("ID") & """><i class=""hr-icon"">&#xec9d;</i></li>"
				tList = tList & "</ol>" & vbCrlf
				tList = tList & "	</div>" & vbCrlf
				rsList.MoveNext
			Loop
		Else
			tList = tList & "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>您还没有发表意见！</h3></div>" & vbCrlf
		End If
	Set rsList = Nothing
	Response.Write "<div class=""hr-panel-box"">" & vbCrlf & tList & "</div>" & vbCrlf
	Response.Write "<div class=""hr-float-btn"">" & vbCrlf
	Response.Write "	<span class=""navExtend open-popup"" data-target=""#editpop""><i class=""hr-icon"">&#xf3c0;</i></span>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div id=""editpop"" class=""weui-popup__container popup-bottom"">" & vbCrlf
	Response.Write "	<div class=""weui-popup__overlay""></div>" & vbCrlf
	Response.Write "	<div class=""weui-popup__modal"">" & vbCrlf
	Response.Write "		<div class=""toolbar""><div class=""toolbar-inner""><span class=""picker-button close-popup"">关闭</span><h1 class=""title"">签写建议内容</h1></div></div>" & vbCrlf
	Response.Write "		<div class=""modal-content"">" & vbCrlf
	Response.Write "			<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "				<div class=""weui-cell"">" & vbCrlf
	Response.Write "					<div class=""weui-cell__bd""><textarea class=""weui-textarea"" placeholder=""请输入建议内容"" rows=""5""></textarea></div>" & vbCrlf
	Response.Write "				</div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""weui-btn-area""><em class=""weui-btn weui-btn_primary"" id=""subPost"">提　交</em></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf


	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	'strHtml = strHtml & "	$("".navExtend"").popup();" & vbCrlf
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
	strHtml = strHtml & "	$("".del"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var tID = $(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "		$.confirm(""您确定要删除吗？"", function() {" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "Propose/Delete.html"", {ID:tID}, function(reStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(reStr.reMessge, function(){location.reload(); });" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub SavePost()
	Dim tContent : tContent = Trim(Request("Content"))
	If HR_IsNull(tContent) Then
		ErrMsg = "建议内容不能为空！"
	Else
		Set rs = Server.CreateObject("ADODB.RecordSet")
			rs.Open("Select * From HR_Propose"), Conn, 1, 3
			rs.AddNew
			rs("ID") = GetNewID("HR_Propose", "ID")
			rs("Content") = tContent
			rs("CreateTime") = Now()
			rs("YGDM") = UserYGDM
			rs("IP") = UserTrueIP
			rs.Update
		Set rs = Nothing
		ErrMsg = "您的建议已经成功提交！"
	End If
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub
Sub Delete()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Conn.Execute("Delete From HR_Propose Where ID in(" & tmpID & ")")
	ErrMsg = "删除成功！"
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub
%>