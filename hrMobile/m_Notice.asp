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
Dim Page_Title : Page_Title = "通知"

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "Message" Call Message()
	Case "View" Call View()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	SiteTitle = "系统通知"
	strHtml = "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.weui-dialog__bd {text-align:initial;max-height:420px; overflow-y: auto;color:#000;padding:2px;}" & vbCrlf
	strHtml = strHtml & "		.weui-dialog__bd img {max-width:100%;}" & vbCrlf
	strHtml = strHtml & "		.weui-dialog__bd .PubTime {color:#777;padding-top:15px;font-size:14px;}" & vbCrlf
	strHtml = strHtml & "		.iconTit {padding-right:5px;color:#f30;font-size:22px}" & vbCrlf
	strHtml = strHtml & "		#fullView {z-index:5000} .viewbox {box-sizing: border-box;margin:10px;padding:10px;background-color:#fff;}" & vbCrlf
	strHtml = strHtml & "		.viewbox .Content {min-height:12rem;padding:5px 0} .viewbox .Content img {max-width:99%;}" & vbCrlf
	strHtml = strHtml & "		.viewbox .Title {text-align:center;border-bottom:1px solid #ccc;font-size: 1.2rem;color:#18d;padding-bottom:3px;}" & vbCrlf
	strHtml = strHtml & "		.viewbox .PubTime {border-top:1px solid #ccc;font-size: 0.9rem;color:#999;padding-top:3px;}" & vbCrlf
	strHtml = strHtml & "		.popbtn {margin:10px;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Dim tUserID : tUserID = UserID
	Dim newMsgNum : newMsgNum = 0
	If HR_Clng(UserYGDM) > 0 Then tUserID = HR_Clng(UserYGDM)
	Set rsTmp = Conn.Execute("Select count(ID) From HR_Notice Where ID>0")
		newMsgNum = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing

	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf

	sqlTmp = "Select Top 100 * From HR_Notice Where ID>0"
	sqlTmp = sqlTmp & " Order By PublishesTime DESC"
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				Response.Write "	<a class=""weui-cell weui-cell_access"" href=""#"">" & vbCrlf
				'If HR_CBool(rsTmp("isRead")) Then
					Response.Write "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xec17;</i></div><div class=""weui-cell__bd weui-cell_primary viewMSG"" data-id=""" & rsTmp("ID") & """><p>" & rsTmp("Title") & "</p></div>" & vbCrlf
				'Else
				'	Response.Write "		<div class=""weui-cell__bd viewMSG"" data-id=""" & rsTmp("ID") & """><p style=""color:#f30""><i class=""hr-icon"">&#xf003;</i>" & rsTmp("Title") & "</p></div>" & vbCrlf
				'End If
				Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
				Response.Write "	</a>" & vbCrlf
				rsTmp.MoveNext
			Loop
		Else
			Response.Write "	<a class=""weui-cell weui-cell_access"" href=""javascript:;"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>暂时没有通知</p></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
			Response.Write "	</a>" & vbCrlf
		End If
	Set rsTmp = Nothing
	Response.Write "</div>" & vbCrlf
	Response.Write "<div id=""fullView"" class=""weui-popup__container"">" & vbCrlf
	Response.Write "	<div class=""weui-popup__overlay""></div>" & vbCrlf
	Response.Write "	<div class=""weui-popup__modal""><div class=""viewbox""></div><div class=""popbtn""><button class=""weui-btn weui-btn_primary close-popup"">关闭</button></div></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf	
	Response.Write "<div class=""hr-shrink-x20""></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".addNew a"").css(""display"",""none"");" & vbCrlf
	strHtml = strHtml & "	$("".viewMSG"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		$.get(""" & ParmPath & "Notice/View.html"",{id:$(this).data(""id"")}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "			$("".viewbox"").html(rsStr);$(""#fullView"").popup();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub
Sub View()
	Dim tmpMSG, tmpID : tmpID = HR_Clng(Request("id"))
	If tmpID > 0 Then
		Set rsTmp = Conn.Execute("Select * From HR_Notice Where ID=" & tmpID )
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				tmpMSG = "<div class=""Title"">" & Trim(rsTmp("Title")) & "</div>"
				tmpMSG = tmpMSG & "<div class=""Content"">" & HR_HtmlDecode(rsTmp("Content")) & "</div>"
				tmpMSG = tmpMSG & "<div class=""PubTime"">发布时间：" & FormatDate(rsTmp("PublishesTime"), 4) & "</div>"
			End If
		Set rsTmp = Nothing
	End If
	Response.Write tmpMSG
End Sub

%>