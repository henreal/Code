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
SiteTitle = "通知管理"

Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "SavePost" Call SavePost()
	Case "Edit" Call EditBody()
	Case "Delete" Call Delete()
	Case Else Response.Write GetErrBody(0)
End Select


Sub MainBody()
	Dim CountNotice		'汇总通知
	Set rsTmp = Conn.Execute("Select count(ID) From HR_Notice Where ID>0")
		CountNotice = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing

	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells {margin:0;}" & vbCrlf

	strHtml = strHtml & "		.hr-float-btn {width:55px;height:55px;text-align:center;font-size:2.6rem;position: fixed;right: 1rem;bottom: 3rem;z-index: 9;}" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn i {color:#4CAF50;}" & vbCrlf
	strHtml = strHtml & "		.navExtend {height: initial;flex-grow:2;text-align:right;}" & vbCrlf
	strHtml = strHtml & "		.navExtend span {font-size:1.2rem;display:line-block;background-color:#f7ce93;padding:2px 3px;color:#035;border-radius: 2px}" & vbCrlf
	strHtml = strHtml & "		.viewMSG .tips {color:#999;font-size:0.7rem;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-form-preview"">" & vbCrlf
	Response.Write "	<div class=""weui-form-preview__hd"">" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item""><label class=""weui-form-preview__label"">全部通知</label><em class=""weui-form-preview__value"">" & CountNotice & " 条</em></div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf

	sqlTmp = "Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.Sender) As Sender From HR_Notice a Where a.ID>0"
	sqlTmp = sqlTmp & " Order By a.PublishesTime DESC"
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageNotice/Edit.html?ID=" & rsTmp("ID") & """>" & vbCrlf
				Response.Write "		<div class=""weui-cell__bd viewMSG"" data-id=""" & rsTmp("ID") & """><p><i class=""hr-icon"">&#xef7b;</i>" & rsTmp("Title") & "</p><p class=""tips"">" & FormatDate(rsTmp("PublishesTime"), 1) & "　发布：" & rsTmp("Sender") & "</p></div>" & vbCrlf
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

	Response.Write "<div id=""half"" class=""weui-popup__container popup-bottom"">" & vbCrlf
	Response.Write "	<div class=""weui-popup__overlay""></div>" & vbCrlf
	Response.Write "	<div class=""weui-popup__modal"">" & vbCrlf
	Response.Write "		<div class=""toolbar""><div class=""toolbar-inner""><span class=""picker-button close-popup"">关闭</span><h1 class=""title"">发布通知</h1></div></div>" & vbCrlf
	Response.Write "		<div class=""modal-content"">" & vbCrlf
	Response.Write "			<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "				<div class=""weui-cell"">" & vbCrlf
	Response.Write "					<div class=""weui-cell__bd""><textarea class=""weui-textarea title"" name=""title"" id=""title"" placeholder=""请输入通知标题"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "				</div>" & vbCrlf
	Response.Write "				<div class=""weui-cell"">" & vbCrlf
	Response.Write "					<div class=""weui-cell__bd""><textarea class=""weui-textarea content"" name=""content"" placeholder=""请输入通知内容"" rows=""5""></textarea></div>" & vbCrlf
	Response.Write "				</div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""weui-btn-area""><em class=""weui-btn weui-btn_primary"" id=""subPost"">提　交</em></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-float-btn"">" & vbCrlf
	Response.Write "	<span class=""navExtend open-popup"" data-target=""#half""><i class=""hr-icon"">&#xf3c0;</i></span>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Manage/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var title = $(""#title"").val(), content = $("".content"").val();" & vbCrlf
	strHtml = strHtml & "		if(title==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""通知标题不能为空"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else if(content==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""通知内容太少"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else{" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "ManageNotice/SavePost.html"", {Title:title,Content:content}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(rsStr.reMessge, function(){ $.closePopup();location.reload(); });" & vbCrlf
	strHtml = strHtml & "			},""json"");" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

Sub SavePost()
	Dim tTitle : tTitle = Trim(Request("Title"))
	Dim tContent : tContent = Trim(Request("Content"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	If HR_IsNull(tTitle) Then
		ErrMsg = "通知标题没有填写！"
	ElseIf HR_IsNull(tContent) Then
		ErrMsg = "通知内容不能为空！"
	Else
		sql = "Select * From HR_Notice Where ID=" & tmpID
		Set rs = Server.CreateObject("ADODB.RecordSet")
			rs.Open sql, Conn, 1, 3
			If rs.BOF And rs.EOF Then
				rs.AddNew
				rs("ID") = GetNewID("HR_Notice", "ID")
				rs("PublishesTime") = Now()
				rs("Hits") = 0
			End If
			rs("Title") = tTitle
			rs("Content") = HR_HTMLEncode(tContent)
			rs("KeyWord") = ""
			rs("Sender") = UserYGDM
			rs.Update
		Set rs = Nothing
		ErrMsg = "通知已经修改成功！"

		'发送消息到企业微信提醒！【所有人】
		tContent = HR_HtmlDecode(Trim(tContent)) : tContent = Replace(nohtml(tContent), " ", "") : tContent = Replace(nohtml(tContent), "&nbsp;", "") : tContent = GetSubStr(tContent, 110, True)
		tContent = "发送时间：" & FormatDate(Now(), 1) & "<br>" & tContent
		tContent = Replace(tContent, "</p><p>", "<br>")
		Call SentWechatMSG_QYCard("@all", tTitle, SiteUrl & "/Touch/Notice/Index.html?ID=2", tContent)
	End If
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub
Sub Delete()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Conn.Execute("Delete From HR_Notice Where ID in(" & tmpID & ")")
	ErrMsg = "通知删除成功！"
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub

Sub EditBody()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tTitle, tContent, tSender, tSenderID, ChrNum : ChrNum = 0
	Set rs = Conn.Execute("Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.Sender) As YGXM From HR_Notice a Where a.ID=" & tmpID)
		If Not(rs.BOF And rs.EOF) Then
			tTitle = Trim(rs("Title"))
			tContent = Trim(rs("Content"))
			If HR_IsNull(tContent) = False Then ChrNum = Len(tContent)
			tSenderID = HR_Clng(rs("Sender"))
			tSender = Trim(rs("YGXM"))
		End If
	Set rs = Nothing
	SiteTitle = "修改通知"
	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	'strHtml = strHtml & "		.editbtn em {width:50%;padding:10px 5px;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<div class=""weui-cells__title"">通知标题</div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input name=""Title"" id=""Title"" value=""" & tTitle & """ class=""weui-input"" type=""text"" placeholder=""请输入通知标题""></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""weui-cells__title"">通知内容</div>" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Content"" id=""Content"" placeholder=""请输入文本"" rows=""10"">" & tContent & "</textarea><div class=""weui-textarea-counter""><span id=""charnum"">" & ChrNum & "</span>/300</div></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""weui-cells__title"">发布者：" & tSender & "</div>" & vbCrlf
	
	Response.Write "<div class=""hr-rows hr-editbtn"">" & vbCrlf
	Response.Write "	<em><i class=""hr-icon"">&#xea3f;</i></em>" & vbCrlf
	If tSenderID = 0 Or tSenderID=UserYGDM Or UserRank > 1 Then
		Response.Write "	<em><button type=""button"" name=""save"" class=""save"" id=""SaveEdit"">保存</button></em>" & vbCrlf
		Response.Write "	<em><button type=""button"" name=""delete"" class=""delete"" id=""Delete"" data-id=""" & tmpID & """>删除</button></em>" & vbCrlf
	Else
		Response.Write "	<em class=""hr-disabled""><button type=""button"" name=""save"">保存</button></em>" & vbCrlf
		Response.Write "	<em class=""hr-disabled""><button type=""button"" name=""delete"">删除</button></em>" & vbCrlf
	End If
	Response.Write "</div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "ManageNotice/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$(""#Content"").bind(""input propertychange"",function(){" & vbCrlf		'统计输入字符数
	strHtml = strHtml & "		$(""#charnum"").text($(this).val().length);" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#SaveEdit"").click(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var title = $(""#Title"").val(), content = $(""#Content"").val();" & vbCrlf
	strHtml = strHtml & "		if(title==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""通知标题不能为空"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else if(content==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""通知内容太少"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else{" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "ManageNotice/SavePost.html"", {Title:title,Content:content,ID:" & tmpID & "}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(rsStr.reMessge, function(){ $.closePopup();location.reload(); });" & vbCrlf
	strHtml = strHtml & "			},""json"");" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#Delete"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var tID = $(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "		$.confirm(""您确定要删除吗？"", function() {" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "ManageNotice/Delete.html"", {ID:tID}, function(reStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(reStr.reMessge, function(){location.href=""" & ParmPath & "ManageNotice/Index.html""; });" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

%>