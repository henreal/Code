<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="incCommon.asp"-->
<!--#include file="incPurview.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "查看消息"
Dim arrMsgType : arrMsgType = Split(XmlText("Common", "MsgType", ""), "|")

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "UpdateMessage" Call UpdateMessage()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim CountMsg : CountMsg = 0
	sqlTmp = "Select Count(ID) From HR_Message Where isRead=" & HR_False & " And ReceiverID=" & UserYGDM & ""
	Set rsTmp = Conn.Execute(sqlTmp)
		CountMsg = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.ReadNO {color:#f00;}" & vbCrlf		'未读消息颜色
	tmpHtml = tmpHtml & "		.pageBar {box-sizing:border-box;padding-top:8px;}" & vbCrlf		'分页
	tmpHtml = tmpHtml& "		.msgTips {line-height:40px;} .msgTips i {color:#f30;font-size:20px;position: relative;top:3px;} .msgTips b {color:#f00;}" & vbCrlf
	tmpHtml = tmpHtml & "		.ShowCourse, .BackApply, .Transfer {cursor: pointer;color:#000}" & vbCrlf	'消息内容链接
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead("Desktop", 1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	tmpHtml = vbCrlf & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	tmpHtml = "<a href=""" & ParmPath & "Message/Index.html"">我的消息</a><a><cite>列表</cite></a>"
	strHtml = strHtml & getFrameNav(1) : strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-body-w800 hr-shrink-x10"">" & vbCrlf
	'Response.Write "	<div class=""layui-tab layui-tab-brief"" lay-filter=""docDemoTabBrief"">" & vbCrlf
	'Response.Write "		<ul class=""layui-tab-title""><li class=""layui-this"">我收到的消息</li><li>我发送的消息</li></ul>" & vbCrlf
	'Response.Write "		<div class=""layui-tab-content""></div>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf



	Response.Write "	<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	Response.Write "		<legend>我的消息</legend>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "	<div class=""msgTips""><i class=""hr-icon"">&#xee79;</i> "
	If CountMsg > 0 Then Response.Write " 您有 " & CountMsg & " 未读消息！"
	Response.Write " [<b>红色标题</b>为未读消息]</div>" & vbCrlf
	Response.Write "	<div class=""layui-collapse"" lay-filter=""myMessage"">" & vbCrlf

	sqlTmp = "Select * From HR_Message Where ReceiverID=" & UserYGDM & ""
	sqlTmp = sqlTmp & " Order By isRead ASC, SendTime DESC"

	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0 : CurrentPage = 1 : MaxPerPage = tLimit
			If tPage > 0 Then CurrentPage = tPage
			If MaxPerPage <= 0 Then MaxPerPage = 10
			strFileName = ParmPath & "MyCenter/Message.html"

			TotalPut = rsTmp.Recordcount
			If TotalPut > 0 Then
				If CurrentPage < 1 Then CurrentPage = 1
				If (CurrentPage - 1) * MaxPerPage > TotalPut Then
					If (TotalPut Mod MaxPerPage) = 0 Then
						CurrentPage = TotalPut \ MaxPerPage
					Else
						CurrentPage = TotalPut \ MaxPerPage + 1
					End If
				End If
				If CurrentPage > 1 Then
					If (CurrentPage - 1) * MaxPerPage < TotalPut Then
						rsTmp.Move (CurrentPage - 1) * MaxPerPage
					Else
						CurrentPage = 1
					End If
				End If
			End If
			Do While Not rsTmp.EOF
				Response.Write "		<div class=""layui-colla-item"">" & vbCrlf
				Response.Write "			<h2 class=""layui-colla-title hr-rows"
				If HR_CBool(rsTmp("isRead")) = False Then Response.Write " ReadNO"
				Response.Write """ data-id=""" & rsTmp("ID") & """><em>" & rsTmp("Title") & "</em><em>发送时间：" & FormatDate(rsTmp("SendTime"), 1) & "</em></h2>" & vbCrlf
				Response.Write "			<div class=""layui-colla-content"">" & vbCrlf
				Response.Write "				<div class=""hr-fix MsgContent"">" & rsTmp("Message") & "</div>" & vbCrlf
				Response.Write "				<h6 class=""hr-rows""></h6>" & vbCrlf
				Response.Write "			</div>" & vbCrlf
				Response.Write "		</div>" & vbCrlf
				rsTmp.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		Else
			Response.Write "		<div class=""layui-colla-item"">" & vbCrlf
			Response.Write "			<h2 class=""layui-colla-title"">您还没有任何消息</h2>" & vbCrlf
			Response.Write "			<div class=""layui-colla-content"">提示：您当前没有任何消息！</div>" & vbCrlf
			Response.Write "		</div>" & vbCrlf
		End If
	Set rsTmp = Nothing
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-rows pageBar"">" & vbCrlf
	Response.Write "		<div class=""Page_left""></div>" & vbCrlf
	Response.Write "		" & ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "条消息", True) & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "		element.on(""collapse(myMessage)"", function(data){" & vbCrlf
	tmpHtml = tmpHtml & "			if(data.show){" & vbCrlf
	tmpHtml = tmpHtml & "				$.getJSON(""" & ParmPath & "Message/UpdateMessage.html"", {ID:data.title.data(""id"")}, function(result){" & vbCrlf
	tmpHtml = tmpHtml & "					data.title.removeClass(""ReadNO"");" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			};" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot("Desktop", 0) & "" & vbCrlf
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub UpdateMessage()
	Dim tmpJson, tmpID : tmpID = HR_Clng(Request("ID"))
	If tmpID > 0 Then Conn.Execute("Update HR_Message Set isRead=" & HR_True & " Where ID=" & tmpID )
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""消息 " & tmpID & " 设置为已读！"",""ReStr"":""操作完成！""}"
	Response.Write tmpJson
End Sub
%>