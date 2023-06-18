<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim scriptCtrl : SiteTitle = "发送消息"

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "AddNew" Call SendFrom()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.sumbar b {color:#F60;padding:0 2px} .sumbar b.sumDebit{color:#080}" & vbCrlf		'表头汇总
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	tmpHtml = "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)

	tmpHtml = "<a href=""" & ParmPath & "Message/Index.html"">" & SiteTitle & "</a><a><cite>发送</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf

	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = Replace(getPageFoot(1), "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

End Sub

Sub SendFrom()
	
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .morebtn {padding:3px 0!important;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .tips {padding-left:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	tmpHtml = "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)

	tmpHtml = "<a href=""" & ParmPath & "Message/Index.html"">" & SiteTitle & "</a><a><cite>编辑新消息</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form class=""layui-form"" id=""EditForm"" name=""EditForm"" lay-filter=""EditForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layui-form-item""><label class=""layui-form-label"">选择接收人：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""text"" name=""YGXM"" id=""ygxm"" value="""" lay-verify=""required"" autocomplete=""on"" title=""查找评价人"" class=""layui-input txt1"" readonly></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux morebtn""><span class=""layui-btn layui-btn-sm getBtn"" data-code=""ygdm"" data-name=""ygxm"">查找</span><span class=""tips"">请输入关键字搜索教师，必填项</span></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">工　　号：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""text"" name=""YGDM"" id=""ygdm"" lay-verify=""required"" value="""" class=""layui-input txt1"" readonly></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item""><label class=""layui-form-label"">内　容：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><textarea name=""Content"" id=""content"" style=""width:100%;height:180px;""></textarea></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<input type=""hidden"" name=""ID"" value=""""><input type=""hidden"" name=""Modify"" value=""True"">"
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><button class=""layui-btn"" lay-submit lay-filter=""SubPost"">发送</button><button type=""reset"" class=""layui-btn layui-btn-primary"">重置</button></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf

	Response.Write "	</form>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	If ChkTokenBobao() = False Then Call GetTokenBobao()

	Dim postUrl : postUrl = "https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=" & boToken
	Response.Write "<br>" & boExpires

	Dim tMsg : tMsg = "{""touser"": ""Brett|207062|798017"",""msgtype"":""text"",""agentid"":" & boAgentId & ",""text"" :{""content"": ""文本消息测试消息，有链接：<a href=\""http://www.thinray.net"">打开消息</a><br>发送时间：" & Formatdate(Now(), 10) & """,""safe"":0}"
	Response.Write PostJsonRemote(postUrl, tMsg, 0)
	tMsg = "{""touser"": ""Brett|207062|798017"",""msgtype"":""news"",""agentid"":" & boAgentId & ",""news"" :{""articles"":["
	tMsg = tMsg & "{""title"":""会议通知消息提示"",""description"":""测试消息，有链接：请点击详情查看。<br/>发送时间：" & Formatdate(Now(), 10) & """,""url"":""https://www.thinray.net"",""picurl"":""https://www.thinray.net/Upload/test1.jpg""}"
	tMsg = tMsg & ",{""title"":""第二条消息提醒"",""description"":""第二条消息提醒测试消息，有链接：请点击详情查看。<br/>发送时间：" & Formatdate(Now(), 10) & """,""url"":""https://www.thinray.net"",""picurl"":""https://www.thinray.net/Upload/test1.jpg""}"
	tMsg = tMsg & "]}"
	Response.Write PostJsonRemote(postUrl, tMsg, 0)


	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element;" & vbCrlf

	tmpHtml = tmpHtml & "		$("".getBtn"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var elcode=$(this).data(""code""), elname=$(this).data(""name"");" & vbCrlf		'返回员工代码及名称时的对象
	tmpHtml = tmpHtml & "			var openurl=""" & InstallDir & "Desktop/Contacts/Float.html?Type=3"";" & vbCrlf
	tmpHtml = tmpHtml & "			parent.layer.open({type:2,id:""getWin"",content:openurl, title:[""选择教师"",""font-size:16""],area:[""500px"", ""80%""],scrollbar:false,success:function(layero, index){" & vbCrlf
	tmpHtml = tmpHtml & "				var objIframe = $(layero).find('iframe')[0].contentWindow.document.body;" & vbCrlf
	tmpHtml = tmpHtml & "				var obj1 = $(objIframe).contents().find(""#listGroup"");" & vbCrlf
	tmpHtml = tmpHtml & "				obj1.attr(""value"",window.name);obj1.attr(""code"", elcode); obj1.attr(""name"", elname);" & vbCrlf		'回车搜索
	tmpHtml = tmpHtml & "				}" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = Replace(getPageFoot(1), "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub




Function SendAppMessage()		'发送消息应用间
	
End Function
%>