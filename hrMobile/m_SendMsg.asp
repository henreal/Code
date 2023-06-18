<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<%
Dim scriptCtrl
Response.Write "到期：" & qyExpires & "<br>"

If ChkTokenBobao() = False Then
	Response.Write "重置Token：" & GetTokenBobao() & "<br>"
	Response.End
End If

Dim tYGDM, tTitle, tURL, tContent

tYGDM = "810000"
tTitle = "测试，由信息播报发送，调换课申请问题已处理"
tURL = "http://jx.wzhealth.com/Touch/Swap/Index.html"
tContent = "您提出的调换课问题技术人员已经解决，您收到此条消息时，请微信回复一下。<br>时间：" & FormatDate(Now(), 10)
Response.Write SentWechatMSG_QYCard(tYGDM, tTitle, tURL, tContent)


'======== 返回无管理权限提示 ========
Function GetManagePermit()
	Dim strFun : strFun = ""
	If UserRank=0 Then
		strFun = "<div class=""hr-permit-tips""><dl class=""hr-rows tipsBox""><dt><i class=""hr-icon"">&#xf05e;</i></dt><dd>您没有访问权限！<br><b>Access to the page is denied!</b></dd></dl>"
		strFun = strFun & "<h4 class=""back-btn""><a href=""" & ParmPath & "Index.html"" class=""hr-btn"">返回首页</a></h4></div>"
	End If
	GetManagePermit = strFun
End Function

'======== 发送文本卡片消息 ========
Function SentWechatMSG_QYCard1(sTouser, sTitle, sURL, sContent)		'发送文本卡片消息
	Dim postJson, strSub : strSub = ""
	If HR_IsNull(sTouser) = False And HR_IsNull(sTitle) = False And HR_IsNull(sURL) = False And HR_IsNull(sContent) = False Then
		sContent = Replace(sContent, """", "\""")
		postJson = "{""touser"":""" & sTouser & """,""msgtype"":""textcard"",""agentid"":" & qyAgentId & ",""textcard"":{""title"":""" & sTitle & """,""description"":""" & sContent & """,""url"":""" & sURL & """,""btntxt"":""查看详情""}}"
		strSub = PostWechatMessageQY1(postJson, 1)
	End If
	SentWechatMSG_QYCard1 = strSub
End Function
'=====================================================================
'函数名：PostWechatMessageQY()	【企业微信发送会员消息】
'返回值：String
'=====================================================================
Function PostWechatMessageQY1(fPostJson, fPostType)
	Dim strFun, funErr, fPostHttp, fPostUrl
	If Not(ChkWechatTokenQY) Then		'判断Access Token
		PostWechatMessageQY = "{""errcode"":500, ""errmsg"":""Access Token 已过期"", ""invaliduser"":""""}"
		Exit Function
	End If
	fPostUrl = "https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=" & qyAccToken
	Set fPostHttp = CreateObject("Msxml2.ServerXMLHTTP")
		With fPostHttp
			.Open "Post", fPostUrl, False
			.setRequestHeader "Content-Type","application/xml;charset=UTF-8"
			.Send fPostJson

			strFun = .ResponseText
			funErr = .status
			Response.Write "<br>readyState：" & .readyState
			Response.Write "<br>statusText：" & .statusText
		End With
	Set fPostHttp = Nothing
	PostWechatMessageQY1 = strFun
	Response.Write "<br>status：" & funErr
	
End Function
%>