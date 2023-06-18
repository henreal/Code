<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incMD5.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl

Dim wxBackUri : wxBackUri = SiteUrl & "/API/WechatQY/Scan.html"					'企业微信扫码回调URL
Dim wxScanUrl : wxScanUrl = "https://open.work.weixin.qq.com/wwopen/sso/qrConnect?appid=" & qyid & "&agentid=" & qyAgentId & "&redirect_uri=" & Server.URLEncode(wxBackUri) & "&state=qr_login@henreal"	'扫码登陆
wxBackUri = SiteUrl & "/API/WechatQY/Login.html"								'企业微信回调URL
Dim qyAuthUrl : qyAuthUrl = "https://open.weixin.qq.com/connect/oauth2/authorize?appid=" & qyid & "&redirect_uri=" & Server.URLEncode(wxBackUri) & "&response_type=code&scope=snsapi_base&agentid=" & qyAgentId & "&state=wx_login@henreal#wechat_redirect"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Login" Call qyLogin()
	Case "Scan" Call ScanBack()			'扫码登陆回调地址
	Case "UserInfo" Call UserInfo()
	Case Else Call MainBody()
End Select

Sub MainBody()
	Response.Redirect qyAuthUrl
	'Response.Write "请点击企业微信登陆"
	'Response.Write "<br><a href=""" & qyAuthUrl & """>企业微信登陆</a>"
End Sub

Sub qyLogin()		'企业微信Web OAuth2.0接入【只能在企业微信APP中打开】
	Dim reCode : reCode = Trim(Request("code"))			'回调code
	Dim reAppid : reAppid = Trim(Request("appid"))		'回调appid
	Dim reState : reState = Trim(Request("state"))		'回调state
	Dim strJson, jsonOBJ, tmpUrl : tmpUrl = InstallDir & "Touch/Index.html?origin=wxqy"
	Dim tRndCode : tRndCode = MD5(GetRndString(6), 16)
	
	If ChkUserLogin() Then		'已登陆，进入首页
		Response.Redirect tmpUrl
		Exit Sub
	End If

	If HR_IsNull(reCode) And reState <> "wx_login@henreal" Then	'非企业微信登陆接口则重新登陆
		Response.Write ShowTipsPage("请用企业微信登陆")
		Exit Sub
	End If

	If Not(ChkWechatTokenQY) Then		'Access Token失效时
		qyApiUrl = qyApiUrl & "gettoken?corpid=" & qyid & "&corpsecret=" & qySecret & ""		'取access_token
		Call GetWechatTokenQY()			'重新获取Access Token
		Response.Redirect "?origin=wxqy&code=" & reCode & "&state=" & reState & "&appid=" & reAppid		'重刷当前面
		Exit Sub
	End If

	'通过code获取用户信息
	Dim rsAdd, tmpYGDM, ChkUserQY : ChkUserQY = False
	qyApiUrl = qyApiUrl & "user/getuserinfo?access_token=" & qyAccToken & "&code=" & reCode
	strJson = GetHttpStr(qyApiUrl, "UTF-8", 2, 10)	'从API获取会员UserId	
	If Instr(strJson, "UserId") > 0 Then			'判断是否为本企业会员
		ChkUserQY = True
	End If

	If Instr(strJson, "UserId") > 0 Then
		Set jsonOBJ = parseJSON(strJson)
			tmpYGDM = jsonOBJ.UserId
			If tmpYGDM = "Brett" Then tmpYGDM = "810000"	'正常运行时删除，为解决恒锐企业微信超管帐号为别名
			Set rsTmp = Server.CreateObject("ADODB.RecordSet")
				rsTmp.Open("Select Top 1 * From HR_Teacher Where YGDM='"& tmpYGDM &"' Order By TeacherID ASC"), Conn, 1, 3
				If Not(rsTmp.BOF And rsTmp.EOF) Then		'已经与系统绑定
					Response.Cookies(Site_Sn)("YGDM") = tmpYGDM
					Response.Cookies(Site_Sn)("UserPass") = rsTmp("LoginPass")
					Response.Cookies(Site_Sn)("RndCode") = tRndCode
					rsTmp("LoginIP") = UserTrueIP
					rsTmp("LoginTime") = Now()
					rsTmp.Update
					Response.Redirect tmpUrl	'登陆成功进入系统首页
				Else
					Session("wxqy_UserID") = jsonOBJ.UserId			'将会员信息存入缓存（用于绑定）
					Session("wxqy_DeviceId") = jsonOBJ.DeviceId
					Response.Write ShowTipsPage("您的帐号暂未绑定，请用帐号登陆")
				End If
			Set rsTmp = Nothing
		Set jsonOBJ = Nothing
	Else
		'取微信用户资料失败，进入帐号登陆
		Response.Write ShowTipsPage("企业微信接口返回数据出错！<br>请关注企业号后再试。")
	End If
	
End Sub

Sub ScanBack()		
	Dim reCode : reCode = Trim(Request("code"))			'扫码回调code
	Dim reAppid : reAppid = Trim(Request("appid"))		'扫码回调appid
	Dim reState : reState = Trim(Request("state"))		'扫码回调state，用于判断是否由自己发出
	Dim strJson, tToken, jsonOBJ, reUserID
	If Not(ChkWechatTokenQY) Then		'Access Token失效时
		qyApiUrl = "https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=" & qyid & "&corpsecret=" & qySecret & ""		'取access_token
		Call GetWechatTokenQY()		'重新获取Access Token
		Response.Redirect "?code=" & reCode & "&state=" & reState & "&appid=" & reAppid		'重刷当前面
		Exit Sub
	End If
	If HR_IsNull(reCode) = False And reState = "qr_login@henreal" Then		'通过code获取用户信息
		qyApiUrl = "https://qyapi.weixin.qq.com/cgi-bin/user/getuserinfo?access_token=" & qyAccToken & "&code=" & reCode
		strJson = GetHttpStr(qyApiUrl, "UTF-8", 2, 10)
		If Instr(strJson, "errcode") > 0 Then
			Set jsonOBJ = parseJSON(strJson)
				If jsonOBJ.errcode > 0 Then
					Response.Write "错误提示：" & jsonOBJ.errmsg
					Response.Write "<br><a href=""" & InstallDir & "Desktop/Login.html"">返回登陆</a>"
				Else		'通过工号登陆
					UserYGDM = Trim(jsonOBJ.UserId)
					If UserYGDM = "Brett" Then UserYGDM = "810000"
					Set rsTmp = Server.CreateObject("ADODB.RecordSet")
						rsTmp.Open("Select Top 1 * From HR_Teacher Where YGDM='" & UserYGDM & "' Order By TeacherID ASC"), Conn, 1, 3
						If Not(rsTmp.BOF And rsTmp.EOF) Then		'已经与系统绑定
							If wxApiUserLogin(UserYGDM, 0) Then
								Response.Redirect InstallDir & "Desktop/Index.html"		'进入系统管理
								Exit Sub
							End If
						Else
							Response.Write ShowTipsPage("工号“" & UserYGDM & "”不存在！")
						End If
					Set rsTmp = Nothing
				End If
			Set jsonOBJ = Nothing
		End If
	Else
		Response.Write ErrMsg
		Response.Write "<br><a href=""" & InstallDir & "Desktop/Login.html"">返回登陆</a>"
	End If
End Sub

Sub UserInfo()
	Dim tYGDM : tYGDM = HR_Clng(Request("YGDM"))
	If tYGDM = 0 And ChkUserLogin() Then tYGDM = UserYGDM
	Response.Write GetWechatUserInfoQY(tYGDM)
End Sub

Function ShowTipsPage(strTips)
	Dim strFun
	strFun = "<!DOCTYPE html>" & vbCrlf
	strFun = strFun & "<html lang=""zh-CN"">" & vbCrlf
	strFun = strFun & "<head>" & vbCrlf
	strFun = strFun & "	<meta charset=""utf-8"">" & vbCrlf
	strFun = strFun & "	<title>错误提示_" & SiteName & "</title>" & vbCrlf
	strFun = strFun & "	<meta name=""viewport"" content=""width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no"">" & vbCrlf		'应用Touch，禁止缩放
	'strFun = strFun & "	<meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & vbCrlf
	strFun = strFun & "	<link rel=""stylesheet"" href=""https://cdn.jsdelivr.net/npm/bootstrap@3.3.7/dist/css/bootstrap.min.css"" integrity=""sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u"" crossorigin=""anonymous"">" & vbCrlf
	strFun = strFun & "	<style type=""text/css"">" & vbCrlf
	strFun = strFun & "		body{padding:20px;} .tips-box{text-align:center;} .tips-icon span {font-size: 64px;color:#f30}" & vbCrlf
	strFun = strFun & "		.tips-desc h3 {font-size:2rem;} .tips-btn {padding-top:30px;}" & vbCrlf
	strFun = strFun & "	</style>" & vbCrlf
	strFun = strFun & "</head>" & vbCrlf
	strFun = strFun & "<body>" & vbCrlf
	strFun = strFun & "<div class=""tips-box"">" & vbCrlf
	strFun = strFun & "	<div class=""tips-icon""><span class=""glyphicon glyphicon-minus-sign"" aria-hidden=""true""></span></div>" & vbCrlf
	strFun = strFun & "	<div class=""tips-desc""><h3>" & strTips & "</h3></div>" & vbCrlf
	Select Case GetUserAgent()
		Case "wxwork","weixin","iPhone","Android" strFun = strFun & "	<div class=""tips-btn""><a class=""btn btn-info"" href=""" & InstallDir & "Touch/Login.html?noBind=1"" role=""button"">返回登陆</a></div>" & vbCrlf
		Case Else strFun = strFun & "	<div class=""tips-btn""><a class=""btn btn-info"" href=""" & InstallDir & "Desktop/Login.html?noBind=1"" role=""button"">返回登陆</a></div>" & vbCrlf
	End Select
	strFun = strFun & "</div>" & vbCrlf
	strFun = strFun & "</body>" & vbCrlf
	strFun = strFun & "</html>" & vbCrlf
	ShowTipsPage = strFun
End Function

Function wxApiUserLogin(fYGDM, fIsAdd)		'仅支持企业微信登陆，fIsAdd：0仅登陆，1添加新会员
	Dim strFun : strFun = False
	Dim rsFun, sqlFun, fRndCode : fRndCode = MD5(GetRndString(6), 16)
	Dim funApiUrl, funJson, funOBJ

	If HR_Clng(fYGDM) > 0 Then		'工号不能为空
		sqlFun = "Select Top 1 * From HR_Teacher Where YGDM='"& fYGDM &"' Order By TeacherID ASC"
		Set rsFun = Server.CreateObject("ADODB.RecordSet")
			rsFun.Open sqlFun, Conn, 1, 3
			If Not(rsFun.BOF And rsFun.EOF) Then		'已经与系统绑定
				Response.Cookies(Site_Sn)("YGDM") = fYGDM
				Response.Cookies(Site_Sn)("UserPass") = rsFun("LoginPass")
				Response.Cookies(Site_Sn)("RndCode") = fRndCode
				rsFun("LoginIP") = UserTrueIP
				rsFun("LoginTime") = Now()
				UserYGXM = rsFun("YGXM")
				rsFun.Update
				strFun = True
				ErrMsg = UserYGXM & "["& fYGDM & "]成功登陆！"
			Else
				ErrMsg = "未添加微信会员“" & fYGDM & "”至系统中！"
			End If
		Set rsFun = Nothing
	Else
		ErrMsg = "没有工号！"
	End If
	wxApiUserLogin = strFun
End Function
%>