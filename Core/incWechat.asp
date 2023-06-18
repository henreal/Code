<%
'********** 微信接口函数【包括企业微信】 **********
' Powered By：Henreal Studio
' Update：Henreal CIWP V1.0.23 Build 20170208
' Website：http://www.henreal.com
' Weixin：Henreal-Net【恒锐网络科技】
' Tel：0831-8239995 / 13700999995
'----------------------------------------------------

'********** 【注意】提前加载配置文件 **********
' Access Token、Token Expires等均存入XML，解决多平台刷新Token时的不同步问题；
' 缓存均不可采用Cookies、Session；
' 本文件可在各模块中载入；
'----------------------------------------------------
'** 微信Token生命周期为7200秒

'** 读取公众平台及企业微信配置：
Dim qyid : qyid = XmlText("WechatConfig", "qyid", "")					'企业ID【corpid】
Dim qyAgentId : qyAgentId = XmlText("WechatConfig", "qyAgentId", "")	'应用ID【AgentId】
Dim qySecret : qySecret = XmlText("WechatConfig", "qySecret", "")		'应用的凭证密钥【corpsecret】
Dim qyApiUrl : qyApiUrl = "https://qyapi.weixin.qq.com/cgi-bin/"		'企业微信接口
Dim qyBackUrl : qyBackUrl = SiteUrl & "/API/WechatQY.html"				'企业微信回调地址【URL】
Dim qyAccToken : qyAccToken = XmlText("WechatConfig", "qyAccessToken", "")
Dim qyExpires : qyExpires = XmlText("WechatConfig", "qyExpires", "")

Dim boAgentId : boAgentId = XmlText("WechatConfig", "boAgentId", "")	'信息播报应用ID【AgentId】
Dim boSecret : boSecret = XmlText("WechatConfig", "boSecret", "")		'信息播报凭证密钥【corpsecret】
Dim boToken : boToken = XmlText("WechatConfig", "boAccToken", "")		'信息播报Token
Dim boExpires : boExpires = XmlText("WechatConfig", "boExpires", "")	'信息播报Token有效期

Dim wxAppid : wxAppid = XmlText("WechatConfig", "wxAppID", "")			'公众号开发者APPID【AppID】
Dim wxSecret : wxSecret = XmlText("WechatConfig", "wxAppSecret", "")	'公众号开发者密钥【AppSecret】
Dim wxApiUrl : wxApiUrl = "https://api.weixin.qq.com/cgi-bin/"			'公众平台接口【ApiUrl】
Dim wxBackUrl : wxBackUrl = SiteUrl & "/API/Wechat.html"				'公众号回调地址【URL】
Dim wxAccToken : wxAccToken = XmlText("WechatConfig", "wxAccessToken", "")
Dim wxExpires : wxExpires = XmlText("WechatConfig", "wxTokenExpires", "")



'----------------------------------------------------
'********** 微信公众号部分 **********
'=====================================================================
'函数名：ChkWechatToken()	【验证Access Token是否有效】
'返回值：True/False
'=====================================================================
Function ChkWechatToken()
	ChkWechatToken = False
	Dim funExpires : funExpires = Trim(XmlText("WechatConfig", "wxTokenExpires", ""))	'取过期时间
	Dim funToken : funToken = Trim(XmlText("WechatConfig", "wxAccessToken", ""))	'取Access Token
	If Isdate(funExpires) And HR_IsNull(funToken) = False Then
		If DateDiff("s", Now(), funExpires) > 0 Then ChkWechatToken = True		'判断生命周期
	End If
End Function
Function GetWechatToken()
	GetWechatToken = ""
	Dim funOBJ, funToken, funExpires
	Dim funApiUrl : funApiUrl = "https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid=" & wxAppid & "&secret=" & wxSecret & ""		'取access_token
	Dim strGet : strGet = GetHttpStr(funApiUrl, "UTF-8", 1, 10)
	If Instr(strGet, "errcode") > 0 Then		'判断返回数据是否正确
		GetWechatToken = strGet
	Else
		Set funOBJ = parseJSON(strGet)			'返回数据解析JSON
			funToken = funOBJ.access_token		'取Access Token
			funExpires = funOBJ.expires_in		'取Expires
		Set funOBJ = Nothing
		GetWechatToken = funToken
		Call UpdateXmlText("WechatConfig", "wxAccessToken", Trim(funToken))						'缓存Access Token
		Call UpdateXmlText("WechatConfig", "wxTokenExpires", DateAdd("s", funExpires, Now()))	'缓存有效期Expires
	End If
End Function
'=====================================================================
'函数名：PostJsonRemote()	【跨域发送JSON】
'返回值：True/False
'=====================================================================
Function PostJsonRemote(fPostUrl, fPostJson, fPostType)
	Dim strFun, fPostHttp
	Set fPostHttp = CreateObject("MSXML2.XMLHTTP")
		With fPostHttp
			.Open "Post", fPostUrl, False
			.setRequestHeader "Connection","keep-alive"
			.setRequestHeader "Content-Type","application/json; encoding=utf-8"
			.setRequestHeader "Content-Length", len(fPostJson)
			.Send fPostJson
			strFun = .ResponseText
		End With
	Set fPostHttp = Nothing
	PostJsonRemote = strFun
End Function

'********** 企业微信部分 **********
'=====================================================================
'函数名：ChkWechatTokenQY()	【验证Access Token是否有效】
'返回值：True/False
'=====================================================================
Function ChkWechatTokenQY()
	ChkWechatTokenQY = False
	Dim funExpires : funExpires = Trim(XmlText("WechatConfig", "qyExpires", ""))	'取过期时间
	Dim funToken : funToken = Trim(XmlText("WechatConfig", "qyAccessToken", ""))	'取Access Token
	If Isdate(funExpires) And HR_IsNull(funToken) = False Then
		If DateDiff("s", Now(), funExpires) > 0 Then ChkWechatTokenQY = True		'判断生命周期
	End If
End Function

'=====================================================================
'函数名：GetWechatTokenQY()	【取accessToken】
'返回值：String Token
'=====================================================================
Function GetWechatTokenQY()
	GetWechatTokenQY = ""
	Dim funOBJ, funToken, funExpires
	Dim funApiUrl : funApiUrl = "https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=" & qyid & "&corpsecret=" & qySecret & ""		'取access_token
	Dim strGet : strGet = GetHttpStr(funApiUrl, "UTF-8", 1, 10)
	If Instr(strGet, "errcode") > 0 And Instr(strGet, "token") > 0 Then		'判断返回数据是否正确
		Set funOBJ = parseJSON(strGet)			'返回数据解析JSON
			funToken = funOBJ.access_token		'取Access Token
			funExpires = funOBJ.expires_in		'取Expires
		Set funOBJ = Nothing
		GetWechatTokenQY = funToken
		qyAccToken = funToken
		Call UpdateXmlText("WechatConfig", "qyAccessToken", Trim(funToken))		'缓存Access Token
		Call UpdateXmlText("WechatConfig", "qyExpires", DateAdd("s", funExpires, Now()))		'缓存有效期Expires
	End If
End Function

'=====================================================================
'函数名：GetWechatUserInfoQY()	【取会员资料】
'返回值：String
'=====================================================================
Function GetWechatUserInfoQY(fUserID)
	Dim fYGDM, strFun : strFun = "{""errcode"":500,""errmsg"":""Failed to get userinfo [Henreal Network]""}"
	fYGDM = fUserID
	If HR_Clng(fUserID) = 810000 Then fYGDM = "Brett"		'可删除，处理恒锐企业微信超管帐号问题
	If ChkWechatTokenQY() And HR_IsNull(fYGDM) = False Then
		Dim funApiUrl : funApiUrl = "https://qyapi.weixin.qq.com/cgi-bin/user/get?access_token=" & qyAccToken & "&userid=" & Trim(fYGDM)
		Dim funGet : funGet = GetHttpStr(funApiUrl, "UTF-8", 1, 10)
		strFun = funGet
	End If
	GetWechatUserInfoQY = strFun
End Function


'=====================================================================
'函数名：ChkTokenBobao()	【验证信息播报Token是否有效】
'返回值：True/False
'=====================================================================
Function ChkTokenBobao()
	ChkTokenBobao = False
	Dim funExpires : funExpires = Trim(XmlText("WechatConfig", "boExpires", ""))	'取信息播报过期时间
	Dim funToken : funToken = Trim(XmlText("WechatConfig", "boAccToken", ""))		'取信息播报Token
	If Isdate(funExpires) And HR_IsNull(funToken) = False Then
		If DateDiff("s", Now(), funExpires) > 0 Then ChkTokenBobao = True			'判断信息播报生命周期
	End If
End Function

'=====================================================================
'函数名：GetTokenBobao()	【取信息播报Token】
'返回值：String Token
'=====================================================================
Function GetTokenBobao()
	GetTokenBobao = ""
	Dim funOBJ, funToken, funExpires
	Dim funApiUrl : funApiUrl = "https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=" & qyid & "&corpsecret=" & boSecret & ""		'取access_token
	Dim funGet : funGet = GetHttpStr(funApiUrl, "UTF-8", 1, 10)
	If Instr(funGet, "errcode") > 0 And Instr(funGet, "token") > 0 Then		'判断返回数据是否正确
		Set funOBJ = parseJSON(funGet)			'返回数据解析JSON
			funToken = funOBJ.access_token		'取Access Token
			funExpires = funOBJ.expires_in		'取Expires
		Set funOBJ = Nothing
		GetTokenBobao = funToken
		boToken = funToken
		Call UpdateXmlText("WechatConfig", "boAccToken", Trim(funToken))						'缓存信息播报 Token
		Call UpdateXmlText("WechatConfig", "boExpires", DateAdd("s", funExpires, Now()))		'缓存信息播报有效期Expires
	End If
End Function
%>
