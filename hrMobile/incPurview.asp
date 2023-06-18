<%
'======== 触屏版公用部分【权限等】 ========

If Not(ChkUserLogin()) Then			'由本系统登陆身份识别
	Dim chkOrigin : chkOrigin = GetUserAgent()
	Dim wxBackUri : wxBackUri = SiteUrl & "/API/WechatQY/Login.html"								'企业微信回调URL
	Dim wxworkUrl : wxworkUrl = "https://open.weixin.qq.com/connect/oauth2/authorize?appid=" & qyid & "&redirect_uri=" & Server.URLEncode(wxBackUri) & "&response_type=code&scope=snsapi_base&agentid=" & qyAgentId & "&state=wx_login@henreal#wechat_redirect"

	If chkOrigin = "wxwork" Or chkOrigin = "weixin" Then			'判断来路是否为企业微信或微信
		Response.Redirect wxworkUrl									'微信身份认证
		Response.End
	End If
	Response.Redirect InstallDir & "Touch/Login.html"
	Response.End
End If
Set rs = Conn.Execute("Select Top 1 * From HR_User Where YGDM=" & UserYGDM & " Order By UserID ASC")
	If Not(rs.BOF And rs.EOF) Then
		UserID = HR_Clng(rs("UserID"))
		UserRank = HR_Clng(rs("ManageRank"))
		HeadFace = Trim(rs("UserFace"))
	End If
Set rs = Nothing

%>