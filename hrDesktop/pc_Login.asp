<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incMD5.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="incCommon.asp"-->
<%
Select Case GetUserAgent()
	Case "weixin","wxwork","iPhone","Android"
		Response.Redirect InstallDir & "Touch/Index.html"
End Select

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "管理登陆"

Dim scriptCtrl
Dim tCompany : tCompany = XmlText("Contact", "Company", "")

If IsNull(strParm) Or strParm = "" Then Call LoginBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "ForgetPass" Call ForgetPass()
	Case "ChkLogin" Call ChkLogin()
	Case "GetQRPic" Call GetQRPic()
	Case Else Call LoginBody()
End Select

Sub LoginBody()
	

	Dim strPop, tmpHtml
	If HR_CBool(Request("Logout")) Then		'退出登陆
		Response.Cookies(Site_Sn)("UserPass") = ""
		Response.Cookies(Site_Sn)("YGDM") = ""
		Response.Cookies(Site_Sn)("RndCode") = ""
		ErrMsg = "您已经正常退出！" : strTmp = "点击返回至登陆页"
		ErrHref = "Login.html" : Response.Write GetErrBody(1) : Response.End
		'//strPop = vbCrlf & "	layer.alert(""您已正常退出登陆！"",{icon:6,title:""系统提示""},function(){location.href =""Login.html"";})" & vbCrlf
	ElseIf ChkUserLogin() Then
		ErrMsg = "您已登陆，无须重新登陆！" : strTmp = "点击返回至登陆页"
		ErrHref = "Index.html" : Response.Write GetErrBody(0) : Response.End
		'//strPop = vbCrlf & "	layer.alert(""您已登陆，无须重新登陆！"",{icon:6,title:""系统提示""},function(){location.href =""Index.html"";})" & vbCrlf
	End If
	'Response.Write ChkUserLogin()
	'Response.End
	TempFile = InstallDir & "Static/template/admin/rbLogin.htm"
	strHtml = ReadFromFile(TempFile, "UTF-8", 1)

	strHtml = Replace(strHtml, "[@HeadHtml_RB_1]", getPageHead("Index", 1))		'Header
	strHtml = Replace(strHtml, "[@FootHtml_RB_1]", getPageFoot("Index", 1))		'Footer
	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background:#081726 url(/Static/images/login_bg.jpg) center no-repeat;background-size:100% auto; overflow:hidden;}" & vbCrlf
	tmpHtml = tmpHtml & "		.wxLoginBox {background-color:rgba(255,255,255,0.2);color:#fff;position: absolute;top: 100px;left: 100px;right: 100px;padding: 20px;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .layui-input-inline {width:100%}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginbm dl {display:flex;align-items:center;justify-content:center;height:50px;} .loginbm dt {padding-right:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>"
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")

	tmpHtml = vbCrlf & "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.qrcode.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "<script type=""text/javascript"" src=""http://rescdn.qqmail.com/node/ww/wwopenmng/js/sso/wwLogin-1.0.0.js""></script>" & vbCrlf	'企业微信扫码
	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	if (window != top) { top.location.href = window.location.href; }" & vbCrlf		'防止被框架
	tmpHtml = tmpHtml & "	$(""#TouchQR"").qrcode({ render: ""canvas"", width: 180, height: 180, text: """ & SiteUrl & "/Touch/Index.html"" });" & vbCrlf		'若非H5浏览器将render更换为“table”
	tmpHtml = tmpHtml & "	layui.use([""layer"", ""form""], function () {" & vbCrlf
	tmpHtml = tmpHtml & "		var layer = layui.layer, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		$("".TouchQR"").on(""click"", function () {" & vbCrlf
	tmpHtml = tmpHtml & "			layer.open({ type: 1, title: false, closeBtn: 0, shadeClose: true, skin:""layui-layer-nobg"", content: $(""#TouchQR"") });" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".ForgetPass"").on(""click"", function () { layer.alert(""请联系管理员重置密码！"", {icon:6,title:""系统提示""}); });" & vbCrlf
	tmpHtml = tmpHtml & "		$("".WechatScan"").on(""click"", function () {" & vbCrlf
	tmpHtml = tmpHtml & "			layer.alert(""<div id=\""loginQYQR\"">企业微信二维码</div>"",{btn:""关闭"",offset:'20%',title:""企业微信扫二维码""});" & vbCrlf
	tmpHtml = tmpHtml & "			window.WwLogin({""id"":""loginQYQR"",""appid"":""" & qyid & """,""agentid"":""" & qyAgentId & """," & vbCrlf
	tmpHtml = tmpHtml & "				""redirect_uri"":""" & Server.URLEncode(SiteUrl & "/API/WechatQY/Scan.html") & """,""state"" :""qr_login@henreal""" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".wxLogin"").on(""click"", function () {" & vbCrlf
	tmpHtml = tmpHtml & "			layer.alert(""<div id=\""loginQR\"">微信二维码</div><div class=\""qrTips\""><em>请用微信扫码</em></div>"",{btn:""关闭"",offset:'20%',title:""微信扫二维码""});" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "Login/GetQRPic.html"", {rand:Math.random()}, function (reResult) {" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#loginQR"").html(reResult.errmsg);" & vbCrlf
	tmpHtml = tmpHtml & "				var i=0, chkScan = setInterval(function(){" & vbCrlf
	tmpHtml = tmpHtml & "					ChkUserScan("".qrTips"", i);i++;" & vbCrlf
	tmpHtml = tmpHtml & "				}, 1000);" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		form.on(""submit(Form1)"", function(PostData){" & vbCrlf	'//*提交数据*/
	tmpHtml = tmpHtml & "			console.log('aaabb');" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & strPop
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	function ChkUserScan(select, fTimer){" & vbCrlf
	tmpHtml = tmpHtml & "		$(select).text(fTimer);" & vbCrlf
	tmpHtml = tmpHtml & "		if(fTimer==30){clearInterval(chkScan);}" & vbCrlf
	tmpHtml = tmpHtml & "	}" & vbCrlf
	
	'tmpHtml = tmpHtml & "	$(""input[name=loginPost]"").on(""click"", function () { /*$(""#LoginForm"").submit();*/ });" & vbCrlf
	tmpHtml = tmpHtml & "	$(document).on(""keypress"", function(event){ if(event.keyCode==""13""){Gologin();} });" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#LoginForm"").submit(function(){" & vbCrlf
	tmpHtml = tmpHtml & "		console.log(""dddd"");" & vbCrlf
	tmpHtml = tmpHtml & "		$.getJSON(""" & ParmPath & "Login/ChkLogin.html"", $(""#EditForm"").serialize(),function (data) {" & vbCrlf
	'tmpHtml = tmpHtml & "			if(data.code==0){ /*location.href = """ & ParmPath & "Index.html"";*/ }else{" & vbCrlf
	'tmpHtml = tmpHtml & "				layer.alert(data.msg,{ icon:6,skin:""layer-hr-wr""}, function(index){ layer.close(index); });" & vbCrlf
	'tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>"
	strHtml = Replace(strHtml, "[@FootScript]", "")
	strHtml = Replace(strHtml, "[@ErrMSG]", "模板代码为空")
	strHtml = Replace(strHtml, "[@qyid]", qyid)
	strHtml = Replace(strHtml, "[@qyAgentId]", qyAgentId)

	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub

Sub Logout()
	Dim strSub, rsSub
	UserID = 0
	Response.Cookies(Site_Sn)("UserPass") = ""
	Response.Cookies(Site_Sn)("YGDM") = ""
	Response.Cookies(Site_Sn)("RndCode") = ""
	strSub = "{""code"":500,""msg"":""您已经安全退出登陆！"",""reStr"":""操作成功！"",""rsUrl"":""" & InstallDir & ManageDir & "Login.html""}"
	Response.Write strSub
End Sub
Sub ChkLogin()
	Dim tRndCode : tRndCode = MD5(GetRndString(6), 16)
	Dim rsChk, sqlChk : ErrMsg = ""
	Dim tName : tName = Trim(ReplaceBadChar(Request("Name")))
	Dim tPass : tPass = Trim(Request("Pass"))
	Dim tVcode : tVcode = LCase(ReplaceBadChar(Request("vcode")))
	Dim tTimes : tTimes = HR_Clng(Request.Cookies(Site_Sn)("TIMES"))
	Dim tErrTime : tErrTime = Trim(Request.Cookies(Site_Sn)("err_time"))

	'//取远程验证码：
	Dim tCaptcha, file, str, hf
	file = "/Upload/SESSION/" & Request.cookies("PHPSESSID") & ".txt"
	If FSO.FileExists(Server.MapPath(file)) Then
		Set hf = FSO.OpenTextFile(Server.MapPath(file), 1, False)
			str = hf.ReadAll
			hf.Close
		Set hf = Nothing
	End If
	If str <> tVcode Then ErrMsg = "验证码不正确！"

	If HR_IsNull(tName) Or HR_IsNull(tPass) Then ErrMsg = "登陆帐号或密码不能为空！"
	If HR_Clng(tName) > 999999 Or HR_Clng(tName) < 100000 Then ErrMsg = "您的工号不正确，请重新输入！"
	If HR_IsNull(tVcode) Then ErrMsg = "请输入验证码！"

	Dim tDateDiff : tDateDiff = 0
	If isdate(tErrTime) Then tDateDiff = HR_Clng(DateDiff("s", tErrTime, Now()))	'//获取已经过去的时间

	If tTimes > 4 Then
		If tDateDiff > 900 Then
			Response.Cookies(Site_Sn)("err_time") = ""		'容错时间重置
			Response.Cookies(Site_Sn)("TIMES") = 0
			tTimes = 0
		Else
			ErrMsg = "您已没有试错机会，请" & HR_Clng((900-tDateDiff)/60) & "分钟后再试！"
		End If
	End If
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""err"":false,""errcode"":500,""errmsg"":""" & ErrMsg & """,""icon"":2}" : Exit Sub

	'------ 判断教师工号登陆
	sqlChk = "Select * From HR_Teacher Where TeacherID>0"
	sqlChk = sqlChk & " And YGDM='" & tName & "' And LoginPass='" & MD5(tPass, 16) & "'"
	Set rsChk = Server.CreateObject("ADODB.RecordSet")
		rsChk.Open(sqlChk), Conn, 1, 3
		If Not(rsChk.BOF And rsChk.EOF) Then
			If HR_Clng(rsChk("ApiType")) = 5 Then
				Response.Write "{""err"":false,""errcode"":500,""errmsg"":""" & ErrMsg & """,""icon"":2}" : Exit Sub
			End If
			UserYGDM = rsChk("YGDM")
			UserYGXM = rsChk("YGXM")
			UserRank = 0
			Response.Cookies(Site_Sn)("YGDM") = rsChk("YGDM")
			Response.Cookies(Site_Sn)("UserPass") = rsChk("LoginPass")
			Response.Cookies(Site_Sn)("RndCode") = tRndCode
			Response.Cookies(Site_Sn)("TIMES") = 0
			rsChk("LoginTime") = Now()
			rsChk("LoginIP") = UserTrueIP
			rsChk.Update
			If Trim(rsChk("LoginPass")) = MD5("12345678", 16) Then
				ErrMsg = "成功登陆！<br>请更改初始密码！"
			Else
				ErrMsg = "成功登陆！"
			End If
			Response.Write "{""err"":true,""errcode"":0,""errmsg"":""" & ErrMsg & """,""icon"":1}"
		Else
			tTimes = tTimes+1
			Response.Cookies(Site_Sn)("TIMES") = tTimes
			Response.Cookies(Site_Sn)("err_time") = Now()
			If tTimes>1 Then
				ErrMsg = "帐号或密码不正确！<br>您还有" & 6-tTimes & "次机会！"
			Else
				ErrMsg = "您帐号或密码验证未通过！"
			End If
			Response.Write "{""err"":false,""errcode"":500,""errmsg"":""" & ErrMsg & """,""icon"":2}"
		End If
		rsChk.Close
	Set rsChk = Nothing
End Sub

Sub GetQRPic()
	'取二维码仅执行一次，取Ticket后立即销毁
	If ChkWechatToken() = False Then
		qyAccToken = GetWechatToken()		'重新获得Access Token
	End If

	ErrMsg = ""
	Dim tRndCode : tRndCode = MD5(GetRndString(6), 16)		'随机码，用于判断本次扫码状态

	Dim httpUrl : httpUrl = SiteUrl & "/API/PostJson.htm"	'Post Json接口
	Dim getUrl : getUrl = "https://api.weixin.qq.com/cgi-bin/qrcode/create?access_token=" & qyAccToken
	Dim tScene : tScene = ToUnixTime(Now(), +8)		'场景值，也可是ID
	Dim jsonOBJ, getJson : getJson = "{""expire_seconds"":604800,""action_name"":""QR_STR_SCENE"",""action_info"":{""scene"":{""scene_str"":""web_scan""}}}"
	httpUrl = httpUrl & "?PostAPI=" & Server.URLEncode(getUrl) & "&PostJson=" & getJson		'必须对URL时行编码，然后再接口中解码

	Response.Write wxAccToken & "<br>"
	Response.Write wxExpires & "<br>" & httpUrl & "<br>"
	Exit Sub

	Dim tTicket, reStr, qrPicUrl, qrUrl
	reStr = GetHttpStr(httpUrl, "UTF-8", 1, 10)
	If Instr(reStr,"ticket")>0 Then
		Set jsonOBJ = parseJSON(reStr)
			tTicket = jsonOBJ.ticket
			Response.Cookies(Site_Sn)("wxTicket") = Trim(tTicket)
			qrPicUrl = "https://mp.weixin.qq.com/cgi-bin/showqrcode?ticket=" & Server.URLEncode(tTicket)
			qrPicUrl = Replace(qrPicUrl, "/", "\/")
			qrUrl = jsonOBJ.url : qrUrl = Replace(qrUrl, "/", "\/")
			Response.Write "{""errcode"":0,""errmsg"":""<img src=\""" & qrPicUrl & "\"">"",""scene_str"":""" & tScene & """,""ticket"":""" & tTicket & """,""expire"":""" & jsonOBJ.expire_seconds & """,""qrurl"":""" & qrUrl & """,""RndCode"":""" & tRndCode & """}"
			Set rsTmp = Server.CreateObject("ADODB.RecordSet")		'保存到临时表
				rsTmp.Open("Select * From HR_wxScanTmp Where Scene='" & tScene & "'"), Conn, 1, 3
				If rsTmp.BOF And rsTmp.EOF Then
					rsTmp.AddNew
					rsTmp("ID") = GetNewID("HR_wxScanTmp", "ID")
					rsTmp("QRCodeID") = tRndCode
					rsTmp("Ticket") = tTicket
					rsTmp("Scene") = tScene
					rsTmp("QRUrl") = Trim(jsonOBJ.url)
					rsTmp("ExpireTime") = DateAdd("s", jsonOBJ.expire_seconds, Now())	'到期时间
					rsTmp("wxOpenID") = ""
					rsTmp.Update
				End If
				rsTmp.Close
			Set rsTmp = Nothing
		Set jsonOBJ = Nothing
	Else
		ErrMsg = "微信二维码获取失败[Ticket Err]"
		'Response.Write "{""errcode"":500,""errmsg"":""" & ErrMsg & """}"
		Response.Write reStr
	End If
End Sub


%>
