<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incMD5.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<%
If GetUserAgent() = "wxwork" And HR_CLng(Request("noBind")) = 0 Then
	Response.Redirect InstallDir & "API/WechatQY.html"
	Response.End
End If

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "管理登陆"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index","Logout" Call MainBody()
	Case "ChkLogin" Call ChkLogin()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	tmpHtml = "<link type=""text/css"" href=""[@Web_Dir]Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	tmpHtml = tmpHtml & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background:#081726 url(/Static/images/login_bg_m.jpg) center bottom no-repeat;background-size:100% 100%; overflow:hidden;}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginNav {background-color:rgba(0,0,0,0.2);height:47px;line-height:47px;width:100%;color:#ccc;border-bottom:1px solid rgba(108, 224, 255, 0.3)}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginNav a {padding:0 5px;color:#ccc;font-size:1rem}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginNav a:hover {color:rgba(108, 224, 255, 0.87);background-color:rgba(255,255,255,0.1)}" & vbCrlf
	tmpHtml = tmpHtml & "		.logo i.hr-icon {padding-right:15px;color:rgba(108, 224, 255, 0.87);font-size:0.8rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginNav .logoPic {display:block;height:47px;width:50px;background:url(/Static/images/uLogo1.png) center no-repeat;background-size:auto 30px}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginNav li{font-size:1rem} .loginNav li b {padding:0 5px;color:#777;font-family:Georgia}" & vbCrlf
	tmpHtml = tmpHtml & "		" & vbCrlf
	tmpHtml = tmpHtml & "		.loginBox {width:90%;margin:0 auto;border:0px solid #ccc;background-color:rgba(108, 224, 255, 0)}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginTips {color:#fff;} .loginTips h1 {text-align:center;padding-top:20px;} .loginTips h1 img {width:60px}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginTips h2 {font-size:1.2rem;text-align:center;padding-bottom: 20px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginform {width:90%;margin:0 auto;background-color:rgba(255,255,255,1);border-radius:5px;padding:15px 0;box-sizing:border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginform li {padding:8px 25px;} .loginTit {font-size:1.1rem;line-height:180%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginTit span {font-size:0.9rem;font-family:Georgia;color:#aaa;padding:0 0 0 15px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .layui-input-inline {width:100%}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginInput {width:95%;height:45px;line-height:45px;padding:0 10px;border-radius:5px ;border:1px solid #098;font-size:0.9rem}" & vbCrlf
	tmpHtml = tmpHtml & "		.LoginFooter {width:100%;height:50px;line-height:50px;overflow:hidden;position:initial;box-sizing:border-box;;background-color:rgba(0,0,0,0.7);margin-top:100px}" & vbCrlf
	tmpHtml = tmpHtml & "		.loginbm {text-align:center;color:#ccc;}" & vbCrlf
	tmpHtml = tmpHtml & "		a:-webkit-any-link {color:#0cd;cursor:auto;text-decoration:none;font-style:normal;display:inline-block}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div id=""mainBody""><div id=""cloud1"" class=""cloud""></div><div id=""cloud2"" class=""cloud""></div><div id=""cloud3"" class=""cloud""></div></div>" & vbCrlf

	Response.Write "<div class=""loginBox"">" & vbCrlf
	Response.Write "	<div class=""loginTips""><h1><img src=""" & InstallDir & "Static/images/Wmu2Logo.png""></h1><h2><span>教师教学业绩考核管理系统</span></h2></div>" & vbCrlf
	Response.Write "	<ul class=""loginform""><li class=""loginTit""><b>教师登陆</b><span>Teacher Login</span></li>" & vbCrlf
	Response.Write "		<li class=""hr-fix""><input name=""loginuser"" id=""loginuser"" type=""number"" class=""loginInput loginuser"" maxlength=""6"" min=""100000"" max=""999999"" value="""" placeholder=""工号"" /></li>" & vbCrlf
	Response.Write "		<li class=""hr-fix""><input name=""loginpwd"" id=""loginpwd"" type=""password"" class=""loginInput loginpwd"" autocomplete=""off"" value="""" pattern=""[a-zA-Z]\w{5,17}"" placeholder=""密码"" /></li>" & vbCrlf
	Response.Write "		<li class=""hr-fix"">" & vbCrlf
	Response.Write "			<button class=""weui-btn weui-btn_primary"" id=""loginPost"">登　录</button>" & vbCrlf
	Response.Write "		</li>" & vbCrlf
	Response.Write "	</ul>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<footer class=""LoginFooter""><div class=""loginbm""><a href=""http://www.wzhealth.com/"" target=""_blank"">&copy;温州医科大学二院</a> V1.0.0</div></footer>" & vbCrlf
	Response.Write "<div id=""TouchQR"" class=""hide"" style=""display: none;""></div>" & vbCrlf

	Response.Write "		" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#loginPost"").on(""click"", function(){ Gologin(); });" & vbCrlf
	If Action = "Logout" Then
		Response.Cookies(Site_Sn)("YGDM") = ""
		Response.Cookies(Site_Sn)("UserPass") = ""
		Response.Cookies(Site_Sn)("RndCode") = ""
		tmpHtml = tmpHtml & "	$.alert(""成功退出登陆！"",function(){ location.href=""" & ParmPath & "Login/Index.html"" });" & vbCrlf
	End If
	tmpHtml = tmpHtml & "	function Gologin() {" & vbCrlf
	tmpHtml = tmpHtml & "		var loginName = $(""#loginuser"").val(), loginPass = $(""#loginpwd"").val();" & vbCrlf
	tmpHtml = tmpHtml & "		$.getJSON(""" & ParmPath & "Login/ChkLogin.html"", { A: ""Login"", Name: loginName, Pass: loginPass, rand: Math.random() }, function (data) {" & vbCrlf
	tmpHtml = tmpHtml & "			if (data.code==0){ location.href = """ & InstallDir & "Touch/Index.html""; } else {" & vbCrlf
	tmpHtml = tmpHtml & "				$.alert(data.msg);" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	}" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1) & "" & vbCrlf
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub ChkLogin()
	Dim tRndCode : tRndCode = MD5(GetRndString(6), 16)
	Dim rsChk, sqlChk : ErrMsg = ""
	Dim tName : tName = HR_Clng(Request("Name"))
	Dim tPass : tPass = Trim(Request("Pass"))

	If HR_IsNull(tName) Or HR_IsNull(tPass) Then ErrMsg = "登陆帐号或密码不能为空！"
	If HR_Clng(tName) > 999999 Or HR_Clng(tName) < 100000 Then ErrMsg = "您的工号不正确，请重新输入！" & tName
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""code"":500,""msg"":""" & ErrMsg & """}" : Exit Sub

	'------ 判断教师工号登陆
	sqlChk = "Select * From HR_Teacher Where TeacherID>0"
	sqlChk = sqlChk & " And YGDM='" & tName & "' And LoginPass='" & MD5(tPass, 16) & "'"
	Set rsChk = Server.CreateObject("ADODB.RecordSet")
		rsChk.Open(sqlChk), Conn, 1, 3
		If Not(rsChk.BOF And rsChk.EOF) Then
			If HR_Clng(rsChk("ApiType")) = 5 Then
				Response.Write "{""code"":500,""msg"":""帐号已锁定，无法登陆""}" : Exit Sub
			End If
			UserYGDM = rsChk("YGDM")
			UserYGXM = rsChk("YGXM")
			UserRank = 0
			Response.Cookies(Site_Sn)("YGDM") = rsChk("YGDM")
			Response.Cookies(Site_Sn)("UserPass") = rsChk("LoginPass")
			Response.Cookies(Site_Sn)("RndCode") = tRndCode
			rsChk("LoginTime") = Now()
			rsChk("LoginIP") = UserTrueIP
			rsChk.Update
			Response.Write "{""code"":0,""msg"":""成功登陆""}"
		Else
			ErrMsg = "登陆帐号或密码不正确！"
			Response.Write "{""code"":500,""msg"":""" & ErrMsg & """}"
		End If
		rsChk.Close
	Set rsChk = Nothing
End Sub
%>