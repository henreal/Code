<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incMD5.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="incCommon.asp"-->
<!--#include file="incPurview.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl : SiteTitle = SiteName

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "Start" Call Start()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tHeadFace, strJson : strJson = GetWechatUserInfoQY(UserYGDM)
	Dim jsonOBJ
	If Instr(strJson, "avatar") > 0 Then
		Set jsonOBJ = parseJSON(strJson)
			If jsonOBJ.errcode = 0 Then
				tHeadFace = Trim(jsonOBJ.avatar)
			End If
		Set jsonOBJ = Nothing
	End If
	If HR_IsNull(tHeadFace) Then tHeadFace = InstallDir & "Static/images/nopic.png"

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body,html {font-size:14px;overflow:hidden;position: relative;} .bordMain{height:100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.nav {position: relative;height:58px;z-index:100;overflow:hidden;border-bottom:1px solid #012;color:#fff;background:#012 url(" & InstallDir & "Static/admin/topBG_right.png) bottom right no-repeat;}" & vbCrlf
	tmpHtml = tmpHtml & "		.nav dl {height:57px;line-height:57px;font-size: 1.1rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.nav dt {padding: 0 10px;color: #700;font-size: 2.1rem;cursor: pointer;} .nav dt img {height:40px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.nav dd.navTitle {font-size: 1.5rem;} .nav dd.more {padding:0 10px; cursor:pointer;}" & vbCrlf
	tmpHtml = tmpHtml & "		.nav dd.navTitle .closeLeft {font-size: 1.2rem;cursor:pointer;color:#f30;} .nav dd.navTitle .closeLeft:hover {color:#f80;}" & vbCrlf
	tmpHtml = tmpHtml & "		.nav dd.search input {height:25px;line-height:25px;padding:0 8px 0 12px;border:1px solid #900;border-top-left-radius:27px;border-bottom-left-radius:27px;color:#900}" & vbCrlf
	tmpHtml = tmpHtml & "		.nav dd.soBtn button {width:35px;height:27px;line-height:27px;cursor:pointer;border:1px solid #900;background-color:#900;color:#fff;border-top-right-radius:27px;border-bottom-right-radius:27px;padding:0 10px 0 5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.nav dd.myinfo em {margin:0 5px 0 10px;width:40px;height:40px;border:1px solid #fff;box-sizing: border-box;border-radius: 100%;background-size: 100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.nav dd.name em {padding-right:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.nav a {color:#fff;} .nav a:hover {color:#fc0;}" & vbCrlf

	tmpHtml = tmpHtml & "		.leftMenu {position:absolute;border-right:1px solid #012;width:150px;left:0;top:59px;bottom:0;background-color:#135;color:#fff;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-nav-tree {width:100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-nav .layui-nav-item a {padding:0 5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-nav-tree .layui-nav-child dd.layui-this, .layui-nav-tree .layui-nav-child dd.layui-this a, .layui-nav-tree .layui-this, .layui-nav-tree .layui-this>a, .layui-nav-tree .layui-this>a:hover {background-color:#520f0f;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-nav-tree .layui-nav-item {border-bottom:1px solid #515663;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-navmenu-main {background-color:#393d49;border-bottom:1px solid #393d49}" & vbCrlf

	tmpHtml = tmpHtml & "		.centerBody {position:absolute;left:151px;top:59px;right:0;bottom:0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.leftMenu .MenuBox, .centerBody .MainBox {width:100%;height:100%;overflow-y:auto;overflow-x:hidden;box-sizing:border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.centerBody .MainBox {overflow:hidden;}" & vbCrlf
	tmpHtml = tmpHtml & "		.iframeMain {width:100%;height:100%;border:0;overflow:hidden;box-sizing:border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.centerBody .content p.title {font-weight:bold;padding-top:30px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead("Index", 1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<header class=""nav"">" & vbCrlf
	Response.Write "	<dl class=""hr-rows"">" & vbCrlf
	Response.Write "		<dt><img src=""" & InstallDir & "Static/images/uLogo1.png"" alt=""" & SiteName & """ /></dt><dd class=""hr-grow navTitle"">" & SiteName & " <span class=""closeLeft"" data-value=""open""><i class=""hr-icon"">&#xed6d;</i></span></dd>" & vbCrlf
	Response.Write "		<dd class=""search""><input type=""text"" name=""soWord"" class=""inputtxt"" id=""soWord"" placeholder=""请输入关键字""></dd>" & vbCrlf
	Response.Write "		<dd class=""soBtn""><button name=""soPost"" class=""btn1"" id=""soPost"" type=""button""><i class=""hr-icon"">&#xf35f;</i></button></dd>" & vbCrlf
	Response.Write "		<dd class=""myinfo""><em class=""headface"" style=""background-image:url(" & tHeadFace & ");"" title=""" & UserYGXM & " [" & UserYGDM & "]""></em></dd>" & vbCrlf
	Response.Write "		<dd class=""name""><em>" & UserYGXM & "[" & UserYGDM & "]</em></dd>" & vbCrlf
	If UserRank > 0 Then Response.Write "		<dd class=""hr-rows manage""><em class=""layui-anim layui-anim-rotate layui-anim-loop""><i class=""hr-icon"""">&#xec85;</i></em><em><a href=""" & InstallDir & ManageDir & "Index.html"">管理面板</a></em></dd>" & vbCrlf
	Response.Write "		<dd class=""more""><i class=""hr-icon"">&#xeca7;</i><a href=""" & ParmPath & "Login.html?Logout=True"">退出</a>　<i class=""hr-icon"">&#xf101;</i></dd>" & vbCrlf
	Response.Write "	</dl>" & vbCrlf
	Response.Write "</header>" & vbCrlf
	Response.Write "<div class=""leftMenu"">" & vbCrlf
	Response.Write "	<div class=""MenuBox"" id=""menubox"">" & vbCrlf
	Response.Write "		<ul class=""layui-nav layui-nav-tree menu-tree"" lay-shrink=""all"">" & vbCrlf
	Response.Write "			<li class=""layui-nav-item menu-parent"" data-id=""A111"">" & vbCrlf
	Response.Write "				<a href=""#"" class=""parenthref"" title=""基础性教学业绩""><i class=""hr-icon"">&#xe1b2;</i>基础性教学</a>" & vbCrlf
	Response.Write "				<dl class=""layui-nav-child"">" & ShowNavMenu(1) & "</dl>" & vbCrlf
	Response.Write "			</li>" & vbCrlf
	Response.Write "			<li class=""layui-nav-item"" data-id=""A111"">" & vbCrlf
	Response.Write "				<a href=""#"" class=""subhref"" title=""激励性教学业绩""><i class=""hr-icon"">&#xe8a3;</i>激励性教学</a>" & vbCrlf
	Response.Write "				<dl class=""layui-nav-child"">" & ShowNavMenu(2) & "</dl>" & vbCrlf
	Response.Write "			</li>" & vbCrlf
	Response.Write "			<li class=""layui-nav-item""><a href=""" & ParmPath & "Achieve/Mine.html""  target=""iframeMain"" title=""我的业绩""><i class=""hr-icon"">&#xe8a3;</i>我的业绩</a></li>" & vbCrlf
	Response.Write "			<li class=""layui-nav-item"" data-id=""A111"">" & vbCrlf
	Response.Write "				<a href=""#"" class=""subhref"" title=""日常工作""><i class=""hr-icon"">&#xe94e;</i>日常工作</a>" & vbCrlf
	Response.Write "				<dl class=""layui-nav-child"">" & vbCrlf
	Response.Write "					<dd data-name=""console""><a href=""" & ParmPath & "Notice/Index.html"" target=""iframeMain""><i class=""hr-icon hr-icon-top"">&#xf31d;</i>查看通知</a></dd>" & vbCrlf
	Response.Write "					<dd data-name=""console""><a href=""" & ParmPath & "Message/Index.html"" target=""iframeMain""><i class=""hr-icon hr-icon-top"">&#xf31d;</i>我的消息</a></dd>" & vbCrlf
	Response.Write "					<dd data-name=""console""><a href=""" & ParmPath & "Setup/ModiPass.html"" target=""iframeMain""><i class=""hr-icon hr-icon-top"">&#xf31d;</i>修改密码</a></dd>" & vbCrlf
	Response.Write "				</dl>" & vbCrlf
	Response.Write "			</li>" & vbCrlf
	Response.Write "		</ul>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""centerBody"">" & vbCrlf
	Response.Write "	<div class=""MainBox"" id=""MainBox""><iframe id=""iframeMain"" src=""Index/Start.html"" class=""iframeMain"" name=""iframeMain""></iframe></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf


	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".closeLeft"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var showLeft = $(this).data(""value"");" & vbCrlf
	tmpHtml = tmpHtml & "		if(showLeft==""open""){" & vbCrlf
	tmpHtml = tmpHtml & "			$("".leftMenu"").animate({width:""0px""});" & vbCrlf
	tmpHtml = tmpHtml & "			$(this).data(""value"",""close"");" & vbCrlf
	tmpHtml = tmpHtml & "			$("".centerBody"").animate({left:""0px""});" & vbCrlf
	tmpHtml = tmpHtml & "			$("".closeLeft"").html(""<i class='hr-icon'>&#xed6c;</i>"");" & vbCrlf
	tmpHtml = tmpHtml & "		}else if(showLeft==""close""){" & vbCrlf
	tmpHtml = tmpHtml & "			$("".leftMenu"").animate({width:""150px""});" & vbCrlf
	tmpHtml = tmpHtml & "			$(this).data(""value"",""open"");" & vbCrlf
	tmpHtml = tmpHtml & "			$("".centerBody"").animate({left:""151px""});" & vbCrlf
	tmpHtml = tmpHtml & "			$("".closeLeft"").html(""<i class='hr-icon'>&#xed6d;</i>"");" & vbCrlf
	tmpHtml = tmpHtml & "		}" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot("Index", 0) & "" & vbCrlf
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub

Sub Start()
	Call ChkInitPass()	'//检查初始密码
	Dim tHeadFace, jsonOBJ, strJson : strJson = GetWechatUserInfoQY(UserYGDM)
	Dim tEmail, tMobile, tQRCode, tLoginTime, tKSMC, tPRZC
	Set jsonOBJ = parseJSON(strJson)
		If jsonOBJ.errcode = 0 Then
			tHeadFace = Trim(jsonOBJ.avatar)
			tEmail = Trim(jsonOBJ.email)
			tMobile = Trim(jsonOBJ.mobile)
			tQRCode = Trim(jsonOBJ.qr_code)
		End If
	Set jsonOBJ = Nothing
	'Response.Write UserYGDM : Response.End
	If HR_IsNull(tHeadFace) Then tHeadFace = InstallDir & "Static/images/nopic.png"
	tQRCode = tHeadFace
	Set rs = Conn.Execute("Select * From HR_Teacher Where YGDM='" & UserYGDM & "'")
		If Not(rs.BOF And rs.EOF) Then
			tLoginTime = Formatdate(rs("LoginTime"), 1)
			tKSMC = Trim(rs("KSMC"))
			tPRZC = Trim(rs("PRZC"))
		End If
	Set rs = Nothing

	Dim SumScore : SumScore =0
	Dim myGrade : myGrade = ""
	Dim msgNum : msgNum = 0
	Dim noticeNum : noticeNum = 0
	Set rsTmp = Conn.Execute("Select top 1 SumScore,Grade From HR_KPI_SUM Where YGDM>0 And YGDM=" & HR_Clng(UserYGDM))
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			SumScore = HR_CDbl(rsTmp(0))
			myGrade = Trim(rsTmp(1))
		End If
	Set rsTmp = Nothing
	Set rsTmp = Conn.Execute("Select count(ID) From HR_Message Where ReceiverID=" & HR_Clng(UserYGDM))
		msgNum = HR_CDbl(rsTmp(0))
	Set rsTmp = Nothing
	'3日内发布的通知
	Set rsTmp = Conn.Execute("Select count(0) From HR_Notice Where ID>0 And DATEDIFF(""d"", PublishesTime, getDate())<4")
		noticeNum = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body,html {background-color:#f1f1f1;} .bordMain{height:100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-panel-tit {line-height:42px;height:42px;font-size:1.1rem}" & vbCrlf
	tmpHtml = tmpHtml & "		.headface em {margin:0 5px 0 10px;width:80px;height:80px;border:1px solid #fff;box-sizing: border-box;border-radius: 100%;background-size: 100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.myInfo {flex-grow:2;padding-left:10px} .myInfo h2{font-size:1.2rem;} .myInfo h3{font-size:1rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.myQRCode em {width:80px;height:80px;border:1px solid #eee;box-sizing: border-box;background-size: 100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.bordMain{height:100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.modiPass {color:#f30;cursor:pointer;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead("Index", 1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf

	tmpHtml = "	<div class=""layui-row layui-col-space15"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-col-xs6"">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-card"">" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-card-header""><dl class=""hr-rows hr-panel-tit""><dt>我的信息</dt><dd class=""modiPass"">修改密码</a><i class=""hr-icon"">&#xef91;</i></dd></dl></div>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-card-body"">" & vbCrlf
	tmpHtml = tmpHtml & "					<div class=""hr-rows"">" & vbCrlf
	tmpHtml = tmpHtml & "						<div class=""headface""><em style=""background-image:url(" & tHeadFace & ");"" title=""" & UserYGXM & " [" & UserYGDM & "]""></em></div>" & vbCrlf
	tmpHtml = tmpHtml & "						<div class=""myInfo"">" & vbCrlf
	tmpHtml = tmpHtml & "							<h2>姓　名：" & UserYGXM & "</h2>" & vbCrlf
	tmpHtml = tmpHtml & "							<h2>工　号：" & UserYGDM & "</h2>" & vbCrlf
	tmpHtml = tmpHtml & "							<h3>科　室：" & tKSMC & "</h3>" & vbCrlf
	tmpHtml = tmpHtml & "							<h3>职　称：" & tPRZC & "</h3>" & vbCrlf
	tmpHtml = tmpHtml & "							<h3>登陆时间：" & tLoginTime & "</h3>" & vbCrlf
	tmpHtml = tmpHtml & "						</div>" & vbCrlf
	tmpHtml = tmpHtml & "						<div class=""myQRCode"">" & vbCrlf
	tmpHtml = tmpHtml & "							<em style=""background-image:url(" & tQRCode & ");"" title=""我的二维码""></em>" & vbCrlf
	tmpHtml = tmpHtml & "						</div>" & vbCrlf
	tmpHtml = tmpHtml & "					</div>" & vbCrlf
	tmpHtml = tmpHtml & "				</div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-col-xs3"">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-card"">" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-card-header""><dl class=""hr-rows hr-panel-tit""><dt>业绩信息</dt></dl></div>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-card-body"">" & vbCrlf
	tmpHtml = tmpHtml & "					<div class=""yj"">学时数：" & SumScore & "<br>等　级：" & myGrade & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "					<div class=""yj"">学　年：" & DefYear-1 & "-" & DefYear & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "					<div class=""yj"">初始密码：AfVHy4k886，MD5值：" & MD5("AfVHy4k886",16) & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "				</div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-col-xs3"">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-card"">" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-card-header""><dl class=""hr-rows hr-panel-tit""><dt>我的提醒</dt><dd>更多</a><i class=""hr-icon"">&#xef91;</i></dd></dl></div>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-card-body"" title=""3日内发布的通知"">新通知：" & NoticeNum & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-card-body"">新消息：" & msgNum & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf

	'//判断密码是否为初始密码
	Dim ArrField : ArrField = GetTableDataQuery("HR_Teacher", "", 1, "YGDM='" & UserYGDM & "'")			'取教师信息

	Response.Write tmpHtml & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-row layui-col-space15"">" & vbCrlf
	'Response.Write "		" & strJson
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".closeLeft"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var showLeft = $(this).data(""value"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var element = layui.element, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		$("".modiPass"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			layer.open({type:2, content:""" & ParmPath & "UserCenter/ModifyPass.html"", title:[""修改密码"",""font-size:16""],area:[""560px"", ""300px""],moveOut:true});" & vbCrlf
	tmpHtml = tmpHtml & "			" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot("Index", 0) & "" & vbCrlf
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub

Function ShowNavMenu(fType)
	Dim strMenu, rsMenu, sqlMenu, fParentID, fItemName, fOrder, fIcon
	fType = HR_Clng(fType) : If fType = 0 Then fType = 1
	sqlMenu = "Select * From HR_Class Where ClassType=" & fType & " Order By RootID ASC, OrderID ASC"
	Set rsMenu = Conn.Execute(sqlMenu)
		If Not(rsMenu.BOF And rsMenu.EOF) Then
			strMenu = strMenu & vbCrlf
			Do While Not rsMenu.EOF
				fParentID = HR_Clng(rsMenu("ParentID"))
				fItemName = Trim(rsMenu("ClassName"))
				If fParentID > 0 Then
					Set rsTmp = Conn.Execute("Select Max(OrderID) From HR_Class Where ParentID=" & fParentID)
						fOrder = HR_Clng(rsTmp(0))
					Set rsTmp = Nothing
				End If
				fIcon = "<i class=""hr-icon"">&#xf31d;</i>"

				If fParentID > 0 And HR_Clng(rsMenu("OrderID")) <> fOrder Then strMenu = strMenu & "		"
				If fParentID > 0 And HR_Clng(rsMenu("OrderID")) <= fOrder Then fIcon = "<i class=""hr-icon"">&#xf328;</i>"
				If HR_Clng(rsMenu("OrderID")) = fOrder And fParentID > 0 Then strMenu = strMenu & "		"
				strMenu = strMenu & "<dd data-name=""console"">"
				If HR_Clng(rsMenu("Child")) > 0 Then strMenu = strMenu & vbCrlf & "	"
				strMenu = strMenu & "<a "
				If HR_Clng(rsMenu("Child")) > 0 Then
					strMenu = strMenu & " class=""hr-navmenu-main"" href=""#"">"
				Else
					strMenu = strMenu & "href=""" & ParmPath & "Course.html?ItemID=" & rsMenu("ClassID") & """ target=""iframeMain"">"
				End If
				strMenu = strMenu & fIcon & fItemName & "</a>"
				If HR_Clng(rsMenu("Child")) > 0 Then strMenu = strMenu & vbCrlf & "	<dl class=""layui-nav-child"">" & vbCrlf
				If HR_Clng(rsMenu("Child")) = 0 Then strMenu = strMenu & "</dd>" & vbCrlf
				If HR_Clng(rsMenu("OrderID")) = fOrder And fParentID > 0 Then strMenu = strMenu & "	</dl>" & vbCrlf & "</dd>" & vbCrlf
				rsMenu.MoveNext
			Loop
		End If
	Set rsMenu = Nothing
	ShowNavMenu = strMenu
End Function
%>