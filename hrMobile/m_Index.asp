<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="incCommon.asp"-->
<!--#include file="incPurview.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl : SiteTitle = "教师业绩考核"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "ForgetPass" Call ForgetPass()
	Case Else Call MainBody()
End Select

Sub MainBody()
	
	Dim SumScore : SumScore =0
	Dim msgNum : msgNum = 0
	Dim myGrade
	If HR_Clng(UserYGDM) > 0 Then
		Set rsTmp = Conn.Execute("Select top 1 SumScore,Grade From HR_KPI_SUM Where YGDM>0 And YGDM=" & HR_Clng(UserYGDM) & " And scYear=" & DefYear-1)
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				SumScore = HR_CDbl(rsTmp(0))
				myGrade = Trim(rsTmp(1))
			End If
		Set rsTmp = Nothing
		Set rsTmp = Conn.Execute("Select count(ID) From HR_Message Where ReceiverID=" & HR_CLng(UserYGDM))
			msgNum = HR_CDbl(rsTmp(0))
		Set rsTmp = Nothing
	End If
	If HR_IsNull(myGrade) Then myGrade = "-"

	Dim newMsgNum : newMsgNum = 0
	Set rsTmp = Conn.Execute("Select count(ID) From HR_Message Where isRead=" & HR_False & " And ReceiverID=" & HR_CLng(UserYGDM))
		newMsgNum = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing
	
	Dim jsonOBJ , userJson : userJson = GetWechatUserInfoQY(UserYGDM)
	If Instr(userJson, "avatar") > 0 Then
		Set jsonOBJ = parseJSON(userJson)
			If jsonOBJ.errcode = 0 Then
				HeadFace = Trim(jsonOBJ.avatar)
			End If
		Set jsonOBJ = Nothing
	End If
	If HR_IsNull(HeadFace) Then HeadFace = InstallDir & "Static/images/nopic.png"

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.TeachYear {text-align:center;font-size:1.2rem;} .TeachYear em {margin:3px auto;font-size:1.5rem; font-weight: bold;color:#fff;width:50%;background-color:#03A9F4;border-radius:1.5rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-grid {padding:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-grid__icon {width:45px;height:auto;font-size:1.6rem;text-align: center}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-grid__label {font-size:1rem;line-height:2rem}" & vbCrlf
	tmpHtml = tmpHtml & "		.msg1 {font-size:1rem;margin-top:0px} .msg1 .weui-cell {padding:8px 5px}" & vbCrlf
	tmpHtml = tmpHtml & "		.msgTit b {padding-right:8px;margin-right:8px;border-right:2px solid #ccc}" & vbCrlf
	tmpHtml = tmpHtml & "		.msgTit p {width:250px;overflow:hidden;-o-text-overflow:ellipsis;text-overflow:ellipsis;white-space:nowrap;}" & vbCrlf

	tmpHtml = tmpHtml & "		.hr-grids {border-bottom:1px solid #ddd;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item {width:33.3%;box-sizing:border-box;text-align:center;padding:8px 0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item em:first-child {width:40px;height:40px;line-height:40px;text-align:center;margin:0 auto;background-color:#f90;color:#fff;border-radius: 40px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item i {font-size:1.5rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item em {font-size:0.8rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item:nth-child(6n+1) em:first-child {background-color:#0bf;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item:nth-child(6n+2) em:first-child {background-color:#2da;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item:nth-child(6n+3) em:first-child {background-color:#b6c;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item:nth-child(6n+4) em:first-child {background-color:#ca4;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-grids .grid-item:nth-child(6n+5) em:first-child {background-color:#5b6;}" & vbCrlf
	tmpHtml = tmpHtml & "		.tips1 {background-color:#ffeec3;color:#f30}" & vbCrlf

	tmpHtml = tmpHtml & "		.pop_tit {padding:10px;border-bottom:1px solid #ccc;}" & vbCrlf
	tmpHtml = tmpHtml & "		.pop_box {padding:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.pop_box dl {padding:5px 0;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<header class=""hr-rows hr-header"">" & vbCrlf
	Response.Write "	<nav class=""navLogo""><em></em></nav>" & vbCrlf
	Response.Write "	<nav class=""navTitle""><span>" & SiteTitle & "</span></nav>" & vbCrlf
	Response.Write "	<nav class=""navMenu""><em><i class=""hr-icon"">&#xeef7;</i></em></nav>" & vbCrlf
	Response.Write "</header>" & vbCrlf

	Dim tNavMenu
	tNavMenu = "<li><a href=""Index.html""><i class=""hr-icon hr-icon-top"">&#xec58;</i>返　回</a></li>" & vbCrlf
	tNavMenu = tNavMenu & "		<li><a href=""Achieve/Index.html?A=List""><i class=""hr-icon hr-icon-top"">&#xec8d;</i>查看业绩<sup></sup></a></li>" & vbCrlf
	tNavMenu = tNavMenu & "		<li><a href=""myCenter/Message.html""><i class=""hr-icon hr-icon-top"">&#xeea0;</i>我的消息<sup></sup></a></li>" & vbCrlf
	If UserRank > 0 Then tNavMenu = tNavMenu & "		<li><a href=""Manage/Index.html""><i class=""hr-icon hr-icon-top"">&#xeab5;</i>管理面板<sup></sup></a></li>" & vbCrlf
	tNavMenu = tNavMenu & "		<li><a href=""" & ParmPath & "Login/Logout.html?noBind=1""><i class=""hr-icon hr-icon-top"">&#xeca7;</i>退出登陆</a><sup></sup></li>" & vbCrlf

	%>
<div class="nctouch-nav-layout layerNav" style="display: none;">
	<div class="nctouch-nav-menu">
	<span class="arrow"></span>
	<ul><%=tNavMenu %></ul>
	</div>
</div>
<div class="hr-fix hr-header-hide"></div>
<div class="scroller-body">
  <div class="scroller-box">
    <div class="member-top">
		<div class="member-info">
			<div class="user-avatar"><img src="<%=HeadFace %>" alt="<%=UserYGXM %>" /> </div>
			<div class="user-name"> <span><%=UserYGXM %><%
			If newMsgNum > 0 Then
				Response.Write "<sup>" & newMsgNum & "</sup>"
			End If
			%></span> </div>
		</div>
		<div class="member-collect">
			<span><a href="Achieve/Index.html"><em><%=SumScore %></em><tt>学时数</tt></a></span>
			<span><a href="myCenter/Message.html"><em><%=msgNum %></em><tt>我的消息</tt></a></span>
			<span><a href="javascript:;" class="open-popup mygrade" data-target="grade-con"><em><%=myGrade %></em><tt>等级</tt></a></span>
		</div>
    </div>
  </div>
  <div class="hr-rows tips1"><em class="hr-item">学时数及等级为上一学年</em><tt class="hr-item offtips1"><i class="hr-icon hr-icon-top">&#xee31;</i></tt></div>
</div>

<%
	Dim rsMsg, msgTitle, msgTime
	Set rsMsg = Conn.Execute("Select Top 1 * From HR_Notice Order By PublishesTime DESC")
		If Not(rsMsg.BOF And rsMsg.EOF) Then
			msgTitle = Trim(rsMsg("Title"))
			msgTime = FormatDate(rsMsg("PublishesTime"), 5)
		End If
	Set rsMsg = Nothing

	Response.Write "<div class=""hr-fix TeachYear""><p>当前学年</p><em>" & DefYear-1 & " - " & DefYear & "</em></div>" & vbCrlf
	Response.Write "<div class=""weui-cells msg1"">" & vbCrlf
	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "Notice/Index.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd"" style=""color:#f30""><i class=""hr-icon"">&#xe972;</i></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd weui-cell_primary msgTit"" data-id=""123""><p><b>通知</b>" & msgTitle & "</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & msgTime & "</div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "<div class=""hr-grids"">" & vbCrlf
	Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "Remain/Passed.html""><em><i class=""hr-icon"">&#xe960;</i></em><em>未审业绩</em></a>" & vbCrlf
	Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "Remain/Retreat.html""><em><i class=""hr-icon"">&#xf338;</i></em><em>退回业绩</em></a>" & vbCrlf
	Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "Remain/Affirm.html""><em><i class=""hr-icon"">&#xeac4;</i></em><em>确认业绩</em></a>" & vbCrlf
	Response.Write "</div>" & vbCrlf
%>
<div class="hr-gap-20 hr-fix"></div>
<div class="hr-panel-item hr-fix">
	<div class="hr-flex_item" data-type="1">
		<a class="hr-navmenu" href="Course/ListItem.html?TypeID=1">
			<em class="title"><i class="hr-icon">&#xe1b2;</i>添加基础性教学</em>
			<em class="tips">共<b>7</b>类</em>
			<em class="more"><i class="hr-icon">&#xf054;</i></em>
		</a>
	</div>
	<div class="hr-flex_item" data-type="2">
		<a class="hr-navmenu" href="Course/ListItem.html?TypeID=2">
			<em class="title"><i class="hr-icon">&#xe8a3;</i>添加激励性教学</em>
			<em class="tips">共<b>11</b>类</em>
			<em class="more"><i class="hr-icon">&#xf054;</i></em>
		</a>
	</div>
</div>
<div class="hr-gap-20 hr-fix"></div>
<div class="NavBar_Body hr-fix">
	<div class="weui-grids">
      <a href="Achieve/Index.html?A=List" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="我的业绩">&#xeee7;</i></div>
        <p class="weui-grid__label">我的业绩</p>
      </a>
	  <a href="myCenter/Index.html" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="我的信息">&#xeeef;</i></div>
        <p class="weui-grid__label">个人档案</p>
      </a>
	  <a href="myCenter/Pass.html" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="密码修改">&#xee75;</i></div>
        <p class="weui-grid__label">密码修改</p>
      </a>
	  <a href="<%=ParmPath %>Notice/Index.html" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="通知公告">&#xe7f7;</i></div>
        <p class="weui-grid__label">通知公告</p>
      </a>
	  <a href="<%=ParmPath %>CourseSelect.html" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="评价">&#xead1;</i></div>
        <p class="weui-grid__label">课堂教学评价</p>
      </a>
	  <a href="<%=ParmPath %>Evaluate/CEX/EditCEX.html?AddNew=True" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="评价">&#xead1;</i></div>
        <p class="weui-grid__label">CEX<sup>plus</sup>记录</p>
      </a>
	  <a href="<%=ParmPath %>Schedule/Index.html" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="日程">&#xf073;</i></div>
        <p class="weui-grid__label">授课日程</p>
      </a>
	  <a href="<%=ParmPath %>Directories/Index.html" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="通讯录">&#xf2ba;</i></div>
        <p class="weui-grid__label">通讯录</p>
      </a>
	  <a href="<%=ParmPath %>Remain/Index.html" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="待办">&#xf022;</i></div>
        <p class="weui-grid__label">待办</p>
      </a>
	  <a href="<%=ParmPath %>SwapCourse/Edit.html?AddNew=True" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="对换课">&#xed1e;</i></div>
        <p class="weui-grid__label">换课</p>
      </a>
	  <a href="<%=ParmPath %>Substitute/Edit.html?AddNew=True" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="代课记录">&#xeeed;</i></div>
        <p class="weui-grid__label">代课</p>
      </a>
	  <a href="<%=ParmPath %>Propose/Index.html" class="weui-grid item_grid">
        <div class="weui-grid__icon"><i class="hr-icon" title="建议意见">&#xf298;</i></div>
        <p class="weui-grid__label">建议意见</p>
      </a>
	</div>
</div>
<div class="hr-gap-20 hr-fix"></div>
<%
	Dim SwapNum, SwapNum1, SwapPass
	Dim startTime, endTime
	startTime = DefYear-1 & "-07-01 00:00:00"
	endTime = DefYear & "-06-30 23:59:59"

	Set rs = Conn.Execute("Select Count(0) From HR_Swap Where (ApplyTime Between '" & startTime & "' And '" & endTime & "') And newItemID>0 And newCourseID>0 And Passer=" & UserYGDM)		'教研主任审核（换课）
		SwapNum = HR_Clng(rs(0))
	Set rs = Nothing
	SwapPass = GetTypeName("HR_User", "SwapPass", "YGDM", UserYGDM)		'教学处或教辅审核
	If HR_CLng(SwapPass) > 0 Then
		Set rs = Conn.Execute("Select Count(0) From HR_Swap Where newItemID>0 And newCourseID>0 And (ApplyTime Between '" & startTime & "' And '" & endTime & "')")
			SwapNum = HR_Clng(rs(0))
		Set rs = Nothing
	End If

	Set rs = Conn.Execute("Select Count(0) From HR_Swap Where (ApplyTime Between '" & startTime & "' And '" & endTime & "') And newItemID=0 And newCourseID=0 And Passer=" & UserYGDM)		'教研主任审核（代课）
		SwapNum1 = HR_Clng(rs(0))
	Set rs = Nothing
	If HR_CLng(SwapPass) > 0 Then
		Set rs = Conn.Execute("Select Count(0) From HR_Swap Where newItemID=0 And newCourseID=0 And (ApplyTime Between '" & startTime & "' And '" & endTime & "')")
			SwapNum1 = HR_Clng(rs(0))
		Set rs = Nothing
	End If

	If SwapNum > 0 Then
		Response.Write "<div class=""hr-panel-item hr-fix"">" & vbCrlf
		Response.Write "	<div class=""hr-flex_item"" data-type=""3"">" & vbCrlf
		Response.Write "		<a class=""hr-navmenu"" href=""" & ParmPath & "SwapPass/Index.html"">" & vbCrlf
		Response.Write "			<em class=""title""><i class=""hr-icon"">&#xec66;</i>审核换课申请</em>" & vbCrlf
		Response.Write "			<em class=""tips"">共<b>" & SwapNum & "</b>条</em><em class=""more""><i class=""hr-icon"">&#xf054;</i></em>" & vbCrlf
		Response.Write "		</a>" & vbCrlf
		Response.Write "	</div>" & vbCrlf
		Response.Write "</div>" & vbCrlf
	End If
	If SwapNum1 > 0 Then
		Response.Write "<div class=""hr-panel-item hr-fix"">" & vbCrlf
		Response.Write "	<div class=""hr-flex_item"" data-type=""3"">" & vbCrlf
		Response.Write "		<a class=""hr-navmenu"" href=""" & ParmPath & "SubstitutePass/Index.html"">" & vbCrlf
		Response.Write "			<em class=""title""><i class=""hr-icon"">&#xe877;</i>审核代课申请</em>" & vbCrlf
		Response.Write "			<em class=""tips"">共<b>" & SwapNum1 & "</b>条</em><em class=""more""><i class=""hr-icon"">&#xf054;</i></em>" & vbCrlf
		Response.Write "		</a>" & vbCrlf
		Response.Write "	</div>" & vbCrlf
		Response.Write "</div>" & vbCrlf
	End If

	Response.Write "<div class=""weui-popup__container popup-bottom"" id=""grade-con"">" & vbCrlf
	Response.Write "	<div class=""weui-popup__overlay""></div>" & vbCrlf
	Response.Write "	<div class=""weui-popup__modal"">" & vbCrlf
	Response.Write "		<div class=""pop_tit""><span>我的历史等级</span></div>" & vbCrlf
	Response.Write "		<div class=""pop_box"">" & vbCrlf
	Set rsTmp = Conn.Execute("Select scYear, Grade From HR_KPI Where YGDM=" & HR_Clng(UserYGDM) & " Order By scYear DESC")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Dim tGrade
			Do While Not rsTmp.EOF
				tGrade = Trim(rsTmp("Grade"))
				If HR_IsNull(tGrade) Then tGrade = "-"
				Response.Write "			<dl class=""hr-rows""><dt>" & rsTmp("scYear") - 1 & " - " & rsTmp("scYear") & " 学年</dt><dd>等级：" & tGrade & "</dd></dl>" & vbCrlf
				rsTmp.MoveNext
			Loop
		End If
	Set rsTmp = Nothing
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$("".layerNav"").toggle();" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""tt.offtips1"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$("".tips1"").hide();" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	$("".mygrade"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#grade-con"").popup();" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub
%>