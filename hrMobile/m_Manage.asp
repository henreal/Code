<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./m_ManageCommon.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl : SiteTitle = "管理面板"

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "Message" Call Message()
	Case "View" Call View()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim jsonOBJ , userJson : userJson = GetWechatUserInfoQY(UserYGDM)
	Dim arrManageRank : arrManageRank = Split(XmlText("Config", "ManageRank", ""), "|")
	Dim mgStuType : mgStuType = GetTypeName("HR_User", "StuType", "UserID", UserID)

	Dim tKSMC, tPRZC, tYGXB, tXZZW, tYGZT
	Set rsTmp = Conn.Execute("Select * From HR_Teacher Where YGDM='" & UserYGDM & "'")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tKSMC = Trim(rsTmp("KSMC"))
			tYGXB = Trim(rsTmp("YGXB"))
			tPRZC = Trim(rsTmp("PRZC"))
			tXZZW = Trim(rsTmp("XZZW"))
			tYGZT = Trim(rsTmp("YGZT"))
		End If
	Set rsTmp = Nothing
	Set jsonOBJ = parseJSON(userJson)
		If jsonOBJ.errcode = 0 Then
			HeadFace = Trim(jsonOBJ.avatar)
		End If
	Set jsonOBJ = Nothing
	If HR_IsNull(HeadFace) Then HeadFace = InstallDir & "Static/images/nopic.png"

	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells {margin:0;}" & vbCrlf
	strHtml = strHtml & "		.yearbar {padding:6px;text-align:center;border-bottom:1px solid #ddd;font-size:1rem;} .yearbar b {font-size:1.5rem;color:#f30}" & vbCrlf
	strHtml = strHtml & "		.hr-cell .weui-cell__hd {padding-right:5px;color:#f60} .hr-cell .weui-cell__hd i {font-size:1.5rem;}" & vbCrlf
	strHtml = strHtml & "		.UserBody {position: relative;height:3.5rem;margin: 0 auto;background: #F60 url(" & InstallDir & "Static/images/UserBG-750-235.jpg) no-repeat top center;background-size: cover;color: #fff;}" & vbCrlf
	strHtml = strHtml & "		.UserBody .user-photo {position:absolute; bottom: -1.75rem;left: .467rem; width: 3.5rem;height: 3.5rem; text-align: center; border-radius: 100%; border: 2px solid #fff;background:#ffe5be center no-repeat;background-size:100%; -webkit-transform: translateZ(1px); transform: translateZ(1px); overflow: hidden;z-index:8}" & vbCrlf
	strHtml = strHtml & "		.UserBody .user-photo img {width: 100%;height: 100%;border-radius: 100%;background-color:#f50}" & vbCrlf
	strHtml = strHtml & "		" & vbCrlf
	strHtml = strHtml & "		.UserBody .user-nick { float: left;width: 100%; padding-left: 5rem; box-sizing: border-box; margin-top: 1.5rem;font-size: 1.1rem; display: -webkit-box; display: -moz-box; display: -ms-box; display: -o-box; display: box;}" & vbCrlf
	strHtml = strHtml & "		.UserBody .user-nick .level {margin-left: 0.4rem;background-position: left center;background-repeat: no-repeat;background-size:1.3rem; -webkit-box-flex: 1;}" & vbCrlf
	strHtml = strHtml & "		.UserBody .user-nick .level5 {background-image: url(/Static/images/Lv5-40-32.png);}" & vbCrlf
	strHtml = strHtml & "		section {padding-bottom:.4rem; border-bottom:1px solid #ddd;background: #fff;}" & vbCrlf
	strHtml = strHtml & "		.user-behavior {height:auto;text-align:center;-webkit-box-pack:center;-moz-box-pack:center;-ms-box-pack:center;-o-box-pack:center;box-pack:center;-webkit-box-align:center;-moz-box-align:center;-ms-box-align:center;-o-box-align:center;box-align:center;display:-webkit-box;display:-moz-box;display:-ms-box;display:-o-box;display:box}" & vbCrlf
	strHtml = strHtml & "		.user-behavior ul {position:relative;margin-left:3.5rem;margin-top:0.2rem;display:-webkit-box;width:14.5rem;box-sizing:border-box;}" & vbCrlf
	strHtml = strHtml & "		.user-behavior ul li {-webkit-box-flex:1;border-left:1px solid #e7e7e7;font-size: 0.85rem;list-style: none;}" & vbCrlf
	strHtml = strHtml & "		.user-behavior ul li:first-child{border:0}" & vbCrlf
	strHtml = strHtml & "		.user-behavior ul li a {color: #051b28; text-decoration: none;}" & vbCrlf
	strHtml = strHtml & "		.hr-grids {border-bottom:1px solid #ddd;}" & vbCrlf
	strHtml = strHtml & "		.hr-grids .grid-item {width:25%;box-sizing:border-box;text-align:center;padding:8px 0;}" & vbCrlf
	strHtml = strHtml & "		.hr-grids .grid-item em:first-child {width:40px;height:40px;line-height:40px;text-align:center;margin:0 auto;background-color:#f90;color:#fff;border-radius: 40px;}" & vbCrlf
	strHtml = strHtml & "		.hr-grids .grid-item i {font-size:1.5rem;}" & vbCrlf
	strHtml = strHtml & "		.hr-grids .grid-item em {font-size:0.8rem;}" & vbCrlf
	strHtml = strHtml & "		.hr-grids .grid-item:nth-child(6n+1) em:first-child {background-color:#0bf;}" & vbCrlf
	strHtml = strHtml & "		.hr-grids .grid-item:nth-child(6n+2) em:first-child {background-color:#2da;}" & vbCrlf
	strHtml = strHtml & "		.hr-grids .grid-item:nth-child(6n+3) em:first-child {background-color:#b6c;}" & vbCrlf
	strHtml = strHtml & "		.hr-grids .grid-item:nth-child(6n+4) em:first-child {background-color:#ca4;}" & vbCrlf
	strHtml = strHtml & "		.hr-grids .grid-item:nth-child(6n+5) em:first-child {background-color:#5b6;}" & vbCrlf

	
	strHtml = strHtml & "		.grid2 {border:0}" & vbCrlf
	strHtml = strHtml & "		.grid2 .grid-item {border-right:1px solid #ddd;border-bottom:1px solid #ddd;font-size:0.8rem;padding:3px 0}" & vbCrlf
	strHtml = strHtml & "		.grid2 .grid-item:nth-child(4n) {border-right:0}" & vbCrlf
	strHtml = strHtml & "		.grid2 .grid-item em:first-child {background-color:transparent;color:#f90;line-height:35px;height:35px}" & vbCrlf
	strHtml = strHtml & "		.grid2 .grid-item:nth-child(6n+1) em:first-child {background-color:transparent;color:#0bf;}" & vbCrlf
	strHtml = strHtml & "		.grid2 .grid-item:nth-child(6n+2) em:first-child {background-color:transparent;color:#2da;}" & vbCrlf
	strHtml = strHtml & "		.grid2 .grid-item:nth-child(6n+3) em:first-child {background-color:transparent;color:#b6c;}" & vbCrlf
	strHtml = strHtml & "		.grid2 .grid-item:nth-child(6n+4) em:first-child {background-color:transparent;color:#ca4;}" & vbCrlf
	strHtml = strHtml & "		.grid2 .grid-item:nth-child(6n+5) em:first-child {background-color:transparent;color:#5b6;}" & vbCrlf

	strHtml = strHtml & "		.hr-rows .row-item {width:50%;font-size:1.1rem;line-height: 2.7rem;border-right: 1px solid #eee;padding-left: 20px;}" & vbCrlf
	strHtml = strHtml & "		.hr-rows .row-item:last-child {border-right:0;}" & vbCrlf
	strHtml = strHtml & "		.hr-rows .row-item i {position: relative;top:6px;color:#29f;padding-right:10px;font-size:1.8rem}" & vbCrlf
	strHtml = strHtml & "		.infobar {box-sizing: border-box;padding:10px 0;}" & vbCrlf
	strHtml = strHtml & "		.infobar .row-info {width:33%;text-align:center;border-right:1px solid #ddd;}" & vbCrlf
	strHtml = strHtml & "		.infobar .row-info:last-child {border-right:0;}" & vbCrlf
	strHtml = strHtml & "		.infobar .row-info dt b {color:#800;padding-right:3px} .infobar .row-info dd {color:#999;}" & vbCrlf
	strHtml = strHtml & "		.tit2 {line-height: 2.7rem;font-size:1.2rem}" & vbCrlf
	strHtml = strHtml & "		.tit2 em {margin:0 10px;width:1.6rem;height:1.6rem;line-height:1.6rem;text-align:center;background-color:#35b;border-radius:100%;color:#fff}" & vbCrlf
	strHtml = strHtml & "		.myAuth {justify-content:center;}" & vbCrlf
	strHtml = strHtml & "		.myAuth em {width:30%;text-align:center;background-color:#2ad;color:#fff;line-height:2rem;border-radius:3px;margin:5px;font-size:0.9rem}" & vbCrlf
	strHtml = strHtml & "		.hr-permit-tips {position:fixed;top:0;left:0;bottom:0;right:0;background-color:rgba(0, 0, 0, 0.8);z-index:1000}" & vbCrlf
	strHtml = strHtml & "		.tipsBox {position:relative;top:20%;left:20px;right:20px;z-index:1001;color:#fff;} .tipsBox dt {font-size:3.5rem;color:#f30;padding-right:10px;}" & vbCrlf
	strHtml = strHtml & "		.tipsBox dd {flex-grow:2;font-size:1rem;} .tipsBox dd b {font-weight: normal;color:#ccc;font-size:0.8rem;}" & vbCrlf
	strHtml = strHtml & "		.back-btn {position:relative;top:30%;text-align:center;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""UserBody"">" & vbCrlf
	Response.Write "	<div class=""user-photo"" style=""background-image:url(" & HeadFace & ");""></div>" & vbCrlf
	Response.Write "	<div class=""user-nick""><p class=""nick"" id=""J_myNick"">" & UserYGXM & "</p><p class=""level level5""></p></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<section class=""user-behavior userBehavior"" id=""userBehavior2"">"
	Response.Write "	<ul data-spm=""1"">" & vbCrlf
	Response.Write "		<li><a href=""javascript:void(0);""><p>部门</p><p>" & tKSMC & "</p> </a> </li>" & vbCrlf
	Response.Write "		<li><a href=""javascript:void(0);""><p>工号</p><p>" & UserYGDM & "</p> </a> </li>" & vbCrlf
	Response.Write "		<li><a href=""javascript:void(0);""><p>级别</p><p>" & arrManageRank(UserRank) & "</p> </a> </li> " & vbCrlf
	Response.Write "	</ul>" & vbCrlf
	Response.Write "</section>" & vbCrlf
	Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	Dim rsNum, numTeacher, numItem, numRecord, numMidiApplay
	Set rsNum = Conn.Execute("Select Count(TeacherID) From HR_Teacher")
		numTeacher = HR_Clng(rsNum(0))
	Set rsNum = Nothing
	sql = "" : numItem = 0
	Set rsNum = Conn.Execute("Select ClassID From HR_Class Where ModuleID=1001 And Child=0")
		If Not(rsNum.BOF And rsNum.EOF) Then
			i = 0
			Do While Not rsNum.EOF
				If i > 0 Then sql = sql & " union all "
				sql = sql & "select count(1) as CNT From HR_Sheet_" & rsNum("ClassID") & " Where scYear=" & DefYear
				rsNum.MoveNext
				i = i + 1
			Loop
			numItem = i
		End If
	Set rsNum = Nothing

	Set rsNum = Conn.Execute("Select Count(0) From HR_Message Where MsgType=1")		'修改申请统计
		numMidiApplay = HR_Clng(rsNum(0))
	Set rsNum = Nothing

	sql="select sum(CNT) from (" & sql & ") as nTab"
	Set rsNum = Conn.Execute(sql)
		numRecord = HR_Clng(rsNum(0))
	Set rsNum = Nothing

	Response.Write "<div class=""yearbar"">学年：<b>" & DefYear-1 & "-" & DefYear & "</b></div>" & vbCrlf
	Response.Write "<div class=""hr-rows infobar"">" & vbCrlf
	Response.Write "	<dl class=""row-info""><dt><b>" & numTeacher & "</b>名</dt><dd>教师数</dd></dl>" & vbCrlf
	Response.Write "	<dl class=""row-info""><dt><b>" & numItem & "</b>项</dt><dd>考核项目</dd></dl>" & vbCrlf
	Response.Write "	<dl class=""row-info""><dt><b>" & numRecord & "</b>条</dt><dd>业绩总数</dd></dl>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "<div class=""hr-grids"">" & vbCrlf
	Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "ManageNotice/Index.html?Type=1""><em><i class=""hr-icon"">&#xe9b4;</i></em><em>发布通知</em></a>" & vbCrlf
	Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "ManageTeacher/Index.html""><em><i class=""hr-icon"">&#xf2f1;</i></em><em>教师查询</em></a>" & vbCrlf
	Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "ManageDepart/Index.html""><em><i class=""hr-icon"">&#xf2f6;</i></em><em>科室查询</em></a>" & vbCrlf
	Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "ManageCourse/Index.html""><em><i class=""hr-icon"">&#xec8c;</i></em><em>业绩管理</em></a>" & vbCrlf
	Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "ManageReport/Index.html""><em><i class=""hr-icon"">&#xeb16;</i></em><em>业绩报表</em></a>" & vbCrlf
	Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "ManageGoback/Index.html""><em><i class=""hr-icon"">&#xf338;</i></em><em>业绩退回</em></a>" & vbCrlf
	Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "ManagePass/Index.html""><em><i class=""hr-icon"">&#xef82;</i></em><em>业绩审核</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "ManageDataTemp/Index.html""><em><i class=""hr-icon"">&#xef82;</i></em><em>数据模板</em></a>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	'Response.Write "<div class=""hr-rows"">" & vbCrlf
	'Response.Write "	<a class=""row-item"" href=""" & ParmPath & "Manage/Course.html?Type=1""><i class=""hr-icon"">&#xe1b2;</i>基础性教学</a>" & vbCrlf
	'Response.Write "	<a class=""row-item"" href=""" & ParmPath & "Manage/Course.html?Type=2""><i class=""hr-icon"">&#xe8a3;</i>激励性教学</a>" & vbCrlf
	'Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells hr-cell"">" & vbCrlf
	Dim numSwapNopass, numSwapTotal
	Set rsNum = Conn.Execute("Select Count(0) From HR_Swap Where Process=0")		'调换课未审
		numSwapNopass = HR_Clng(rsNum(0))
	Set rsNum = Nothing
	Set rsNum = Conn.Execute("Select Count(0) From HR_Swap Where Process>0")		'调换课已审
		numSwapTotal = HR_Clng(rsNum(0))
	Set rsNum = Nothing

	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageSwap/Index.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><i class=""hr-icon"">&#xe1b2;</i></div><div class=""weui-cell__bd weui-cell_primary"" data-id=""""><p>调换课审核</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">待审<b style=""color:#f30"">" & numSwapNopass & "</b>/已审<b style=""color:#090"">" & numSwapTotal & "</b>条</div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageModify/Index.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd"" style=""color:#a3c""><i class=""hr-icon"">&#xe9b3;</i></div><div class=""weui-cell__bd weui-cell_primary"" data-id=""""><p>修改申请</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><b style=""color:#f30"">" & HR_CLng(numMidiApplay) & "</b>条</div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Dim noPassCourse : sql = ""		'未审课程
	Set rs = Conn.Execute("Select ClassID From HR_Class Where ModuleID=1001 And Child=0")
		If Not(rs.BOF And rs.EOF) Then
			i = 0
			Do While Not rs.EOF
				If i > 0 Then sql = sql & " union all "
				sql = sql & "select count(1) as CNT From HR_Sheet_" & rs("ClassID") & " Where Passed=" & HR_False & " And scYear=" & DefYear
				rs.MoveNext
				i = i + 1
			Loop
		End If
	Set rs = Nothing

	sql="select sum(CNT) from (" & sql & ") as nTab"
	Set rs = Conn.Execute(sql)
		noPassCourse = HR_Clng(rs(0))
	Set rs = Nothing

	Dim noApplyCourse : sql = ""		'未确认课程
	Set rs = Conn.Execute("Select ClassID From HR_Class Where ModuleID=1001 And Child=0")
		If Not(rs.BOF And rs.EOF) Then
			i = 0
			Do While Not rs.EOF
				If i > 0 Then sql = sql & " union all "
				sql = sql & "select count(1) as CNT From HR_Sheet_" & rs("ClassID") & " Where State=0 And scYear=" & DefYear
				rs.MoveNext
				i = i + 1
			Loop
		End If
	Set rs = Nothing

	sql="select sum(CNT) from (" & sql & ") as nTab"
	Set rs = Conn.Execute(sql)
		noApplyCourse = HR_Clng(rs(0))
	Set rs = Nothing

	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManagePass/Index.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd"" style=""color:#da1""><i class=""hr-icon"">&#xe07f;</i></div><div class=""weui-cell__bd weui-cell_primary"" data-id=""""><p>未审业绩</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><b style=""color:#f30"">" & HR_CLng(noPassCourse) & "</b>条</div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageAffirm/Index.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd"" style=""color:#0af""><i class=""hr-icon"">&#xf35e;</i></div><div class=""weui-cell__bd weui-cell_primary"" data-id=""""><p>业绩确认</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><b style=""color:#f30"">" & HR_CLng(noApplyCourse) & "</b>/<b style=""color:#090"">" & HR_CLng(numRecord) - HR_CLng(noApplyCourse) & "</b>条</div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Dim iPropose, iEval, iEvalCEX
	Set rsTmp = Conn.Execute("Select Count(0) From HR_Propose")
		iPropose = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing
	Set rsTmp = Conn.Execute("Select Count(0) From HR_Evaluate")
		iEval = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing
	Set rsTmp = Conn.Execute("Select Count(0) From HR_EvaluateCEX")
		iEvalCEX = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing
	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManagePropose/Index.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd"" style=""color:#c61""><i class=""hr-icon"">&#xf298;</i></div><div class=""weui-cell__bd weui-cell_primary"" data-id=""""><p>意见建议</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><b style=""color:#f30"">" & iPropose & "</b>条</div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageEvaluate/TeachQuality.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd"" style=""color:#0b0""><i class=""hr-icon"">&#xe9dc;</i></div><div class=""weui-cell__bd weui-cell_primary"" data-id=""""><p>课堂教学质量评价</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><b style=""color:#f30"">" & iEval & "</b>条</div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageEvaluate/CEX.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd"" style=""color:#e47""><i class=""hr-icon"">&#xe9dc;</i></div><div class=""weui-cell__bd weui-cell_primary"" data-id=""""><p>mini-CEX<sup>plus</sup>记录</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""><b style=""color:#f30"">" & iEvalCEX & "</b>条</div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	'Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	'Response.Write "<div class=""hr-grids tit2""><em><i class=""hr-icon"">&#xf013;</i></em><tt>系统管理</tt></div>" & vbCrlf
	'Response.Write "<div class=""hr-grids grid2"">" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "ManageItem/Index.html?Type=1""><em><i class=""hr-icon"">&#xe9b4;</i></em><em>考核项管理</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xeb16;</i></em><em>考核系数</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xf393;</i></em><em>级别管理</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xe992;</i></em><em>等级管理</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xf2f1;</i></em><em>教师管理</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xf2f6;</i></em><em>科室管理</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xe9e8;</i></em><em>校[院]区</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xebf4;</i></em><em>授课点</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xebf4;</i></em><em>教学班级</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xeedb;</i></em><em>节次管理</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xea2c;</i></em><em>课程管理</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xf30b;</i></em><em>评分表管理</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xf30b;</i></em><em>学生类别</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xe429;</i></em><em>考核指标</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/ManageSetup/Switch.html""><em><i class=""hr-icon"">&#xf205;</i></em><em>业绩开关</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xe8d6;</i></em><em>等级导入</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/Manage/Course.html?Type=2""><em><i class=""hr-icon"">&#xea7d;</i></em><em>系统参数</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/ManageSetup/DataBack.html?Type=2""><em><i class=""hr-icon"">&#xf1c0;</i></em><em>数据备份</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/ManageSetup/API.html""><em><i class=""hr-icon"">&#xf37d;</i></em><em>接口参数</em></a>" & vbCrlf
	'Response.Write "	<a class=""grid-item"" href=""" & ParmPath & "/ManageSetup/Wechat.html""><em><i class=""hr-icon"">&#xec60;</i></em><em>微信参数</em></a>" & vbCrlf
	'Response.Write "</div>" & vbCrlf

	mgStuType = "<em>" & Replace(mgStuType, ",", "</em><em>") & "</em>"
	Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "<div class=""hr-grids tit2""><em style=""background-color:#a61""><i class=""hr-icon"">&#xef2f;</i></em><tt>我的权限</tt></div>" & vbCrlf
	Response.Write "<div class=""hr-grids myAuth""><em>" & arrManageRank(UserRank) & "</em>" & mgStuType & "</div>" & vbCrlf
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	'Response.Write GetManagePermit() & vbCrlf	'显示无权限

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub
%>