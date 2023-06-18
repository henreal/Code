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
Dim scriptCtrl : SiteTitle = "管理首页_" & SiteName
Dim jsonOBJ, strJson : strJson = GetWechatUserInfoQY(UserYGDM)
If Instr(strJson, "errcode") > 0 Then
	Set jsonOBJ = parseJSON(strJson)
		If jsonOBJ.errcode = 0 Then
			UserFace = Trim(jsonOBJ.avatar)
		End If
	Set jsonOBJ = Nothing
End If
If HR_IsNull(UserFace) Then UserFace = InstallDir & "Static/images/nopic.png"


If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "Start" Call IndexBody()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()		'非frameset框架
	strHtml = getPageHead(1)
	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.frame_Left {width: 170px;} .frame_Main {left:170px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.statBox {display:none;width:0;height:0;overflow:hidden;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-side {position:initial}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-side-scroll {width:100%}" & vbCrlf
	tmpHtml = tmpHtml & "		.MenuTitle .ShrinkMenu i {color:rgba(255,255,255,0.7);cursor: pointer;}" & vbCrlf
	tmpHtml = tmpHtml & "		.LeftSwitch {position:absolute;width:0;height:39px;line-height:39px;background-color:#003e65;color:#fff;z-index:150;font-size:1.5rem;cursor: pointer;overflow: hidden;opacity:0}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	tmpHtml = tmpHtml & "	<meta name=""keywords"" content=""公司网站管理后台,恒锐公司网站"" />" & vbCrlf
	tmpHtml = tmpHtml & "	<meta name=""description"" content=""恒锐网络公司网站管理后台，仅适用于门户网站、业绩考核、员工考勤、企业微信管理、微信公众号管理模块"" />"

	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")

	tmpHtml = "<header class=""frame_Head"">" & vbCrlf
	tmpHtml = tmpHtml & "	<dl class=""hr-rows frame_top"">" & vbCrlf
	tmpHtml = tmpHtml & "		<dt></dt><dd class=""hr-row-fill navTitle""></dd>" & vbCrlf
	tmpHtml = tmpHtml & "		<dd class=""top_nav"">" & vbCrlf
	tmpHtml = tmpHtml & "			<ul class=""hr-rows nav"">" & vbCrlf
	tmpHtml = tmpHtml & "				<li><a href=""" & InstallDir & """ target=""_blank"" class=""hr-hover selected""><img src=""" & InstallDir & "Static/admin/icon09.png"" title=""教师首页"" /><h3>教师首页</h3></a></li>" & vbCrlf
	'tmpHtml = tmpHtml & "				<li><a href=""" & ParmPath & "User/ModiPass.html"" target=""rightFrame"" class=""hr-hover""><img src=""" & InstallDir & "Static/admin/icon06.png"" title=""修改密码"" /><h3>修改密码</h3></a></li>" & vbCrlf
	tmpHtml = tmpHtml & "				<li><a href=""" & ParmPath & "MyCenter/Message.html"" target=""rightFrame"" class=""hr-hover""><img src=""" & InstallDir & "Static/admin/icon10.png"" title=""我的消息"" /><h3>我的消息</h3></a></li>" & vbCrlf
	tmpHtml = tmpHtml & "				<li><a href=""" & ParmPath & "Help/Index.html"" target=""rightFrame"" class=""hr-hover""><img src=""" & InstallDir & "Static/admin/ico05.png"" title=""使用手册"" /><h3>使用手册</h3></a></li>" & vbCrlf
	tmpHtml = tmpHtml & "				<li><a href=""" & InstallDir & "Desktop/Login.html?Logout=True"" target=""_parent"" class=""hr-hover""><img src=""" & InstallDir & "Static/admin/icon08.png"" title=""退出系统"" /><h3>退出系统</h3></a></li>" & vbCrlf
	tmpHtml = tmpHtml & "			</ul>" & vbCrlf
	tmpHtml = tmpHtml & "		</dd>" & vbCrlf
	tmpHtml = tmpHtml & "		<dd class=""search""><input type=""text"" name=""soWord"" class=""inputtxt"" id=""soWord"" placeholder=""请输入关键字""></dd>" & vbCrlf
	tmpHtml = tmpHtml & "		<dd class=""soBtn""><button name=""soPost"" class=""btn1"" id=""soPost"" type=""button""><i class=""hr-icon"">&#xf35f;</i></button></dd>" & vbCrlf
	tmpHtml = tmpHtml & "		<dd class=""userbar""><em style=""background-image:url(" & UserFace & ");""></em><tt>" & UserYGXM & "</tt></dd>" & vbCrlf
	tmpHtml = tmpHtml & "		<dd class=""more""><i class=""hr-icon"">&#xf351;</i>帮助</dd>" & vbCrlf
	tmpHtml = tmpHtml & "	</dl>" & vbCrlf
	tmpHtml = tmpHtml & "</header>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""frame_Left"" title=""左栏菜单"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""MenuBox"" id=""menubox"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""hr-rows MenuTitle""><em><i class=""hr-icon"">&#xe1bd;</i>导航菜单</em><em class=""ShrinkMenu"" title=""关闭左栏""><i class=""hr-icon"">&#xed10;</i></em></div>" & vbCrlf
	tmpHtml = tmpHtml & ShowLeftBody()
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""frame_Main"">" & vbCrlf
	tmpHtml = tmpHtml & "	<iframe src=""" & ParmPath & "Index/Start.html"" class=""MainBox"" name=""rightFrame"" id=""rightFrame"" title=""rightFrame""></iframe>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""LeftSwitch""><div class=""switch"" title=""展开左栏""><i class=""hr-icon"">&#xed11;</i></div></div>" & vbCrlf

	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".ShrinkMenu"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$("".frame_Left"").animate({width:0}); $("".frame_Main"").animate({left:0});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".LeftSwitch"").animate({width:""25px"",opacity:""0.8""});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$("".LeftSwitch"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$("".frame_Left"").animate({width:""170px""}); $("".frame_Main"").animate({left:""170px""});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".LeftSwitch"").animate({width:""0px"",opacity:""0""});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = strHtml & tmpHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", "")
	strHtml = Replace(strHtml, "[@ErrMSG]", "模板代码为空")

	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub

Sub IndexBody()
	Dim arrManageRank : arrManageRank = Split(XmlText("Config", "ManageRank", ""), "|")
	Dim mgStuType : mgStuType = GetTypeName("HR_User", "StuType", "UserID", UserID)

	Dim TotalNum, meUploadNum, ItemNum, noPassNum, noAffirmNum, TeacherNum
	Set rs = Conn.Execute("Select Count(TeacherID) From HR_Teacher")
		TeacherNum = HR_Clng(rs(0))
	Set rs = Nothing
	sql = ""
	Set rs = Conn.Execute("Select ClassID From HR_Class Where ModuleID=1001 And Child=0")
		If Not(rs.BOF And rs.EOF) Then
			i = 0
			Do While Not rs.EOF
				If i > 0 Then sql = sql & " union all "
				sql = sql & "select count(1) as CNT From HR_Sheet_" & rs("ClassID") & " Where scYear=" & DefYear
				rs.MoveNext
				i = i + 1
			Loop
			ItemNum = i
		End If
	Set rs = Nothing

	sql="select sum(CNT) from (" & sql & ") as nTab"
	Set rs = Conn.Execute(sql)
		TotalNum = HR_Clng(rs(0))
	Set rs = Nothing


	strHtml = getPageHead(1)
	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.home-body {box-sizing:border-box;background-color:#f1f1f1;padding:15px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.manager-face {padding-right:20px;text-align:center;flex-shrink:0} .manager-face em {width:70px;height:70px;background:#fff center no-repeat;background-size:100%;border-radius:8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.manager-info {flex-grow:2;} .sys-info dt {width:120px;text-align:right;color:#999}" & vbCrlf
	tmpHtml = tmpHtml & "		.manager-info dt {width:100px;text-align:right;color:#999}" & vbCrlf
	tmpHtml = tmpHtml & "		" & vbCrlf
	tmpHtml = tmpHtml & "		" & vbCrlf
	tmpHtml = tmpHtml & "		" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	strHtml = strHtml & getFrameNav(1)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
	
	Dim LastLoginTime, LoginIP, LastLoginIP
	Set rsTmp = Conn.Execute("Select Top 1 * From HR_Teacher Where YGDM='" & UserYGDM & "'")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			LastLoginTime = rsTmp("LoginTime")
			LastLoginIP = rsTmp("LoginIP")
		End If
	Set rsTmp = Nothing

	Response.Write "<div class=""home-body"">" & vbCrlf
	Response.Write "	<div class=""layui-row layui-col-space15 home-box1"">" & vbCrlf
	Response.Write "		<div class=""layui-col-md9"">" & vbCrlf
	Response.Write "			<div class=""layui-card"">" & vbCrlf
	Response.Write "				<div class=""layui-card-header"">管理员信息</div>" & vbCrlf
	Response.Write "				<div class=""layui-card-body hr-rows hr-item-top"">" & vbCrlf
	Response.Write "					<div class=""manager-face""><em style=""background-image:url(" & UserFace & ");""></em></div>" & vbCrlf
	Response.Write "					<div class=""manager-info"">" & vbCrlf
	Response.Write "						<dl class=""hr-grids""><dt>姓名：</dt><dd>" & UserYGXM & " [管理员ID：" & UserID & "]</dd></dl>" & vbCrlf
	Response.Write "						<dl class=""hr-grids""><dt>工号：</dt><dd>" & UserYGDM & "</dd></dl>" & vbCrlf
	Response.Write "						<dl class=""hr-grids""><dt>管理级别：</dt><dd>" & arrManageRank(UserRank) & "</dd></dl>" & vbCrlf
	Response.Write "						<dl class=""hr-grids""><dt>管理权限：</dt><dd>" & mgStuType & "</dd></dl>" & vbCrlf
	Response.Write "					</div>" & vbCrlf
	Response.Write "					<div class=""sys-info"">" & vbCrlf
	Response.Write "						<dl class=""hr-grids""><dt>当前版本：</dt><dd>" & XmlText("Config", "Ver", "") & "</dd></dl>" & vbCrlf
	Response.Write "						<dl class=""hr-grids""><dt>上次登陆：</dt><dd>" & FormatDate(LastLoginTime, 10) & "</dd></dl>" & vbCrlf
	Response.Write "						<dl class=""hr-grids""><dt>登陆IP：</dt><dd>" & LastLoginIP & "</dd></dl>" & vbCrlf
	Response.Write "					</div>" & vbCrlf
	Response.Write "				</div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf

	Response.Write "			<div class=""layui-row layui-col-space15"">" & vbCrlf
	Response.Write "				<div class=""layui-col-md6"">" & vbCrlf
	Response.Write "					<div class=""layui-card"">" & vbCrlf
	Response.Write "						<div class=""layui-card-header"">当前学年</div>" & vbCrlf
	Response.Write "						<div class=""layui-card-body"">" & vbCrlf
	Response.Write "							<dl class=""hr-grids""><dt>学年：</dt><dd>" & DefYear-1 & "-" & DefYear & "</dd></dl>" & vbCrlf
	Response.Write "						</div>"
	Response.Write "					</div>"
	Response.Write "				</div>"
	Response.Write "				<div class=""layui-col-md6"">" & vbCrlf
	Response.Write "					<div class=""layui-card"">" & vbCrlf
	Response.Write "						<div class=""layui-card-header"">系统信息</div>" & vbCrlf
	Response.Write "						<div class=""layui-card-body sys-info"">" & vbCrlf
	Response.Write "							<dl class=""hr-grids""><dt>总业绩记录：</dt><dd>" & TotalNum & " 条</dd></dl>" & vbCrlf
	Response.Write "							<dl class=""hr-grids""><dt>我上传的数据：</dt><dd>" & meUploadNum & " 条</dd></dl>" & vbCrlf
	Response.Write "							<dl class=""hr-grids""><dt>考核项目：</dt><dd>" & ItemNum & " 项</dd></dl>" & vbCrlf
	Response.Write "							<dl class=""hr-grids""><dt>教师数：</dt><dd>" & TeacherNum & " 名</dd></dl>" & vbCrlf
	Response.Write "							<dl class=""hr-grids""><dt>未审核：</dt><dd>" & noPassNum & " 条</dd></dl>" & vbCrlf
	Response.Write "							<dl class=""hr-grids""><dt>未确认：</dt><dd>" & noAffirmNum & " 条</dd></dl>" & vbCrlf
	Response.Write "						</div>" & vbCrlf
	Response.Write "					</div>" & vbCrlf
	Response.Write "				</div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-col-md3"">" & vbCrlf
	Response.Write "			<div class=""layui-card"">" & vbCrlf
	Response.Write "				<div class=""layui-card-header"">最新消息</div>" & vbCrlf
	Response.Write "				<div class=""layui-card-body"">" & vbCrlf

	Response.Write "				</div>"
	Response.Write "			</div>"
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""element"", ""form""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var element = layui.element, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		$("".BandWeChatQY"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			layer.open({type:2, id:""bindWin"",content:""" & ParmPath & "User/BindWeChatQY.html"",title:[""扫描下方二维码"",""font-size:16""],area:[""340px"", ""500px""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	strHtml = Replace(strHtml, "[@ErrMSG]", "模板代码为空")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub

Function ShowLeftBody()
	Dim tmpStr
	tmpStr = "<div class=""layui-side layui-side-menu"">" & vbCrlf
	tmpStr = tmpStr & "	<div class=""layui-side-scroll"">" & vbCrlf
	tmpStr = tmpStr & "		<ul class=""layui-nav layui-nav-tree hr-nav-skin"" lay-shrink=""all"" id=""right-menu"" lay-filter=""right-side-menu"">" & vbCrlf
	tmpStr = tmpStr & "			<li data-name=""home"" class=""layui-nav-item layui-nav-itemed"">" & vbCrlf
	tmpStr = tmpStr & "				<a href=""javascript:void(0);"" class=""hr-navmenu-main""><i class=""hr-icon hr-icon-top"">&#xe1b2;</i><u>基础性教学</u></a>" & vbCrlf
	tmpStr = tmpStr & "				<dl class=""layui-nav-child"">" & vbCrlf
	tmpStr = tmpStr & "					" & ShowNavMenu(1)
	tmpStr = tmpStr & "				</dl>" & vbCrlf
	tmpStr = tmpStr & "			</li>" & vbCrlf
	tmpStr = tmpStr & "			<li data-name=""home"" class=""layui-nav-item"">" & vbCrlf
	tmpStr = tmpStr & "				<a href=""javascript:void(0);"" class=""hr-navmenu-main""><i class=""hr-icon hr-icon-top"">&#xe8a3;</i><u>激励性教学</u></a>" & vbCrlf
	tmpStr = tmpStr & "				<dl class=""layui-nav-child"">" & vbCrlf
	tmpStr = tmpStr & "					" & ShowNavMenu(2)
	tmpStr = tmpStr & "				</dl>" & vbCrlf
	tmpStr = tmpStr & "			</li>" & vbCrlf

	tmpStr = tmpStr & "			<li data-name=""home"" class=""layui-nav-item"">" & vbCrlf
	tmpStr = tmpStr & "				<a href=""javascript:void(0);"" class=""hr-navmenu-main""><i class=""hr-icon hr-icon-top"">&#xeb16;</i><u>查看业绩</u></a>" & vbCrlf
	tmpStr = tmpStr & "				<dl class=""layui-nav-child"">" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & InstallDir & "Desktop/Achieve/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>业绩报表</a></dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "Achieve/Collect.html?teacher=" & UserYGDM & """ target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>学时汇总</a></dd>" & vbCrlf
	tmpStr = tmpStr & "				</dl>" & vbCrlf
	tmpStr = tmpStr & "			</li>" & vbCrlf

	tmpStr = tmpStr & "			<li data-name=""home"" class=""layui-nav-item"">" & vbCrlf
	tmpStr = tmpStr & "				<a href=""javascript:void(0);"" class=""hr-navmenu-main""><i class=""hr-icon hr-icon-top"">&#xee36;</i><u>日常工作</u></a>" & vbCrlf
	tmpStr = tmpStr & "				<dl class=""layui-nav-child"">" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "Notice/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>通知公告</a></dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "Swap/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>换课申请</a></dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "Substitute/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>代课申请</a></dd>" & vbCrlf	
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "Evaluate/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>课堂教学质量评价</a></dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "EvaluateCEX/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>形成性评价</a></dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "Remind/ListLog.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>上课提醒记录</a></dd>" & vbCrlf
	tmpStr = tmpStr & "				</dl>" & vbCrlf
	tmpStr = tmpStr & "			</li>" & vbCrlf

	tmpStr = tmpStr & "			<li data-name=""home"" class=""layui-nav-item"">" & vbCrlf
	tmpStr = tmpStr & "				<a href=""javascript:void(0);"" class=""hr-navmenu-main""><i class=""hr-icon hr-icon-top"">&#xf085;</i><u>系统管理</u></a>" & vbCrlf
	tmpStr = tmpStr & "				<dl class=""layui-nav-child"">" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""navmenu"">" & vbCrlf
	tmpStr = tmpStr & "						<a class=""hr-navmenu-main"" href=""javascript:void(0);"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>人员管理</a>" & vbCrlf
	tmpStr = tmpStr & "						<dl class=""layui-nav-child"">" & vbCrlf
	tmpStr = tmpStr & "							<dd data-name=""child""><a href=""" & ParmPath & "Teacher/List.html"" target=""rightFrame""><i class=""hr-icon"">&#xf328;</i>教师管理</a></dd>" & vbCrlf
	tmpStr = tmpStr & "							<dd data-name=""child""><a href=""" & ParmPath & "Department/List.html"" target=""rightFrame""><i class=""hr-icon"">&#xf328;</i>科室管理</a></dd>" & vbCrlf
	tmpStr = tmpStr & "							<dd data-name=""child""><a href=""" & ParmPath & "SwapPasser/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf328;</i>调换课审核员</a></dd>" & vbCrlf
	tmpStr = tmpStr & "							<dd data-name=""child""><a href=""" & ParmPath & "User/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf328;</i>管理员管理</a></dd>" & vbCrlf
	tmpStr = tmpStr & "						</dl>" & vbCrlf
	tmpStr = tmpStr & "					</dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""navmenu"">" & vbCrlf
	tmpStr = tmpStr & "						<a class=""hr-navmenu-main"" href=""javascript:void(0);"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>考核项管理</a>" & vbCrlf
	tmpStr = tmpStr & "						<dl class=""layui-nav-child"">" & vbCrlf
	tmpStr = tmpStr & "							<dd data-name=""child""><a href=""" & ParmPath & "ExamItems/Index.html?Type=1"" target=""rightFrame""><i class=""hr-icon"">&#xf328;</i>A类考核项目</a></dd>" & vbCrlf
	tmpStr = tmpStr & "							<dd data-name=""child""><a href=""" & ParmPath & "ExamItems/Index.html?Type=2"" target=""rightFrame""><i class=""hr-icon"">&#xf328;</i>B类考核项目</a></dd>" & vbCrlf
	tmpStr = tmpStr & "						</dl>" & vbCrlf
	tmpStr = tmpStr & "					</dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""navmenu"">" & vbCrlf
	tmpStr = tmpStr & "						<a class=""hr-navmenu-main"" href=""javascript:void(0);"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>系统配置</a>" & vbCrlf
	tmpStr = tmpStr & "						<dl class=""layui-nav-child"">" & vbCrlf
	tmpStr = tmpStr & "							<dd data-name=""child""><a href=""" & ParmPath & "Setup/Period.html"" target=""rightFrame""><i class=""hr-icon"">&#xf328;</i>校区节次管理</a></dd>" & vbCrlf
	tmpStr = tmpStr & "							<dd data-name=""child""><a href=""" & ParmPath & "Setup/Course.html?Type=2"" target=""rightFrame""><i class=""hr-icon"">&#xf328;</i>课程管理</a></dd>" & vbCrlf
	tmpStr = tmpStr & "							<dd data-name=""child""><a href=""" & ParmPath & "Setup/TeachClass.html"" target=""rightFrame""><i class=""hr-icon"">&#xf328;</i>授课对象管理</a></dd>" & vbCrlf
	tmpStr = tmpStr & "							<dd data-name=""child""><a href=""" & ParmPath & "Setup/ClassRoom.html?Type=2"" target=""rightFrame""><i class=""hr-icon"">&#xf328;</i>授课教室管理</a></dd>" & vbCrlf
	tmpStr = tmpStr & "						</dl>" & vbCrlf
	tmpStr = tmpStr & "					</dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "Setup/SetupSwitch.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>业绩开关</a></dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "DataModel/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>数据模型管理</a></dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "Setup/BackData.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>数据备份</a></dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "DataDict/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>数据字典</a></dd>" & vbCrlf
	'tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "WechatQY/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>企业微信管理</a></dd>" & vbCrlf
	tmpStr = tmpStr & "					<dd data-name=""console""><a href=""" & ParmPath & "Interface/Index.html"" target=""rightFrame""><i class=""hr-icon"">&#xf31d;</i>接口管理</a></dd>" & vbCrlf
	tmpStr = tmpStr & "				</dl>" & vbCrlf
	tmpStr = tmpStr & "			</li>" & vbCrlf
	tmpStr = tmpStr & "		</ul>" & vbCrlf
	tmpStr = tmpStr & "	</div>" & vbCrlf
	tmpStr = tmpStr & "</div>" & vbCrlf
	ShowLeftBody = tmpStr
End Function

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
					strMenu = strMenu & "href=""" & ParmPath & "Course.html?ItemID=" & rsMenu("ClassID") & """ target=""rightFrame"">"
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