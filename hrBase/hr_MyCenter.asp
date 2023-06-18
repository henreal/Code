<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incMD5.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./incKPI.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")

Dim Page_Title : Page_Title = "个人中心"
Dim SubButTxt : SubButTxt = "个人信息"
Dim arrSex : arrSex = Split(XmlText("Config", "Sex", ""), "|")

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))				'Get head template code.
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))				'Get foot template code.
Dim strNavPath : strNavPath = ReplaceCommonLabel(getFrameNav(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index", "List" Call MainBody()
	Case "ShowInfo" Call ShowInfo()
	Case "Password" Call Password()
	Case "SavePass" Call SavePass()
	Case "Message" Call Message()
	Case "UpdateMessage" Call UpdateMessage()
	Case "BackApply" Call BackApply()
	Case "SendBackApply" Call SendBackApply()
	Case "TransferApply" Call TransferApply()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	strHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />" & vbCrlf
	tmpHtml = tmpHtml & "	<style type=""text/css"">" & vbCrlf
	'tmpHtml = tmpHtml & "		.iframe-nav .navBtn .navLayer {font-size: 16px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@PageTitle]", Page_Title)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var layer = layui.layer, element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "			layer.config({skin:""layer-hr""});var loadInit = layer.load();" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf

	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	strHtml = "<a href=""" & ParmPath & "Achieve/List.html"">" & Page_Title & "</a><a><cite>" & SubButTxt & "</cite></a>"
	strNavPath = Replace(strNavPath, "[@ModulePath]", strHtml)
	Response.Write strNavPath

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf

	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){" & vbCrlf
	'strHtml = strHtml & "		$("".navLayer"").html(""<i class=\""hr-icon\"">&#xf067;</i>添加项目"")" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub ShowInfo()
	Dim tmpHtml : SubButTxt = "添加" : ErrMsg = ""
	Dim tYGDM : tYGDM = GetTypeName("HR_User", "TeacherNum", "UserID", UserID)
	If UserYGDM <> "" Then tYGDM = UserYGDM

	Dim FieldLen, arrFieldName, arrFieldValue
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select Top 1 a.*,b.KSMC From HR_Teacher a Inner Join HR_Department b On a.KSDM=b.KSDM Where a.YGDM='" & tYGDM & "'"), Conn, 1, 1
			Redim arrFieldName(rsTmp.Fields.Count-1)
			Redim arrFieldValue(rsTmp.Fields.Count-1)
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				For i = 0 To rsTmp.Fields.Count-1
					arrFieldValue(i) = rsTmp.Fields(i).Value
					arrFieldName(i) = rsTmp.Fields(i).Name
				Next
			End If
			ChkTeacherKPI(tYGDM)	'添加员工信息至业绩表
			UpdateTeacherTotalKPI(tYGDM)	'更新员工总计数据
		rsTmp.Close
	Set rsTmp = Nothing
	tmpHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />" & vbCrlf
	tmpHtml = tmpHtml & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding: 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-color-true {color:#080;} .hr-color-false {color:#F30;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", "")
	strHeadHtml = Replace(strHeadHtml, "[@Head_style]", tmpHtml)

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		$(document).ready(function(){ });" & vbCrlf
	tmpHtml = tmpHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var layer = layui.layer, element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "			layer.config({skin:""layer-hr""});var loadInit = layer.load();" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Head_script]", tmpHtml)
	Response.Write strHeadHtml
	strHtml = "<a><cite>个人信息</cite></a>"
	strNavPath = Replace(strNavPath, "[@Module_Path]", strHtml)
	Response.Write strNavPath
	Response.Write "<div class=""hr-body-w800 hr-shrink-x10"">" & vbCrlf

	strHtml = "<fieldset class=""layui-elem-field layui-field-title""><legend>我的 个人信息</legend>" & vbCrlf
	strHtml = strHtml & "	<div class=""layui-form layer-hr-box"">" & vbCrlf
	strHtml = strHtml & "		<table class=""layui-table"">" & vbCrlf
	strHtml = strHtml & "			<colgroup><col width=""120""><col><col width=""120""><col></colgroup>" & vbCrlf
	strHtml = strHtml & "			<tbody>" & vbCrlf
	If UserID>0 Then
		strHtml = strHtml & "				<tr><td style=""text-align:right;"">帐　号：</td><td>" & UserName & "</td><td style=""text-align:right;"">ID：</td><td>" & UserID & "</td></tr>" & vbCrlf
		strHtml = strHtml & "				<tr><td style=""text-align:right;"">姓　名：</td><td>" & TrueName & "</td><td style=""text-align:right;"">姓　名：</td><td>" & UserID & "</td></tr>" & vbCrlf
		strHtml = strHtml & "				<tr><td style=""text-align:right;"">提示：</td><td colspan=""3"">系统管理员暂无个人信息</td></tr>" & vbCrlf
	Else
		strHtml = strHtml & "				<tr><td style=""text-align:right;"">工　号：</td><td>" & arrFieldValue(3) & "</td><td style=""text-align:right;"">姓　名：</td><td>" & arrFieldValue(4) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "				<tr><td style=""text-align:right;"">性　别：</td><td>" & Trim(arrFieldValue(6)) & "</td><td style=""text-align:right;"">状　态：</td><td>" & arrFieldValue(7) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "				<tr><td style=""text-align:right;"">职　称：</td><td>" & arrFieldValue(12) & "</td><td style=""text-align:right;"">职　务：</td><td>" & arrFieldValue(14) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "				<tr><td style=""text-align:right;"">科　室：</td><td colspan=""3"">" & Trim(arrFieldValue(9)) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "				<tr><td colspan=""4""></td></tr>" & vbCrlf
		strHtml = strHtml & "				<tr><td style=""text-align:right;"">总　分：</td><td id=""totalScore"">0</td><td style=""text-align:right;"">排　序：</td><td><span id=""Grade""></span></td></tr>" & vbCrlf
	End If
	strHtml = strHtml & "			</tbody>" & vbCrlf
	strHtml = strHtml & "		</table>" & vbCrlf
	strHtml = strHtml & "	</div>" & vbCrlf
	strHtml = strHtml & "	<div class=""layer-hr_searchBox layui-form"">" & vbCrlf
	If HR_Clng(UserYGDM) > 0 Then
		strHtml = strHtml & "		<div class=""layui-input-block"">" & vbCrlf
		strHtml = strHtml & "			<button class=""layui-btn layui-btn-normal"" data-type=""myAchieve"" id=""myAchieve"" title=""我的业绩""><i class=""hr-icon"">&#xecb8;</i>我的业绩</button>" & vbCrlf
		strHtml = strHtml & "		</div>" & vbCrlf
	End If
	strHtml = strHtml & "	</div>" & vbCrlf
	strHtml = strHtml & "</fieldset>" & vbCrlf
	strHtml = strHtml & "</div>" & vbCrlf
	Response.Write strHtml

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	If HR_Clng(UserYGDM) > 0 Then
		strHtml = strHtml & "		$.getJSON(""" & ParmPath & "Ajax/ExportData.html"", { tid:" & arrFieldValue(3) & "}, function(tjData){" & vbCrlf
		strHtml = strHtml & "			$(""#totalScore"").text(tjData[0].totalScore);" & vbCrlf
		strHtml = strHtml & "			$(""#Grade"").html(tjData[0].tid);" & vbCrlf
		strHtml = strHtml & "		});" & vbCrlf
	End If
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "		$(""#myAchieve"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "			location.href=""" & ParmPath & "Tab.html?type=1&word=" & arrFieldValue(3) & """;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub Password()
	Dim tmpHtml : SubButTxt = "修改密码" : ErrMsg = ""
	Dim tYGDM : tYGDM = GetTypeName("HR_User", "TeacherNum", "UserID", UserID)
	If UserYGDM <> "" Then tYGDM = UserYGDM

	Dim FieldLen, arrFieldName, arrFieldValue
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select Top 1 * From HR_Teacher Where YGDM='" & tYGDM & "'"), Conn, 1, 1
			Redim arrFieldName(rsTmp.Fields.Count-1)
			Redim arrFieldValue(rsTmp.Fields.Count-1)
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				For i = 0 To rsTmp.Fields.Count-1
					arrFieldValue(i) = rsTmp.Fields(i).Value
					arrFieldName(i) = rsTmp.Fields(i).Name
				Next
			End If
		rsTmp.Close
	Set rsTmp = Nothing

	tmpHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />" & vbCrlf
	tmpHtml = tmpHtml & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding: 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-color-true {color:#080;} .hr-color-false {color:#F30;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", "")
	strHeadHtml = Replace(strHeadHtml, "[@Head_style]", tmpHtml)

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		$(document).ready(function(){ });" & vbCrlf
	tmpHtml = tmpHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var layer = layui.layer, element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "			layer.config({skin:""layer-hr""});var loadInit = layer.load();" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Head_script]", tmpHtml)
	Response.Write strHeadHtml
	strHtml = "<a><cite>个人信息</cite></a>"
	strNavPath = Replace(strNavPath, "[@Module_Path]", strHtml)
	Response.Write strNavPath
	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf

	strHtml = "<fieldset class=""layui-elem-field layui-field-title""><legend>密码修改</legend>" & vbCrlf
	strHtml = strHtml & "	<div class=""layui-form layer-hr-box"">" & vbCrlf

	strHtml = strHtml & "	<form class=""layui-form layui-form-pane"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	strHtml = strHtml & "		<div class=""layui-form-item""><label class=""layui-form-label"">原密码:</label>" & vbCrlf
	strHtml = strHtml & "			<div class=""layui-input-inline""><input type=""password"" name=""OldPass"" value="""" placeholder=""原密码不能为空"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>"
	strHtml = strHtml & "			<div class=""layui-form-mid layui-word-aux"">请输入您的原密码</div>"
	strHtml = strHtml & "		</div>" & vbCrlf
	strHtml = strHtml & "		<div class=""layui-form-item""><label class=""layui-form-label"">新密码:</label>" & vbCrlf
	strHtml = strHtml & "			<div class=""layui-input-inline""><input type=""password"" name=""NewPass"" value="""" placeholder=""新密码不能为空"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>"
	strHtml = strHtml & "			<div class=""layui-form-mid layui-word-aux"">密码由由英文半角的字母和数字构成</div>"
	strHtml = strHtml & "		</div>" & vbCrlf
	strHtml = strHtml & "		<div class=""layui-form-item""><label class=""layui-form-label"">密码确认:</label>" & vbCrlf
	strHtml = strHtml & "			<div class=""layui-input-inline""><input type=""password"" name=""NewPass2"" value="""" placeholder=""再次输入新密码"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>"
	strHtml = strHtml & "			<div class=""layui-form-mid layui-word-aux"">确认新密码</div>"
	strHtml = strHtml & "		</div>" & vbCrlf
	strHtml = strHtml & "		<input type=""hidden"" name=""Modify"" value=""True"">"
	strHtml = strHtml & "		<div class=""layui-form-item"">"
	strHtml = strHtml & "			<div class=""layui-input-block""><button class=""layui-btn"" lay-submit lay-filter=""SubPost"">修改密码</button><button type=""reset"" class=""layui-btn layui-btn-primary"">重置</button></div>"
	strHtml = strHtml & "		</div>" & vbCrlf
	strHtml = strHtml & "	</form>" & vbCrlf

	strHtml = strHtml & "	</div>" & vbCrlf
	strHtml = strHtml & "</fieldset>" & vbCrlf
	strHtml = strHtml & "</div>" & vbCrlf
	Response.Write strHtml

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	strHtml = strHtml & "		form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "MyCenter/SavePass.html"", $(""#EditForm"").serialize(), function(result){" & vbCrlf
	strHtml = strHtml & "				var reData = eval(""("" + result + "")"");" & vbCrlf
	strHtml = strHtml & "				if(reData.Return){" & vbCrlf
	strHtml = strHtml & "					layer.alert(reData.reMessge, {icon:1,title: ""修改结果""},function(layero, index){layer.closeAll();window.location.reload();});" & vbCrlf
	strHtml = strHtml & "				}else{" & vbCrlf
	strHtml = strHtml & "					layer.alert(reData.reMessge, {icon:2,title: ""错误提示""});" & vbCrlf
	strHtml = strHtml & "				}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@Foot_script]", strHtml)
	Response.Write strFootHtml
End Sub

Sub SavePass()
	Dim tmpJson, tmpData, rsGet, sqlGet, vCount, vMSG
	Dim OldPass, NewPass, NewPass2
	OldPass = Trim(Request("OldPass"))
	NewPass = Trim(Request("NewPass"))
	NewPass2 = Trim(Request("NewPass2"))
	vMSG = "{""Return"":false,""Err"":500,""reMessge"":""[@ErrMsg]"",""ReStr"":""操作失败！""}" : ErrMsg = ""
	If NewPass <> NewPass2 Then
		ErrMsg = "两次新密码不一致，请重新输入！"
		Call RecordFrontLog(1, ScriptName, "管理员ID：" & UserID & "，操作：修改密码", 0, Request.QueryString())
		Response.Write Replace(vMSG, "[@ErrMsg]", ErrMsg) :Exit Sub
	End If
	If UserID > 0 Then
		Set rsTmp = Server.CreateObject("ADODB.RecordSet")
			rsTmp.Open("Select * From HR_User Where UserID=" & UserID), Conn, 1, 3
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				If rsTmp("UserPass") <> MD5(OldPass, 16) Then
					ErrMsg = "原密码不正确，请重新输入！"
					Call RecordFrontLog(1, ScriptName, "管理员ID：" & UserID & "，操作：修改密码(" & ErrMsg & ")", 0, Request.QueryString())
					tmpJson = Replace(vMSG, "[@ErrMsg]", ErrMsg)
				Else
					rsTmp("UserPass") = MD5(NewPass, 16)
					rsTmp.Update
					ErrMsg = "修改密码成功，请用新密码登陆！"
					vMSG = Replace(vMSG, "false", "true") : vMSG = Replace(vMSG, "500", "0")
					tmpJson = Replace(vMSG, "[@ErrMsg]", ErrMsg)
					Call RecordFrontLog(1, ScriptName, "管理员ID：" & UserID & "，操作：修改密码(" & ErrMsg & ")", 1, Request.QueryString())
				End If
			End If
		Set rsTmp = Nothing
	ElseIf UserYGDM <> "" Then
		Set rsTmp = Server.CreateObject("ADODB.RecordSet")
			rsTmp.Open("Select * From HR_Teacher Where YGDM='" & UserYGDM & "'"), Conn, 1, 3
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				If rsTmp("LoginPass") <> MD5(OldPass, 16) Then
					ErrMsg = "原密码不正确，请重新输入！"
					Call RecordFrontLog(1, ScriptName, "员工" & UserYGXM & "[工号：" & UserYGDM & "]，操作：修改密码(" & ErrMsg & ")", 0, Request.QueryString())
					tmpJson = Replace(vMSG, "[@ErrMsg]", ErrMsg)
				Else
					rsTmp("LoginPass") = MD5(NewPass, 16)
					rsTmp.Update
					ErrMsg = "修改密码成功，请用新密码登陆！"
					vMSG = Replace(vMSG, "false", "true") : vMSG = Replace(vMSG, "500", "0")
					tmpJson = Replace(vMSG, "[@ErrMsg]", ErrMsg)
					Call RecordFrontLog(1, ScriptName, "员工" & UserYGXM & "[工号：" & UserYGDM & "]，操作：修改密码(" & ErrMsg & ")", 1, Request.QueryString())
				End If
			End If
		Set rsTmp = Nothing
	End If
	Response.Write tmpJson
End Sub

Sub Message()
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	SiteTitle = "我的消息"

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.ReadNO {color:#f00;}" & vbCrlf
	tmpHtml = tmpHtml & "		.pageBar {box-sizing:border-box;padding-top:8px;}" & vbCrlf
	tmpHtml = tmpHtml& "		.msgTips {line-height:40px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.msgTips i {color:#f30;font-size:20px;position: relative;top:3px;} .msgTips b {color:#f00;}" & vbCrlf
	tmpHtml = tmpHtml & "		.ShowCourse, .BackApply, .Transfer {cursor: pointer;color:#000}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)

	tmpHtml = vbCrlf & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "			layer.config({skin:""layer-hr""});var loadInit = layer.load();" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf

	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	tmpHtml = "<a href=""" & ParmPath & "MyCenter/Message.html"">" & SiteTitle & "</a><a><cite>全部消息</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Dim CountMsg : CountMsg = 0
	sqlTmp = "Select Count(ID) From HR_Message Where isRead=" & HR_False & " And ReceiverID=" & UserYGDM & ""

	Set rsTmp = Conn.Execute(sqlTmp)
		CountMsg = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing

	Response.Write "<div class=""hr-body-w800 hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	Response.Write "		<legend>我的消息</legend>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "	<div class=""msgTips""><i class=""hr-icon"">&#xee79;</i> "
	If CountMsg > 0 Then Response.Write " 您有 " & CountMsg & " 未读消息！"
	Response.Write " [<b>红色标题</b>为未读消息]</div>" & vbCrlf
	Response.Write "	<div class=""layui-collapse"" lay-filter=""myMessage"">" & vbCrlf

	sqlTmp = "Select * From HR_Message Where ReceiverID=" & UserID & ""
	sqlTmp = sqlTmp & " Order By isRead ASC, SendTime DESC"

	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0 : CurrentPage = 1 : MaxPerPage = tLimit
			If tPage > 0 Then CurrentPage = tPage
			If MaxPerPage <= 0 Then MaxPerPage = 10
			strFileName = ParmPath & "MyCenter/Message.html"

			TotalPut = rsTmp.Recordcount
			If TotalPut > 0 Then
				If CurrentPage < 1 Then CurrentPage = 1
				If (CurrentPage - 1) * MaxPerPage > TotalPut Then
					If (TotalPut Mod MaxPerPage) = 0 Then
						CurrentPage = TotalPut \ MaxPerPage
					Else
						CurrentPage = TotalPut \ MaxPerPage + 1
					End If
				End If
				If CurrentPage > 1 Then
					If (CurrentPage - 1) * MaxPerPage < TotalPut Then
						rsTmp.Move (CurrentPage - 1) * MaxPerPage
					Else
						CurrentPage = 1
					End If
				End If
			End If
			Do While Not rsTmp.EOF
				Response.Write "		<div class=""layui-colla-item"">" & vbCrlf
				Response.Write "			<h2 class=""layui-colla-title"
				If HR_CBool(rsTmp("isRead")) = False Then Response.Write " ReadNO"
				Response.Write """ data-id=""" & rsTmp("ID") & """>" & rsTmp("Title") & "　时间：" & FormatDate(rsTmp("SendTime"), 1) & "</h2>" & vbCrlf
				Response.Write "			<div class=""layui-colla-content""><p>" & rsTmp("Message") & "</p><h6></h6></div>" & vbCrlf
				Response.Write "		</div>" & vbCrlf
				rsTmp.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		Else
			Response.Write "		<div class=""layui-colla-item"">" & vbCrlf
			Response.Write "			<h2 class=""layui-colla-title"">您还没有任何消息</h2>" & vbCrlf
			Response.Write "			<div class=""layui-colla-content"">提示：您当前没有任何消息！</div>" & vbCrlf
			Response.Write "		</div>" & vbCrlf
		End If
	Set rsTmp = Nothing
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-rows pageBar"">" & vbCrlf
	Response.Write "		<div class=""Page_left""></div>" & vbCrlf
	Response.Write "		" & ShowPage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "条消息", True) & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		layer.closeAll(""loading"");" & vbCrlf
	tmpHtml = tmpHtml & "		element.on(""collapse(myMessage)"", function(data){" & vbCrlf
	tmpHtml = tmpHtml & "			if(data.show){" & vbCrlf
	tmpHtml = tmpHtml & "				$.getJSON(""" & ParmPath & "MyCenter/UpdateMessage.html"", {ID:data.title.data(""id"")}, function(result){" & vbCrlf
	tmpHtml = tmpHtml & "					data.title.removeClass(""ReadNO"");" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			};" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".ShowCourse"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var itemID = $(this).data(""itemid""), id = $(this).data(""id""), ygdm = $(this).data(""sender"");" & vbCrlf
	tmpHtml = tmpHtml & "			layer.open({type:2,id:""ShowWin"",content:""" & ParmPath & "Course/Preview.html?SendBtn=2&ItemID="" + itemID + ""&ID=""+ id, title:[""查看课程业绩"",""font-size:16""], area:[""680px"", ""82%""]});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".BackApply"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var itemID = $(this).data(""itemid""), id = $(this).data(""id""), ygdm = $(this).data(""sender"");" & vbCrlf
	tmpHtml = tmpHtml & "			layer.open({type:1,id:""BackWin"",content:"""", title:[""退回申请"",""font-size:16""], area:[""680px"", ""360px""]});" & vbCrlf
	tmpHtml = tmpHtml & "			$.get(""" & ParmPath & "MyCenter/BackApply.html"",{ItemID:itemID, ID:id, YGDM:ygdm}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#BackWin"").html(strForm);" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#ApplyPost"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "					var strExplain = $(""#Explain"").val();" & vbCrlf
	tmpHtml = tmpHtml & "					if(strExplain == """"){layer.msg(""您没有填写回退原因！"");return false;}" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "MyCenter/SendBackApply.html"",$(""#ApplyForm"").serialize(), function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.msg(strForm.reMessge,{icon:6,time:0,btn:""关闭""},function(){" & vbCrlf
	tmpHtml = tmpHtml & "							layer.closeAll();" & vbCrlf
	tmpHtml = tmpHtml & "							return false;" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".Transfer"").on(""click"", function(){" & vbCrlf
	'tmpHtml = tmpHtml & "			layer.open({type:1, id:""TranWin"", title:[""转交超管取消息审核"",""font-size:16""], area:[""650px"", ""360px""]});" & vbCrlf
	tmpHtml = tmpHtml & "			var itemID = $(this).data(""itemid""), id = $(this).data(""id""), ygdm = $(this).data(""sender"");" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "MyCenter/TransferApply.html"",{ItemID:itemID, ID:id, YGDM:ygdm}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.msg(strForm.reMessge,{icon:6,time:0,btn:""关闭""},function(){ });" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	'tmpHtml = tmpHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub
Sub UpdateMessage()
	Dim tmpJson, tmpID : tmpID = HR_Clng(Request("ID"))
	If tmpID > 0 Then Conn.Execute("Update HR_Message Set isRead=" & HR_True & " Where ID=" & tmpID )
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""消息 " & tmpID & " 设置为已读！"",""ReStr"":""操作完成！""}"
	Response.Write tmpJson
End Sub

Sub BackApply()		'退回申请
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tYGDM : tYGDM = HR_Clng(Request("YGDM"))
	ErrMsg = ""

	Dim tItemName, tTemplate, lenField, tFieldHead, tSheetName
	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			lenField = HR_Clng(rsTmp("FieldLen"))
			tFieldHead = Trim(rsTmp("FieldHead"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing
	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title"" style=""margin-top:1px;""><legend>退回您在 " & tItemName & " 业绩修改申请</legend></fieldset>" & vbCrlf
	Response.Write "	<form class=""layui-form layui-form-pane"" id=""ApplyForm"" name=""ApplyForm"" lay-filter=""ApplyForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layui-form-item layui-form-text"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">退回原因：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><textarea name=""Explain"" id=""Explain"" placeholder=""备注"" class=""layui-textarea""></textarea></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<input name=""ItemID"" type=""hidden"" value=""" & tItemID & """><input name=""ID"" type=""hidden"" value=""" & tmpID & """>" & vbCrlf
	Response.Write "		<input name=""ygdm"" type=""hidden"" value=""" & tYGDM & """><input name=""userid"" type=""hidden"" value=""" & UserID & """>" & vbCrlf
	Response.Write "		<div class=""searchBtn"">" & vbCrlf
	Response.Write "			<button class=""layui-btn"" type=""button"" id=""ApplyPost"" title=""发送退回消息""><i class=""hr-icon"">&#xebc5;</i>发送消息</button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "		layer.closeAll(""loading"");" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#ApplyPost"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var strExplain = $(""#Explain"").val();" & vbCrlf
	tmpHtml = tmpHtml & "			if(strExplain ==""""){layer.msg(""您没有填写申请的理由！"");return false;}" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "MyCenter/SendBackApply.html"",$(""#ApplyForm"").serialize(), function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.msg(strForm.reMessge,{icon:6,time:0,btn:""关闭""},function(){" & vbCrlf
	'tmpHtml = tmpHtml & "				var index1 = parent.layer.getFrameIndex(window.name);" & vbCrlf
	'tmpHtml = tmpHtml & "				parent.layer.close(index1);" & vbCrlf		'关闭自身，在iframe页面
	tmpHtml = tmpHtml & "				parent.layer.closeAll();" & vbCrlf
	tmpHtml = tmpHtml & "				return false;" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	'strFootHtml = Replace(strFootHtml, "[@Foot_script]", strHtml)
	'Response.Write strFootHtml

End Sub

Sub SendBackApply()
	Dim tExplain : tExplain = Trim(Request("Explain"))
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tID : tID = HR_Clng(Request("ID"))

	Dim tItemName, tTemplate, lenField, tFieldHead, tSheetName
	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			lenField = HR_Clng(rsTmp("FieldLen"))
			tFieldHead = Trim(rsTmp("FieldHead"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = "业绩考核项目不存在！【ID：" & tItemID & "】<br>"
		End If
	Set rsTmp = Nothing

	If Not(ChkTable(tSheetName)) Then
		ErrMsg = ErrMsg & "数据表 " & tSheetName & " 不存在！<br>"
	End If

	Dim tYGXM, tYGDM, SentMsg
	Set rsTmp = Conn.Execute("Select * From " & tSheetName & " Where ID=" & tID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tYGXM = Trim(rsTmp("VA2"))
			tYGDM = HR_Clng(rsTmp("VA1"))
		Else
			ErrMsg = ErrMsg & tItemName & "课程业绩不存在或已删除！【ID：" & tID & "】<br>"
		End If
	Set rsTmp = Nothing
	If ErrMsg <> "" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}"
		Exit Sub
	End If
	ErrMsg = tYGXM & "老师，您申请修改课程业绩，考核项目：" & tItemName & "，序号：" & tID & "，教师" & tYGXM & "[工号 " & tYGDM & "]不能修改。"
	ErrMsg = ErrMsg & " <span class=""ShowCourse"" data-ItemID=""" & tItemID & """ data-id=""" & tID & """ data-sender=""" & tYGDM & """>【查看】</span>"
	ErrMsg = ErrMsg & "<br>退回原因：" & tExplain
	ErrMsg = ErrMsg & "<br>退回时间：" & FormatDate(Now(), 1)
	SentMsg = SendMessage(0, tYGDM, "修改课程业绩申请被退回", ErrMsg, 0)
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""退回消息已经发送给 " & tYGXM & " 老师！<br />"",""ReStr"":""操作成功！""}"
End Sub

Sub TransferApply()
	Dim tExplain : tExplain = Trim(Request("Explain"))
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tID : tID = HR_Clng(Request("ID"))

	Dim tItemName, tTemplate, tSheetName
	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = "业绩考核项目不存在！【ID：" & tItemID & "】<br>"
		End If
	Set rsTmp = Nothing

	Dim tYGXM, tYGDM, SentMsg
	Set rsTmp = Conn.Execute("Select * From " & tSheetName & " Where ID=" & tID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tYGXM = Trim(rsTmp("VA2"))
			tYGDM = HR_Clng(rsTmp("VA1"))
		Else
			ErrMsg = tItemName & "课程业绩不存在或已删除！【ID：" & tID & "】<br>"
		End If
	Set rsTmp = Nothing
	If UserRank > 1 Then
		ErrMsg = ErrMsg & "您已有取消审核的权限，无需转发！<br>"
	End If

	If ErrMsg <> "" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}"
		Exit Sub
	End If


	Set rsTmp = Conn.Execute("Select * From HR_User Where ManageRank>1")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				ErrMsg = tYGXM & "老师申请修改课程业绩初审通过，请您取消审核，考核项目：" & tItemName & "，序号：" & tID & "，教师" & tYGXM & "[工号 " & tYGDM & "]，本消息由管理员“" & UserName & "”发送。"
				ErrMsg = ErrMsg & " <a href=""" & ParmPath & "Course.html?ItemID=" & tItemID & "&SearchWord=" & tYGDM & """>【查看】</a>"
				SentMsg = SendMessage(1, rsTmp("UserID"), tYGXM & "申请修改课程业绩，请取消审核！", ErrMsg, 0)
				rsTmp.MoveNext
			Loop
		End If
	Set rsTmp = Nothing
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""向高级管理员发送取消审核消息成功！<br />"",""ReStr"":""操作成功！""}"
End Sub
%>