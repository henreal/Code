<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim SubButTxt : SiteTitle = "业绩报表"
Dim arrStudentType : arrStudentType = Split(XmlText("Common", "StudentType", ""), "|")

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index", "List" Call MainBody()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveForm" Call SaveForm()			'Disable
	Case "AllData" Call getList()			'Disable
	Case "Preview" Call Preview()			'Disable
	Case "Collect" Call Collect()			'学时汇总
	Case "ImportGrade" Call ImportGrade()
	Case "ImportSave" Call ImportSave()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tExcelFileName : tExcelFileName = "PB" & FormatDate(Date(), 2) & ".xls" 
	'tmpHtml = "<style type=""text/css"">" & vbCrlf
	'tmpHtml = tmpHtml & "	</style>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		layui.use([""layer""], function(){ layer.load(1); });" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	tmpHtml = "<a href=""" & ParmPath & "Achieve/List.html"">" & SiteTitle & "</a><a><cite>报表首页</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-sides-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field site-demo-button"" style=""margin-top: 30px;"">" & vbCrlf
	Response.Write "		<legend>选择报表</legend>" & vbCrlf
	Response.Write "		<div class=""hr-shrink-x20"">" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-lg layui-btn-normal"" id=""show1"">查看报表</button>" & vbCrlf
	If UserRank > 0 Then
		Response.Write "			<button class=""layui-btn layui-btn-lg layui-btn-normal"" id=""show2"">报表下载</button>" & vbCrlf
		Response.Write "			<button class=""layui-btn layui-btn-lg layui-btn-normal"" id=""show3"">前100名</button>" & vbCrlf
		Response.Write "			<button class=""layui-btn layui-btn-lg layui-btn-normal"" id=""show4"">等级导入</button>" & vbCrlf
	End If
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "	$(""#show1"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		location.href = """ & ParmPath & "Tab/ExportExcel.html?total=false&teacher=&sort=&ksdm=&limit="";" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#show2"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		layer.confirm(""业绩报表导出可能要花点时间，请不要关闭本窗口！<br>导出的文件会自动保存到“我的电脑”的“下载”目录中<br>文件名为：" & tExcelFileName & """,{icon: 3, title:""导出提示""},function(index){" & vbCrlf
	strHtml = strHtml & "			location.href = """ & ParmPath & "Tab/ExportExcel.html?excel=true&total=false&teacher=&sort=&ksdm=&limit=10000"";" & vbCrlf
	strHtml = strHtml & "			layer.close(layer.index);" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#show3"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		location.href = """ & ParmPath & "Tab/ExportExcel.html?total=false&teacher=&sort=2&ksdm=&limit=100"";" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#show4"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		layer.open({type:2, id:""imGrade"", content:""" & ParmPath & "Achieve/ImportGrade.html"",title:[""导入等级"",""font-size:16""],offset:[""20%"", ""15%""],area:[""560px"", ""360px""]});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$("".navBtn a"").html(""<i class='hr-icon hr-icon-top'>&#xf351;</i>报表帮助"");" & vbCrlf
	strHtml = strHtml & "	$("".navBtn"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		location.href = """ & ParmPath & "Help.html?file=helpAchieve.pdf"";" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table; element = layui.element, form = layui.form;" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "		$(""#SearchBtn"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:1, id:""ReloadWin"",content:""<div class='reload'><p><i class='hr-icon'>&#xefe3;</i>正在生成业绩报表，可能需要几分钟，请稍等……</p></div>"", title:[""生成业绩报表"",""font-size:16""],area:[""630px"", ""420px""]});" & vbCrlf
	strHtml = strHtml & "			var ygdm=$(this).data(""ygdm""), loadTips = layer.load(1);" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "Tab/Reload.html"",{YGDM:ygdm}, function(strForm){" & vbCrlf
	strHtml = strHtml & "				$(""#ReloadWin"").html(strForm.reMessge);layer.close(loadTips);" & vbCrlf
	strHtml = strHtml & "				layer.close(loadTips);" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub

Sub Collect()
	SiteTitle = "项目学时汇总"
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background:#f3f3f3} " & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)

	tmpHtml = "<a href=""" & ParmPath & "Achieve/Collect.html"">" & SiteTitle & "</a><a><cite>汇总表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-card"">" & vbCrlf
	Response.Write "		<div class=""layui-card-header"">学时汇总【" & DefYear & "】</div>" & vbCrlf
	Response.Write "		<div class=""layui-card-body"">" & vbCrlf
	Dim tSumVA3, noSumVA3, CountYGDM, tItemName, tSheetName
	Set rs = Conn.Execute("Select * From HR_Class Where ModuleID=1001 Order By ClassType ASC, RootID ASC, OrderID ASC")
		If Not(rs.BOF And rs.EOF) Then
			Response.Write "	<table class=""layui-table"">" & vbCrlf
			Response.Write "		<thead><tr><th>项目名称</th><th>已审学时数</th><th>未审学时数</th><th>总学时数</th><th>教师数</th></tr></thead>" & vbCrlf
			Response.Write "		<tbody>" & vbCrlf
			Do While Not rs.EOF
				tSumVA3 = 0 : noSumVA3 = 0 : CountYGDM=0
				tItemName = Trim(rs("ClassName"))
				If rs("ParentID") > 0 Then tItemName = "　" & tItemName
				Response.Write "<tr><td>" & tItemName & "</td>" & vbCrlf


				If rs("Child") > 0 Then
					Response.Write "	<td></td>" & vbCrlf
				Else
					tSheetName = "HR_Sheet_" & rs("ClassID")
					If ChkTable(tSheetName) Then
						Set rsTmp = Conn.Execute("Select Sum(VA3) From " & tSheetName & " Where scYear=" & DefYear & " And Passed=" & HR_True)
							tSumVA3 = rsTmp(0)
						Set rsTmp = Nothing
						Set rsTmp = Conn.Execute("Select Sum(VA3) From " & tSheetName & " Where scYear=" & DefYear & " And Passed=" & HR_False)
							noSumVA3 = rsTmp(0)
						Set rsTmp = Nothing
						Set rsTmp = Server.CreateObject("ADODB.RecordSet")
							rsTmp.Open("Select Count(VA1) From " & tSheetName & " Where scYear=" & DefYear & " Group By VA1"), Conn, 1, 1
							CountYGDM = rsTmp.Recordcount
						Set rsTmp = Nothing
					End If
					Response.Write "	<td>" & HR_CDbl(tSumVA3) & "</td>" & vbCrlf
					Response.Write "	<td>" & HR_CDbl(noSumVA3) & "</td>" & vbCrlf
					Response.Write "	<td>" & HR_CDbl(tSumVA3) + HR_CDbl(noSumVA3) & "</td>" & vbCrlf
					Response.Write "	<td>" & CountYGDM & "</td>" & vbCrlf
				End If
				Response.Write "</tr>" & vbCrlf
				rs.MoveNext
			Loop
			Response.Write "		</tbody>" & vbCrlf
			Response.Write "	</table>" & vbCrlf
		End If
	Set rs = Nothing
	Response.Write "		" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element"", ""form""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table; element = layui.element, form=layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub ImportGrade()
	Dim tmpHtml, xlsUrl
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@Page_Title]", "导入业绩等级_" & SiteName)

	tmpHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />" & vbCrlf
	tmpHtml = tmpHtml & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.soBox .layui-btn {height: 30px;line-height: 30px;padding: 0 10px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = Replace(strHtml, "[@Head_style]", tmpHtml)
	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js?v=1.11.2""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js?v=2.3.0""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		$(document).ready(function(){ });" & vbCrlf
	tmpHtml = tmpHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var layer = layui.layer, element = layui.element; layer.config({skin:""layer-hr""}); var loadInit = layer.load();" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf
	strHtml = Replace(strHtml, "[@Head_script]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field site-demo-button"">" & vbCrlf
	Response.Write "		<legend>导入业绩等级</legend>" & vbCrlf
	Response.Write "		<form class=""layui-form layui-form-pane"" id=""ImportForm"" name=""ImportForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layer-hr-box"" id=""xlsBox"">" & vbCrlf
	Response.Write "			<div class=""layui-form-item""><label class=""layui-form-label"">Excel文件：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-block""><input type=""text"" name=""xlsUrl"" id=""xlsUrl"" value=""" & xlsUrl & """ placeholder=""请上传Excel文件"" class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-form-item soBox"">" & vbCrlf
	Response.Write "				<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-danger"" id=""upExcel""><i class=""hr-icon hr-icon-top"">&#xedd3;</i>上传</button></div>" & vbCrlf
	Response.Write "				<div class=""layui-inline searchBtn""><button class=""layui-btn"" lay-submit lay-filter=""stepNext""><i class=""hr-icon hr-icon-top"">&#xf051;</i>保存</button></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		</form>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = getPageFoot(1)
	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""form"", ""element"", ""upload""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var element = layui.element, form = layui.form, upload = layui.upload; layer.closeAll(""loading"");" & vbCrlf
	tmpHtml = tmpHtml & "		upload.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem: '#upExcel',url: '" & ParmPath & "UploadFile.htm?UploadDir=Excel'" & vbCrlf
	tmpHtml = tmpHtml & "			,multiple: false,accept:'file',done: function(res, index){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#xlsUrl"").val(res.data.src);" & vbCrlf
	tmpHtml = tmpHtml & "			},error: function (index, upload){console.log(index);}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		form.on(""submit(stepNext)"", function(PostData){" & vbCrlf
	tmpHtml = tmpHtml & "			var xls1 = PostData.field.xlsUrl;" & vbCrlf
	tmpHtml = tmpHtml & "			if(xls1==""""){layer.tips(""您还没有上传Excel数据文件！"",""#xlsUrl"",{tips: [1, ""#393D49""]});return false;}" & vbCrlf
	tmpHtml = tmpHtml & "			parent.layer.open({type:2, id:""ImportWin"",content:""" & ParmPath & "Achieve/ImportSave.html?xlsUrl="" + xls1" & vbCrlf
	tmpHtml = tmpHtml & "				, title:[""保存上传数据"",""font-size:16""],area:[""630px"", ""420px""],cancel:function(){parent.layer.closeAll();}" & vbCrlf
	tmpHtml = tmpHtml & "				,btn:""关闭"",yes:function(){parent.layer.closeAll();}" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "			return false;" & vbCrlf		'禁止提交
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = Replace(strHtml, "[@Foot_script]", tmpHtml)
	Response.Write strHtml
End Sub

Sub ImportSave()
	Dim xlsFile : xlsFile = Trim(Request("xlsUrl"))
	Dim tmpHtml, strTmp, tmpItemArr, tmpArr, tmpCount
	tmpCount = 0

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@Page_Title]", "导入业绩等级_" & SiteName)
	tmpHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />" & vbCrlf
	tmpHtml = tmpHtml & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.reload p {margin-top:2rem;text-align:center;font-size:1rem;} .reload p i{font-size:2rem;padding-right:15px;color:#f30}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = Replace(strHtml, "[@Head_style]", tmpHtml)
	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js?v=1.11.2""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js?v=2.3.0""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		$(document).ready(function(){ });" & vbCrlf
	tmpHtml = tmpHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var layer = layui.layer, element = layui.element; layer.config({skin:""layer-hr""});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf
	strHtml = Replace(strHtml, "[@Head_script]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class='reload'><p><i class='layui-icon layui-anim layui-anim-rotate layui-anim-loop'>&#xe63d;</i>正在保存数据，可能需要几分钟，请稍等……</p></div>" & vbCrlf
	Response.Flush

	If fso.FileExists(Server.MapPath(xlsFile)) Then
		strTmp = GetHttpStr(apiHost & "/Manage/ReadExcel.htm?type=1&xlsFile=" & xlsFile, "", 3, 10)
		If HR_IsNull(strTmp) Then
			ErrMsg = "Excel数据文件 " & xlsFile & " 没有记录！"
		Else
			tmpItemArr = Split(strTmp, "@@")
			ErrMsg = "<ul>"
			For i = 0 To Ubound(tmpItemArr)
				If HR_IsNull(tmpItemArr(i)) = False Then 
					tmpArr = Split(tmpItemArr(i), "||")
					If Ubound(tmpArr) = 3 Then
						'Set rsTmp = Server.CreateObject("ADODB.RecordSet")
						'	rsTmp.Open("Select * From HR_KPI Where YGDM=" & HR_Clng(tmpArr(0))), Conn, 1, 3
						'	If rsTmp.BOF And rsTmp.EOF Then
						'		ErrMsg = ErrMsg & "<li>" & i + 2 & "、教师“" & tmpArr(1) & "[" & tmpArr(0) & "]”没有业绩或业绩未审核！</li>"
						'	Else
						'		rsTmp("Grade") = tmpArr(3)
						'		rsTmp.Update
						'		Conn.Execute("Update HR_KPI_SUM Set Grade='" & tmpArr(3) & "' Where YGDM=" & HR_Clng(tmpArr(0)))	'同步学时缓存库
								tmpCount = tmpCount + 1
						'	End If
						'Set rsTmp = Nothing
					Else
						ErrMsg = ErrMsg & "<li>数据格式错误</li>"
						Exit For
					End If
				Else
					ErrMsg = ErrMsg & "<li>" & i + 2 & "、空数据</li>"
				End If
			Next
			ErrMsg = "<li class=""count"">共有 " & tmpCount & " 条数据导入成功！</li>" & ErrMsg
			ErrMsg = ErrMsg & "</ul>"
		End If
	Else
		ErrMsg = "Excel数据文件 " & xlsFile & " 不存在！"
	End If
	Response.Write ErrMsg & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var element = layui.element, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		$("".reload"").html("""");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@Foot_script]", tmpHtml)
	Response.Write strHtml
End Sub
%>