<%
Sub CourseBody()
	SiteTitle = "课程管理"
	Dim arrCourse : arrCourse = Split(XmlText("Common", "Course", ""), "|")
	Dim strRows : strRows = ""
	For i = 0 To Ubound(arrCourse)
		If i > 0 Then strRows = strRows & ","
		strRows = strRows & "{""ID"":" & i + 1 & ",""Course"":""" & arrCourse(i) & """}"
	Next

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		#EditBox em {padding:20px 20px 0 0;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		#EditBox em .itemBar {height:100px;background:#fff;border:5px solid #eee;box-sizing: border-box;padding:15px;}" & vbCrlf
	tmpHtml = tmpHtml & "		#EditBox em .itemBar h3 {line-height:30px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	'tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "Setup/Index.html"">系统管理</a><a><cite>" & SiteTitle & "</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:10px;"">" & vbCrlf
	Response.Write "	<legend>课程管理</legend>" & vbCrlf
	Response.Write "	<div class=""layer-hr-box"">" & vbCrlf
	Response.Write "	<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""hr-grids"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">课程名称：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""Course"" value="""" placeholder=""请输入课程"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""hr-sides-x10""><button class=""layui-btn layui-btn-sm"" lay-submit lay-filter=""SubPost"">添加课程</button></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<input type=""hidden"" name=""Modify"" value=""False"">" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf
	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form layui-row"" id=""EditBox""></div>"
	Response.Write "</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "	{{#  layui.each(d.rows, function(index, item){ }}" & vbCrlf
	Response.Write "	<em class=""layui-col-xs5 layui-col-sm4 layui-col-md3""><div class=""itemBar""><h3 class=""rowTitle"">{{ item.Course }}</h3>"
	Response.Write "<h4><button class=""layui-btn layui-btn-sm layui-btn-normal ModifyBtn"" type=""button"" data-id=""{{ item.ID }}"" data-value=""{{ item.Course }}"" title=""修改""><i class=""hr-icon"">&#xebf7;</i></button>"
	Response.Write "<button class=""layui-btn layui-btn-sm layui-btn-normal DeleteBtn"" type=""button"" data-id=""{{ item.ID }}"" data-value=""{{ item.Course }}"" title=""删除""><i class=""hr-icon"">&#xf014;</i></button>"
	Response.Write "</h4></div></em>" & vbCrlf
	Response.Write "	{{#  }); }}" & vbCrlf
	Response.Write "</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf


	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "	layui.use([""form"", ""laytpl"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var form = layui.form, laytpl = layui.laytpl;" & vbCrlf
	strHtml = strHtml & "		element = layui.element;" & vbCrlf
	strHtml = strHtml & "		var data = {""title"":""课程"",""rows"":[" & strRows & "]};" & vbCrlf
	strHtml = strHtml & "		$(""#AddCourse"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:2,id:""popWin"",content:""" & ParmPath & "Setup/EditCourse.html"",title:[""课程管理"",""font-size:16""],area:[""700px"", ""550px""],maxmin:true});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		var getTpl = barTable.innerHTML, view = document.getElementById(""EditBox"");" & vbCrlf
	strHtml = strHtml & "		laytpl(getTpl).render(data, function(html){" & vbCrlf
	strHtml = strHtml & "			view.innerHTML = html;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$("".ModifyBtn"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			var tValue = $(this).data(""value""), tID = $(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "			layer.prompt({title:""编辑课程名称"",value:tValue},function(value, index, elem){" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "Setup/SaveCourse.html"",{ID:tID,Course:value}, function(reData){ window.location.reload(); });" & vbCrlf
	strHtml = strHtml & "				window.location.reload();" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$("".DeleteBtn"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			var tValue = $(this).data(""value""), tID = $(this).data(""ID"");" & vbCrlf
	strHtml = strHtml & "			layer.confirm(""真的删除选中的课程吗？"", {icon:3,title: ""删除提醒""}, function(index){" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "Setup/DeleteCourse.html"",{ID:tID,Course:tValue}, function(reData){ window.location.reload(); });" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Setup/SaveCourse.html"",PostData.field, function(result){ });" & vbCrlf
	strHtml = strHtml & "			window.location.reload();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub
Sub SaveCourse()
	Dim tmpCourse, tCourse, tResult, tmpJson, strTmp
	Dim tmpID :tmpID = HR_Clng(Request("ID"))
	Dim arrCourse : arrCourse = Split(XmlText("Common", "Course", ""), "|")

	tCourse = Trim(ReplaceBadChar(Request("Course")))
	tmpCourse = XmlText("Common", "Course", "")
	tResult = False
	If tmpID > 0 Then
		For i = 0 To Ubound(arrCourse)
			If tmpID = i + 1 And Trim(tCourse) <> "" Then
				strTmp = strTmp & Trim(tCourse) & "|"
			Else
				strTmp = strTmp & Trim(arrCourse(i)) & "|"
			End If
		Next
		strTmp = FilterArrNull(strTmp, "|")
		tResult = UpdateXmlText("Common", "Course", strTmp)
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""课程 修改成功！"",""ReStr"":""操作成功！""}"
	Else
		If Instr(tmpCourse, "|") > 0 And tCourse <> "" Then
			tmpCourse = tmpCourse & "|" & tCourse
			tResult = UpdateXmlText("Common", "Course", tmpCourse)
			tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""课程 添加成功！"",""ReStr"":""操作成功！""}"
		Else
			tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""课程 添加失败！"",""ReStr"":""操作失败！""}"
		End If
	End If
	Response.Write tmpJson
End Sub
Sub DeleteCourse()
	Dim tResult, tmpJson, tmpCourse : strTmp = ""
	Dim tCourse : tCourse = Trim(ReplaceBadChar(Request("Course")))
	Dim arrCourse : arrCourse = Split(XmlText("Common", "Course", ""), "|")
	tResult = False

	For i = 0 To Ubound(arrCourse)
		If tCourse <> Trim(arrCourse(i)) Then strTmp = strTmp & Trim(arrCourse(i)) & "|"
	Next
	strTmp = FilterArrNull(strTmp, "|")
	If strTmp <> "" Then
		tResult = UpdateXmlText("Common", "Course", strTmp)
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""课程 删除成功！"",""ReStr"":""操作成功！""}"
	Else
		tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""课程 删除失败！"",""ReStr"":""操作失败！""}"
	End If
	Response.Write tmpJson
End Sub

Sub SetupSwitch()		'系统开关
	Dim tAddChecked, IsAdd : IsAdd = HR_CBool(XmlText("Common", "AddSwitch", "0"))
	If IsAdd Then tAddChecked = "checked"
	Dim tImportChecked, IsImport : IsImport = HR_CBool(XmlText("Common", "ImportSwitch", "0"))
	If IsImport Then tImportChecked = "checked"

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "Setup/Index.html"">系统管理</a><a><cite>课程开关</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:10px;"">" & vbCrlf
	Response.Write "	<legend>课程开关</legend>" & vbCrlf
	Response.Write "	<div class=""layer-hr-box"">" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">录入开关：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""checkbox"" " & tAddChecked & " name=""addSwitch"" lay-skin=""switch"" lay-filter=""addswitch"" lay-text=""开|关""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">导入开关：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""checkbox"" " & tImportChecked & " name=""importswitch"" lay-skin=""switch"" lay-filter=""importswitch"" lay-text=""开|关""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf

	Response.Write "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:10px;"">" & vbCrlf
	Response.Write "	<legend>业绩等级导入</legend>" & vbCrlf
	Response.Write "	<div class=""layer-hr-box"">" & vbCrlf
	Response.Write "		<button type=""button""  class=""layui-btn hr-btn_darkgreen"" id=""ImportGrade""><i class=""hr-icon"">&#xe890;</i>导入业绩等级</button>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf

	Response.Write "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:10px;"">" & vbCrlf
	Response.Write "	<legend>设置默认学年</legend>" & vbCrlf
	Response.Write "	<div class=""layer-hr-box"">" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">设置学年：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline"">" & vbCrlf
	Response.Write "					<select name=""yearType""  lay-filter=""scyear"" id=""yearType"">" & GetYearOption(0, DefYear) & "</select>" & vbCrlf
	Response.Write "				</div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf
	Response.Write "</form>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "	layui.use([""form"", ""laytpl"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var form = layui.form, laytpl = layui.laytpl;" & vbCrlf
	strHtml = strHtml & "		element = layui.element;" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "		form.on(""switch(addswitch)"", function(reData){" & vbCrlf			'添加开关
	strHtml = strHtml & "			var switch1 = this.checked ? 'true' : 'false';" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "Setup/SaveSwitch.html"",{switch:switch1}, function(reData){" & vbCrlf
	strHtml = strHtml & "				layer.msg(reData.reMessge, {time:0,btn:""关闭"",icon:1,offset: ['100px', '100px']});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		form.on(""switch(importswitch)"", function(reData){" & vbCrlf		'导入开关
	strHtml = strHtml & "			var switch1 = this.checked ? 'true' : 'false';" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "Setup/SaveSwitchImport.html"",{switch:switch1}, function(reData){" & vbCrlf
	strHtml = strHtml & "				layer.msg(reData.reMessge, {time:0,btn:""关闭"",icon:1,offset: ['100px', '300px']});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		form.on(""select(scyear)"", function(data){" & vbCrlf			'默认学年
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "Setup/SaveYear.html"",{scyear:data.value}, function(reData){" & vbCrlf
	strHtml = strHtml & "				layer.msg(reData.reMessge, {time:0,btn:""关闭"",icon:1,offset: ['100px', '300px']});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#ImportGrade"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:2, id:""imGrade"", content:""" & ParmPath & "Setup/ImportGrade.html"",title:[""导入等级"",""font-size:16""],offset:[""20%"", ""15%""],area:[""560px"", ""360px""]});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	'strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub

Sub SaveSwitch()
	Dim tmpJson, tSwitch : tSwitch = HR_CBool(Request("switch"))
	If tSwitch Then
		tmpJson = "添加课程启用"
		Call UpdateXmlText("Common", "AddSwitch", "1")
	Else
		tmpJson = "添加课程关闭"
		Call UpdateXmlText("Common", "AddSwitch", "0")
	End If
	tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""" & tmpJson & """,""ReStr"":""操作失败！""}"
	Response.Write tmpJson
End Sub

Sub SaveSwitchImport()
	Dim tmpJson, tSwitch : tSwitch = HR_CBool(Request("switch"))
	If tSwitch Then
		tmpJson = "导入启用"
		Call UpdateXmlText("Common", "ImportSwitch", "1")
	Else
		tmpJson = "导入关闭"
		Call UpdateXmlText("Common", "ImportSwitch", "0")
	End If
	tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""" & tmpJson & """,""ReStr"":""操作失败！""}"
	Response.Write tmpJson
End Sub
Sub SaveYear()
	Dim tmpJson, tYear : tYear = HR_CLng(Request("scyear"))
	ErrMsg = "更新学年失败！"
	If tYear>2018 And tYear<2030 Then
		Call UpdateXmlText("Common", "Year", tYear)
		ErrMsg = "学年已经更新为 " & tYear & "！"
	End If
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub

Sub BackBody()		'备份数据库
	Dim layUrl : layUrl = ParmPath & "Setup/GetBackdataJson.html"

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.ResultTips {display:none;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-field-title {margin:0;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")

	tmpHtml = "<a href=""" & ParmPath & "Setup/Index.html"">" & SiteTitle & "</a><a><cite>备份数据库</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "<fieldset class=""layui-elem-field"" style=""margin:10px;"">" & vbCrlf
	Response.Write "	<legend>备份数据库</legend>" & vbCrlf
	Response.Write "	<div class=""layer-hr-box"">" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline"">" & vbCrlf
	Response.Write "				<button class=""layui-btn layui-bg-cyan"" data-type=""backup"" id=""backup"" title=""备份数据""><i class=""hr-icon"">&#xf1c0;</i>备份数据</button>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""layer-hr-box ResultTips"">" & vbCrlf
	Response.Write "		<div class=""tips""></div>" & vbCrlf
	Response.Write "		<div class=""dataurl""><b><i class=""hr-icon"">&#xecb8;</i>下载地址：</b><a href=""""></a></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-workZones hr-sides-x20"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	Response.Write "		<legend>备份数据列表</legend>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf	'行工具
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-xs hr-btn_olive"" lay-event=""down"" title=""下载""><i class=""hr-icon"">&#xed27;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""form"", ""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var form = layui.form, table = layui.table;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#backup"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var load = layer.load(0,{offset:[""100px"",""200px""]});" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "Setup/getBackupData.html"", function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "				$("".tips"").html(reData.reMessge);" & vbCrlf
	tmpHtml = tmpHtml & "				$("".dataurl a"").attr(""href"",reData.bakfile);$("".dataurl a"").text(reData.bakfile);" & vbCrlf
	tmpHtml = tmpHtml & "				$("".ResultTips"").show(); layer.close(load);;" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",page:true,limit:30,skin:'',limits:[10,15,20,30,50,100,200]" & vbCrlf
	tmpHtml = tmpHtml & "			,text:{none:'暂时没有数据备份文件'},cols: [[" & vbCrlf				'设置表头
	tmpHtml = tmpHtml & "				{type:'checkbox',unresize:true,align:'center',width:50}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'name',title:'备份数据文件',minwidth:250}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'size',title:'大小',align:'right',sort:true,width:105,style:'color:#080',templet:function(res){return res.size + ' MB'}}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'date',title:'备份日期',sort:true,width:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'filetype',title:'文件类型',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{title:'下载',align:'center',unresize:true,width:70, toolbar:'#barTable'}" & vbCrlf
	tmpHtml = tmpHtml & "			]],url:""" & layUrl & """" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf		'监听工具栏
	tmpHtml = tmpHtml & "			if(obj.event === ""down""){" & vbCrlf
	tmpHtml = tmpHtml & "				location.href=obj.data.url;" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub BackupData()
	Dim tmpJson
	tmpJson = BackSQLData(0, "")		'备份数据库，取备份文件名
	tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""数据库备份成功！"",""ReStr"":""操作失败！"", ""bakfile"":""" & tmpJson & """}"
	Response.Write tmpJson
End Sub

Sub GetBackdataJson()	'取备份数据文件
	Dim tLimit : tLimit = HR_Clng(Request("limit"))
	Dim tPage : tPage = HR_Clng(Request("page"))
	Dim tmpJson, rsFolder, f, dataDir : dataDir = InstallDir & "BackData/"
	Dim iCount : iCount=0 : ErrMsg = ""
	If FSO.FolderExists(Server.MapPath(dataDir)) Then
		Set rsFolder = FSO.GetFolder(Server.MapPath(dataDir))
			iCount = rsFolder.Files.Count
			If iCount > 0 Then
				i = 0
				For Each f in rsFolder.files
					If i > 0 Then tmpJson = tmpJson & ","
					tmpJson = tmpJson & "{""name"":""" & f.Name & """,""size"":" & FormatNumber(f.Size/1024/1024,2,,,0) & ",""date"":""" & FormatDate(f.DateCreated, 1) & """,""filetype"":""" & f.Type & """,""url"":""" & dataDir & f.Name & """}"
					i = i + 1
				Next
			End If
		Set rsFolder = Nothing
		Response.Write "{""code"":0,""msg"":""查询成功"",""count"":" & iCount & ",""limit"":" & tLimit & ",""page"":" & tPage & ",""data"":[" & tmpJson & "]}"
	Else
		Response.Write "{""code"":404,""msg"":""备份目录 " & dataDir & " 不存在！"",""count"":" & iCount & ",""limit"":" & tLimit & ",""page"":" & tPage & ",""data"":[" & tmpJson & "]}"
	End If
End Sub

'班级管理
Sub TeachClassBody()
	SiteTitle = "班级管理"
	Dim arrTeachClass : arrTeachClass = Split(XmlText("Common", "TeachClass", ""), "|")
	Dim strRows : strRows = ""
	For i = 0 To Ubound(arrTeachClass)
		If i > 0 Then strRows = strRows & ","
		strRows = strRows & "{""ID"":" & i + 1 & ",""TeachClass"":""" & arrTeachClass(i) & """}"
	Next

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "	#EditBox em {padding: 20px 20px 0 0;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "	#EditBox em .itemBar {border:5px solid #eee;box-sizing: border-box;padding: 10px;}" & vbCrlf
	tmpHtml = tmpHtml & "	#EditBox em .itemBar h3 {min-height:3.5rem;font-size:1.1rem;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "Setup/Index.html"">系统管理</a><a><cite>" & SiteTitle & "</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:10px;"">" & vbCrlf
	Response.Write "	<legend>授课对象管理</legend>" & vbCrlf
	Response.Write "	<div class=""layer-hr-box"">" & vbCrlf
	Response.Write "	<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""hr-grids"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">班级名称：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""TeachClass"" value="""" placeholder=""请输入班级"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""hr-sides-x10""><button class=""layui-btn layui-btn-sm"" lay-submit lay-filter=""SubPost"">添加班级</button></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<input type=""hidden"" name=""Modify"" value=""False"">" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf
	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form layui-row"" id=""EditBox""></div>"
	Response.Write "</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "	{{#  layui.each(d.rows, function(index, item){ }}" & vbCrlf
	Response.Write "	<em class=""layui-col-xs5 layui-col-sm4 layui-col-md3""><div class=""itemBar"">" & vbCrlf
	Response.Write "		<h3 class=""rowTitle"">{{ item.TeachClass }}</h3>" & vbCrlf
	Response.Write "		<h4><button class=""layui-btn layui-btn-sm layui-btn-normal DeleteBtn"" type=""button"" data-id=""{{ item.ID }}"" data-value=""{{ item.TeachClass }}"" title=""删除""><i class=""hr-icon"">&#xf014;</i></button></h4>" & vbCrlf
	Response.Write "	</div></em>" & vbCrlf
	Response.Write "	{{#  }); }}" & vbCrlf
	Response.Write "</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "	layui.use([""form"", ""laytpl"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var form = layui.form, laytpl = layui.laytpl;" & vbCrlf
	strHtml = strHtml & "		element = layui.element;" & vbCrlf
	strHtml = strHtml & "		var data = {""title"":""班级"",""rows"":[" & strRows & "]};" & vbCrlf
	strHtml = strHtml & "		var getTpl = barTable.innerHTML, view = document.getElementById(""EditBox"");" & vbCrlf
	strHtml = strHtml & "		laytpl(getTpl).render(data, function(html){" & vbCrlf
	strHtml = strHtml & "			view.innerHTML = html;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$("".ModifyBtn"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			var tValue = $(this).data(""value""), tID = $(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "			layer.prompt({title:""编辑班级名称"",value:tValue},function(value, index, elem){" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "Setup/SaveTeachClass.html"",{ID:tID, TeachClass:value}, function(reData){ window.location.reload(); });" & vbCrlf
	strHtml = strHtml & "				window.location.reload();" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$("".DeleteBtn"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			var tValue = $(this).data(""value""), tID = $(this).data(""ID"");" & vbCrlf
	strHtml = strHtml & "			layer.confirm(""真的删除选中的班级吗？"", {icon:3,title: ""删除提醒""}, function(index){" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "Setup/DelTeachClass.html"",{ID:tID, Value:tValue}, function(reData){ window.location.reload(); });" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Setup/SaveTeachClass.html"",PostData.field, function(result){ });" & vbCrlf
	strHtml = strHtml & "			window.location.reload();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub
Sub SaveTeachClass()
	Dim tmpJson, strTmp, tmpID :tmpID = HR_Clng(Request("ID"))
	Dim tmpValue : tmpValue = Trim(ReplaceBadChar(Request("TeachClass")))
	Dim tmpNode : tmpNode = XmlText("Common", "TeachClass", "")

	If HR_IsNull(tmpNode) Then		'检查是否包含
		Call UpdateXmlText("Common", "TeachClass", tmpValue)
	ElseIf FoundInArr(tmpNode, tmpValue, "|") = False Then
		Call UpdateXmlText("Common", "TeachClass", tmpNode & "|" & tmpValue)
	End If
	Response.Write "{""Return"":true, ""Err"":0, ""reMessge"":""班级 保存成功！"", ""ReStr"":""操作成功！""}"
End Sub
Sub DelTeachClass()			'删除班级
	Dim strTmp : strTmp = ""
	Dim tmpValue : tmpValue = Trim(ReplaceBadChar(Request("Value")))
	Dim arrNode : arrNode = Split(XmlText("Common", "TeachClass", ""), "|")

	For i = 0 To Ubound(arrNode)
		If tmpValue <> Trim(arrNode(i)) Then strTmp = strTmp & Trim(arrNode(i)) & "|"
	Next
	If Ubound(arrNode) = 0 Then
		If tmpValue = Trim(arrNode(0)) Then strTmp = ""
	End If
	strTmp = FilterArrNull(strTmp, "|")
	Call UpdateXmlText("Common", "TeachClass", strTmp)
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""班级 删除成功！"",""ReStr"":""操作成功！""}"
End Sub

'教室管理
Sub ClassRoomBody()
	SiteTitle = "授课教室管理"
	Dim arrTeachClass : arrTeachClass = Split(XmlText("Common", "Classroom", ""), "|")
	Dim strRows : strRows = ""
	For i = 0 To Ubound(arrTeachClass)
		If i > 0 Then strRows = strRows & ","
		strRows = strRows & "{""ID"":" & i + 1 & ",""ClassRoom"":""" & arrTeachClass(i) & """}"
	Next

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "	#EditBox em {padding: 20px 20px 0 0;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "	#EditBox em .itemBar {border:5px solid #eee;box-sizing: border-box;padding: 10px;}" & vbCrlf
	tmpHtml = tmpHtml & "	#EditBox em .itemBar h3 {min-height:3.5rem;font-size:1.1rem;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "Setup/Index.html"">系统管理</a><a><cite>" & SiteTitle & "</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:10px;"">" & vbCrlf
	Response.Write "	<legend>" & SiteTitle & "管理</legend>" & vbCrlf
	Response.Write "	<div class=""layer-hr-box"">" & vbCrlf
	Response.Write "	<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""hr-grids"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">教室名称：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""ClassRoom"" value="""" placeholder=""请输入教室"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""hr-sides-x10""><button class=""layui-btn layui-btn-sm"" lay-submit lay-filter=""SubPost"">添加教室</button></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<input type=""hidden"" name=""Modify"" value=""False"">" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf
	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form layui-row"" id=""EditBox""></div>"
	Response.Write "</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "	{{#  layui.each(d.rows, function(index, item){ }}" & vbCrlf
	Response.Write "	<em class=""layui-col-xs5 layui-col-sm4 layui-col-md3""><div class=""itemBar"">" & vbCrlf
	Response.Write "		<h3 class=""rowTitle"">{{ item.ClassRoom }}</h3>" & vbCrlf
	Response.Write "		<h4><button class=""layui-btn layui-btn-sm layui-btn-normal DeleteBtn"" type=""button"" data-id=""{{ item.ID }}"" data-value=""{{ item.ClassRoom }}"" title=""删除""><i class=""hr-icon"">&#xf014;</i></button></h4>" & vbCrlf
	Response.Write "	</div></em>" & vbCrlf
	Response.Write "	{{#  }); }}" & vbCrlf
	Response.Write "</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "	layui.use([""form"", ""laytpl"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var form = layui.form, laytpl = layui.laytpl;" & vbCrlf
	strHtml = strHtml & "		element = layui.element;" & vbCrlf
	strHtml = strHtml & "		var data = {""title"":""班级"",""rows"":[" & strRows & "]};" & vbCrlf
	strHtml = strHtml & "		var getTpl = barTable.innerHTML, view = document.getElementById(""EditBox"");" & vbCrlf
	strHtml = strHtml & "		laytpl(getTpl).render(data, function(html){" & vbCrlf
	strHtml = strHtml & "			view.innerHTML = html;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$("".ModifyBtn"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			var tValue = $(this).data(""value""), tID = $(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "			layer.prompt({title:""编辑教室名称"",value:tValue},function(value, index, elem){" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "Setup/SaveClassRoom.html"",{ID:tID, ClassRoom:value}, function(reData){ window.location.reload(); });" & vbCrlf
	strHtml = strHtml & "				window.location.reload();" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$("".DeleteBtn"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			var tValue = $(this).data(""value""), tID = $(this).data(""ID"");" & vbCrlf
	strHtml = strHtml & "			layer.confirm(""真的删除选中的教室吗？"", {icon:3,title: ""删除提醒""}, function(index){" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "Setup/DelClassRoom.html"",{ID:tID, ClassRoom:tValue}, function(reData){ window.location.reload(); });" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Setup/SaveClassRoom.html"",PostData.field, function(result){ });" & vbCrlf
	strHtml = strHtml & "			window.location.reload();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub
Sub SaveClassRoom()
	Dim tmpJson, strTmp, tmpID :tmpID = HR_Clng(Request("ID"))
	Dim tmpValue : tmpValue = Trim(ReplaceBadChar(Request("ClassRoom")))
	Dim tmpNode : tmpNode = XmlText("Common", "Classroom", "")

	If HR_IsNull(tmpNode) Then		'检查是否包含
		Call UpdateXmlText("Common", "Classroom", tmpValue)
	ElseIf FoundInArr(tmpNode, tmpValue, "|") = False Then
		Call UpdateXmlText("Common", "Classroom", tmpNode & "|" & tmpValue)
	End If
	Response.Write "{""Return"":true, ""Err"":0, ""reMessge"":""班级 保存成功！"", ""ReStr"":""操作成功！""}"
End Sub
Sub DelClassRoom()			'删除班级
	Dim strTmp : strTmp = ""
	Dim tmpValue : tmpValue = Trim(ReplaceBadChar(Request("ClassRoom")))
	Dim arrNode : arrNode = Split(XmlText("Common", "Classroom", ""), "|")

	For i = 0 To Ubound(arrNode)
		If tmpValue <> Trim(arrNode(i)) Then strTmp = strTmp & Trim(arrNode(i)) & "|"
	Next
	If Ubound(arrNode) = 0 Then
		If tmpValue = Trim(arrNode(0)) Then strTmp = ""
	End If
	strTmp = FilterArrNull(strTmp, "|")
	Call UpdateXmlText("Common", "Classroom", strTmp)
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""教室 删除成功！"",""ReStr"":""操作成功！""}"
End Sub

Sub ImportGrade()		'导入业绩等级
	Dim xlsUrl
	
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)
	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field site-demo-button"">" & vbCrlf
	Response.Write "		<legend>导入业绩等级</legend>" & vbCrlf
	Response.Write "		<form class=""layui-form layui-form-pane"" id=""ImportForm"" name=""ImportForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layer-hr-box"" id=""xlsBox"">" & vbCrlf
	Response.Write "			<div class=""layui-form-item""><label class=""layui-form-label"">选择学年：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline"">" & vbCrlf
	Response.Write "					<select name=""EduYear""  lay-filter=""EduYear"" id=""EduYear"">" & GetYearOption(0, DefYear) & "</select>" & vbCrlf
	Response.Write "				</div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-form-item""><label class=""layui-form-label"">Excel文件：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-block""><input type=""text"" name=""xlsUrl"" id=""xlsUrl"" value=""" & xlsUrl & """ placeholder=""请上传Excel文件"" class=""layui-input"" ></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-form-item soBox"">" & vbCrlf
	Response.Write "				<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-danger"" id=""upExcel""><i class=""hr-icon hr-icon-top"">&#xedd3;</i>上传</button></div>" & vbCrlf
	Response.Write "				<div class=""layui-inline searchBtn""><button class=""layui-btn"" lay-submit lay-filter=""stepNext""><i class=""hr-icon hr-icon-top"">&#xf051;</i>保存</button></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		</form>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""form"", ""element"", ""upload""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var element = layui.element, form = layui.form, upload = layui.upload; layer.closeAll(""loading"");" & vbCrlf
	tmpHtml = tmpHtml & "		upload.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem: '#upExcel',url: '" & InstallDir & "API/UploadFile.htm?UploadDir=Excel'" & vbCrlf
	tmpHtml = tmpHtml & "			,multiple: false,accept:'file',done: function(res, index){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#xlsUrl"").val(res.data.src);" & vbCrlf
	tmpHtml = tmpHtml & "			},error: function (index, upload){console.log(index);}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		form.on(""submit(stepNext)"", function(PostData){" & vbCrlf
	tmpHtml = tmpHtml & "			var xls1 = PostData.field.xlsUrl, eduYear = PostData.field.EduYear;" & vbCrlf
	tmpHtml = tmpHtml & "			if(xls1==""""){layer.tips(""您还没有上传Excel数据文件！"",""#xlsUrl"",{tips: [1, ""#393D49""]});return false;}" & vbCrlf
	tmpHtml = tmpHtml & "			parent.layer.open({type:2, id:""ImportWin"",content:""" & ParmPath & "Setup/ImportSave.html?EduYear="" + eduYear + ""&xlsUrl="" + xls1" & vbCrlf
	tmpHtml = tmpHtml & "				, title:[""保存上传数据"",""font-size:16""],area:[""630px"", ""420px""],cancel:function(){parent.layer.closeAll();}" & vbCrlf
	tmpHtml = tmpHtml & "				,btn:""关闭"",yes:function(){parent.layer.closeAll();}" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "			return false;" & vbCrlf		'禁止提交
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub ImportSave()
	Server.ScriptTimeout = 1200			'脚本超时
	Dim xlsFile : xlsFile = Trim(Request("xlsUrl"))
	Dim tEduYear : tEduYear = HR_CLng(Request("EduYear"))
	Dim ajaxMethod : ajaxMethod = HR_CBool(Request("ajax"))
	Dim strTmp, tmpItemArr, tmpArr, tmpCount : tmpCount = 0
	ErrMsg = ""

	If fso.FileExists(Server.MapPath(xlsFile)) = False Then
		ErrMsg = "<i class='layui-icon'>&#xe69c;</i>Excel数据文件 " & xlsFile & " 不存在！"
	End If
	If tEduYear > Year(Date()) Or tEduYear < 2018 Then
		ErrMsg = "<i class='layui-icon'>&#xe69c;</i>您选择的学年 " & tEduYear & " 不正确！"
	End If

	If ajaxMethod = False Then	'若非异步请求时
		If fso.FileExists(Server.MapPath(xlsFile)) Then
			strTmp = GetHttpStr(apiHost & "/API/ReadExcel.htm?type=1&xlsFile=" & xlsFile, "", 1, 10)
			If HR_IsNull(strTmp) Then
				ErrMsg = "<i class='layui-icon'>&#xe63d;</i>Excel数据文件 " & xlsFile & " 没有记录！"
			Else
				tmpItemArr = Split(strTmp, "@@")
				'ErrMsg = "<i class='layui-icon'>&#xe6af;</i>共有 " & Ubound(tmpItemArr)+1 & "条数据"
			End If
		End If
		tmpHtml = "<style type=""text/css"">" & vbCrlf
		tmpHtml = tmpHtml & "		body,html,div,dl,dt,dd,ul,li,cite,em,tt,cite, form, ol, p, h1, h2 {box-sizing:border-box; font-style:normal;}" & vbCrlf
		tmpHtml = tmpHtml & "		.reload {text-align:center; position:relative; top:50px; display:flex; justify-content:center; flex-direction:column; font-size:18px;}" & vbCrlf
		tmpHtml = tmpHtml & "		.reload p i {font-size:30px; color:#f30; position: relative; top:5px; right:5px;}" & vbCrlf
		tmpHtml = tmpHtml & "		.reload span {color:#f60; position: relative; top:5px; right:5px;}" & vbCrlf
		tmpHtml = tmpHtml & "	</style>" & vbCrlf

		strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
		strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
		Response.Write ReplaceCommonLabel(strHtml)

		Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
		Response.Write "	<div class='reload'><p><i class='layui-icon layui-anim layui-anim-rotate layui-anim-loop'>&#xe63d;</i>正在保存数据，可能需要几分钟，请稍等……</p><span></span></div>" & vbCrlf
		Response.Write "</div>" & vbCrlf

		tmpHtml = "<script type=""text/javascript"">" & vbCrlf
		tmpHtml = tmpHtml & "	layui.use([""form"", ""element""], function(){" & vbCrlf
		tmpHtml = tmpHtml & "		var element = layui.element, form = layui.form;" & vbCrlf
		If HR_IsNull(ErrMsg) = False Then tmpHtml = tmpHtml & "		$("".reload p"").html(""" & ErrMsg & """); return false;" & vbCrlf			'中止

		tmpHtml = tmpHtml & "		$.post(""" & ParmPath & "Setup/ImportSave.html"", {ajax:""True"",EduYear:" & tEduYear & ",xlsUrl:""" & xlsFile & """}, function(res){" & vbCrlf
		tmpHtml = tmpHtml & "			console.log(res);" & vbCrlf
		tmpHtml = tmpHtml & "			$("".reload p"").html(res.errmsg); return false;" & vbCrlf
		tmpHtml = tmpHtml & "		},""json"");" & vbCrlf

		tmpHtml = tmpHtml & "	});" & vbCrlf
		tmpHtml = tmpHtml & "</script>" & vbCrlf
		strHtml = getPageFoot(1)
		strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
		Response.Write ReplaceCommonLabel(strHtml)
		Response.End
	End If
	'Response.End
	'计算总数据量
	If fso.FileExists(Server.MapPath(xlsFile)) Then
		strTmp = GetHttpStr(apiHost & "/API/ReadExcel.htm?type=1&xlsFile=" & xlsFile, "", 1, 10)
		If HR_IsNull(strTmp) Then
			ErrMsg = "<i class='layui-icon'>&#xe69c;</i>Excel数据文件 " & xlsFile & " 没有记录！"
		Else
			'Response.Write strTmp
			tmpItemArr = Split(strTmp, "@@")
			ErrMsg = "<ul>"
			For i = 1 To Ubound(tmpItemArr)		'0行为标题栏
				If HR_IsNull(tmpItemArr(i)) = False And Instr(tmpItemArr(i), "工号") = 0 Then
					tmpArr = Split(tmpItemArr(i), "||")
					If Ubound(tmpArr) >= 4 Then
						Call ChkTeacherKPI(tmpArr(0))		'更新员工业绩表
						Set rsTmp = Server.CreateObject("ADODB.RecordSet")
							rsTmp.Open("Select * From HR_KPI Where scYear=" & tEduYear & " And YGDM=" & HR_Clng(tmpArr(0))), Conn, 1, 3
							If Not(rsTmp.BOF And rsTmp.EOF) Then
								rsTmp("Grade") = tmpArr(4)
								rsTmp.Update
								Conn.Execute("Update HR_KPI_SUM Set Grade='" & tmpArr(4) & "' Where scYear=" & tEduYear & " And YGDM=" & HR_Clng(tmpArr(0)))	'同步学时缓存库
								tmpCount = tmpCount + 1
							End If
						Set rsTmp = Nothing
					Else
						ErrMsg = ErrMsg & "<li>数据格式错误</li>"
						Exit For
					End If
				Else
					ErrMsg = ErrMsg & "<li>" & i + 2 & "、空数据</li>"
				End If
			Next
			ErrMsg = ErrMsg & "<li class=\""count\"">共有 " & tmpCount & " 条数据导入成功！</li>"
			ErrMsg = ErrMsg & "</ul>"
		End If
	End If
	Response.Write "{""err"":false, ""errcode"":500, ""errmsg"":""" & ErrMsg & """, ""icon"":1}"
End Sub
%>