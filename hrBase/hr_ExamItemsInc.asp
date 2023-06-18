<%
Sub RatioForm()
	Dim tmpHtml : SubButTxt = "添加" : ErrMsg = ""
	Dim tTypeID : tTypeID = HR_Clng(Request("TypeID"))
	Dim tStuType : tStuType = Trim(Request("stuType"))
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tItemName, tFieldName, tMaxScore, tUnit, tRatio, arrRatio, tIntro, arrStuType, arrValue
	
	If tItemID = 0 Then
		ErrMsg = ErrMsg & "新增考核项目时不能设置考核系数，<br>请添加成功后在修改时设置！"
		Response.Write GetErrBody(2) : Exit Sub
	End If

	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If rsTmp.BOF And rsTmp.EOF Then
			ErrMsg = ErrMsg & "您要修改的考核项目【ID：" & tmpID & "】不存在！<br>"
		Else
			SubButTxt = "修改"
			tItemID = HR_Clng(rsTmp("ClassID"))
			tItemName = Trim(rsTmp("ClassName"))
			tMaxScore = HR_Clng(rsTmp("MaxScore"))
			tUnit = Trim(rsTmp("Unit"))
			tRatio = Trim(rsTmp("Ratio"))
			tIntro = Trim(rsTmp("Readme"))
		End If
	Set rsTmp = Nothing

	If ErrMsg <> "" Then Response.Write GetErrBody(2) : Exit Sub

	'判断是否有学生类别
	If tStuType <> "" Then
		tStuType = FilterArrNull(tStuType, ",")
		arrStuType = Split(tStuType, ",")
		Redim arrValue(Ubound(arrStuType))
		tRatio = FilterArrNull(tRatio, ",")
		arrRatio = Split(tRatio, ",")
		If Ubound(arrRatio) = Ubound(arrStuType) Then
			For m = 0 To Ubound(arrStuType)
				arrValue(m) = FormatNumber(HR_CDbl(arrRatio(m)), 1, -1)
			Next
		Else
			For m = 0 To Ubound(arrStuType)
				arrValue(m) = 0
			Next
		End If
	Else
		Redim arrValue(0)
		arrValue(0) = tRatio
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {padding: 10px;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding: 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	tmpHtml = "<fieldset class=""layui-elem-field site-demo-button"">" & vbCrlf
	tmpHtml = tmpHtml & "	<legend>" & SubButTxt & " " & tItemName & " 考核系数</legend>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layer-hr-box"">" & vbCrlf
	tmpHtml = tmpHtml & "		<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	If tStuType <> "" Then
		For m = 0 To Ubound(arrStuType)
			tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
			tmpHtml = tmpHtml & "				<label class=""layui-form-label"">" & arrStuType(m) & "：</label>"
			tmpHtml = tmpHtml & "				<div class=""layui-input-inline""><input type=""number"" name=""FieldName" & m + 1 & """ value=""" & arrValue(m) & """ placeholder=""系数值只能为数字"" lay-verify=""number"" class=""layui-input""></div>" & vbCrlf
			tmpHtml = tmpHtml & "				<div class=""layui-form-mid layui-word-aux"">若此项无考核系数，请设置为0</div>" & vbCrlf
			tmpHtml = tmpHtml & "			</div>" & vbCrlf
		Next
		tmpHtml = tmpHtml & "			<input type=""hidden"" name=""StuTypeLen"" value=""" & Ubound(arrStuType) + 1 & """>" & vbCrlf
	Else
		tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "				<label class=""layui-form-label"">考核系数：</label>"
		tmpHtml = tmpHtml & "				<div class=""layui-input-inline""><input type=""number"" name=""FieldName"" value=""" & arrValue(0) & """ placeholder=""系数值只能为数字"" lay-verify=""number"" class=""layui-input""></div>" & vbCrlf
		tmpHtml = tmpHtml & "				<div class=""layui-form-mid""><em class=""hr-help""><i class=""hr-icon"">&#xecfd;</i>若此项无考核系数，请设置为1</em></div>" & vbCrlf
		tmpHtml = tmpHtml & "			</div>" & vbCrlf
	End If
	tmpHtml = tmpHtml & "			<input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""ItemID"" value=""" & tItemID & """><input type=""hidden"" name=""Modify"" value=""True"">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""hr-pop-fix"">" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-inline""><button class=""layui-btn layui-btn-sm"" lay-submit lay-filter=""EditPost"">" & SubButTxt & "</button><button type=""reset"" class=""layui-btn layui-btn-sm layui-btn-primary"">重置</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "		</form>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
	Response.Write tmpHtml

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	strHtml = strHtml & "		form.on(""submit(EditPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "			var loadMsg = layer.load(1,{shade:[0.1, ""#000""]});" & vbCrlf
	strHtml = strHtml & "			$.ajax({type:""post"",url:""" & ParmPath & "ExamItems/SaveRatio.html"", data:$(""#EditForm"").serialize(),timeout:0,dataType:""json"",success:function(result){" & vbCrlf
	strHtml = strHtml & "				var icon=2; if(result.Return){icon=1}" & vbCrlf
	strHtml = strHtml & "				layer.alert(result.reMessge, {icon:icon},function(layero, index){" & vbCrlf
	strHtml = strHtml & "					layer.closeAll();form.render();" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "				},error:function(xhr){console.log(xhr)}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub
Sub SaveRatio()
	Dim StuTypeLen : StuTypeLen = HR_Clng(Request("StuTypeLen"))
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim rsAdd, sqlAdd, tmpJson
	Dim tStudentType, arrStuType, arrValue, tRatio, tmpValue
	ErrMsg = "" : SubButTxt = "修改"
	'校验系数
	If StuTypeLen > 0 Then
		For m = 1 To StuTypeLen
			tRatio = Request("FieldName" & m)
			If HR_IsNumeric(tRatio) = False Then
				ErrMsg = "您输入的值不是数字，请重新输入！"
			End If
			tmpValue = tmpValue & Request("FieldName" & m) & ","
		Next
		tmpValue = FilterArrNull(tmpValue, ",")
	Else
		tmpValue = Request("FieldName")
	End If
	sqlAdd = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsAdd = Server.CreateObject("ADODB.RecordSet")
		rsAdd.Open(sqlAdd), Conn, 1, 3
		If rsAdd.BOF And rsAdd.EOF Then
			ErrMsg = "业绩项目不存在！"
		Else
			tStudentType = Trim(rsAdd("StudentType"))
			rsAdd("Ratio") = tmpValue
			rsAdd.Update
			rsAdd.Close
			Call UpdateItemKPI(tItemID)		'更新项目KPI
		End If
	Set rsAdd = Nothing

	If ErrMsg <> "" Then
		tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}"
	Else
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""考核系数 " & SubButTxt & "成功！"",""ReStr"":""操作成功！""}"
	End If
	Response.Write tmpJson
End Sub

Sub TemplateForm()
	Dim rsList, tmpHtml, strData, tItemName, tTemplate, tExcelFile, tFieldsLen
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tTypeID : tTypeID = HR_Clng(Request("TypeID"))

	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
		End If
	Set rsTmp = Nothing
	Set rsTmp = Conn.Execute("Select Top 1 * From HR_DataModel Where ModelName='" & tTemplate & "'")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tFieldsLen = HR_Clng(rsTmp("FieldsLen"))
			tExcelFile = Trim(rsTmp("ExcelFile"))
		End If
	Set rsTmp = Nothing

	If tExcelFile <> "" Then		'检测模板文件是否存在
		If FSO.FileExists(Server.MapPath(tExcelFile)) Then
			tExcelFile = apiHost & tExcelFile
		End If
	End If


	If ChkTable(tTemplate) Then
		Dim FieldNum
		Set rsList = Server.CreateObject("ADODB.RecordSet")
			rsList.Open("Select * From " & tTemplate & ""), Conn, 1, 1
			FieldNum = rsList.Fields.Count		'字段总数
			strData = "<thead><tr>"
			For i = 4 To rsList.Fields.Count - 2
				strData = strData & "<th>" & rsList.Fields(i).Name & "</th>"
			Next
			strData = strData & "</tr></thead>"
			If Not(rsList.BOF And rsList.EOF) Then
				strData = strData & "<tbody>"
				Do While Not rsList.EOF
					strData = strData & "<tr>"
					For i = 4 To rsList.Fields.Count - 2
						strData = strData & "<td>" & rsList(rsList.Fields(i).Name) & "</td>"
					Next
					strData = strData & "</tr>"
					rsList.MoveNext
				Loop
				strData = strData & "</tbody>"
			End If
		Set rsList = Nothing
		strData = "<fieldset class=""layui-elem-field layui-field-title"" style=""margin-top: 10px;""><legend>" & FieldNum & "</legend></fieldset><table class=""layui-table"">" & strData & "</table>"
	Else
		strData = strData & "<div class=""hr-rows hr-err""><em><i class=""hr-icon"">&#xf071;</i></em><em class=""hr-err-tips""><h3>提示：</h3><h4>没有找到数据模版！</h4></em></div>"
	End If
	tmpHtml = "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:10px;"">" & vbCrlf
	tmpHtml = tmpHtml & "<legend>" & tItemName & "数据模板 " & tTemplate & "</legend>"
	tmpHtml = tmpHtml & "<div class=""xlsData"" id=""xlsData"">" & tExcelFile & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layer-hr-box ExportBox"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div>请点击鼠标右键后选择“另存为”：<br>数据模板：<a=""" & tTemplate & """><input type=""hidden"" name=""ItemID"" id=""ItemID"" value=""" & tItemID & """>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""Export""><button class=""layui-btn"" name=""ExportPost"" id=""ExportPost"">下载Excel模板</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	Response.Write tmpHtml
End Sub
Sub DownExcel()
	Dim strUrl, tmpJson, xlsValue, urlParam
	Dim Template : Template = Trim(ReplaceBadChar(Request("Template")))
	Dim ItemID : ItemID = HR_Clng(Request("ItemID"))
	Dim tItemName : tItemName = Trim(Request("ItemName"))
	If ChkTable("HR_" & Template) Then
		urlParam = "Template=" & Template & "&ItemName=" & tItemName
		strUrl = apiHost & ParmPath & "DownExcel.htm?" & urlParam
		xlsValue = GetHttpPage(strUrl, 1)
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""生成Excel模板成功！<br>请点击鼠标右键后选择“另存为”"",""fileUrl"":""" & xlsValue & """}"
	Else
		tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""发生未知错误！"",""fileUrl"":""""}"
	End If
	Response.Write tmpJson
End Sub

Sub ShowTemp()
	Dim xlsUrl, tSheetName, tFieldID, tUnit, tExcelFile, tFieldsLen, tDescr
	Dim rsList, strData, arr1, strTmp, jsonOBJ, j
	Dim tItemName : tItemName = Trim(Request("itemName"))
	Dim tTemplate : tTemplate = Trim(Request("Temp"))

	Set rsTmp = Conn.Execute("Select Top 1 * From HR_DataModel Where ModelName='" & tTemplate & "'")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tFieldsLen = HR_Clng(rsTmp("FieldsLen"))
			tExcelFile = Trim(rsTmp("ExcelFile"))
			tDescr = Trim(rsTmp("ModelDescr"))
		End If
	Set rsTmp = Nothing

	If tExcelFile <> "" Then		'检测模板文件是否存在
		If FSO.FileExists(Server.MapPath(tExcelFile)) Then
			strTmp = GetHttpPage(apiHost & "/API/ReadExcel.htm?xlsFile=" & tExcelFile, 1)
			If strTmp <> "" Then
				Set jsonOBJ = parseJSON(strTmp)
					If jsonObj.data.length > 0 Then
						tFieldsLen = HR_Clng(jsonObj.data.get(0).fLen)
						strData = "<thead><tr>"
						strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA0) & "</th>"
						strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA1) & "</th>"
						strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA2) & "</th>"
						strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA3) & "</th>"
						strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA4) & "</th>"
						strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA5) & "</th>"
						strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA6) & "</th>"
						If tFieldsLen > 7 Then
							strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA7) & "</th>"
							If tFieldsLen > 8 Then
								strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA8) & "</th>"
								If tFieldsLen > 9 Then
									strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA9) & "</th>"
									If tFieldsLen > 10 Then
										strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA10) & "</th>"
										If tFieldsLen > 11 Then
											strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA11) & "</th>"
											strData = strData & "<th>" & Trim(jsonObj.data.get(0).VA12) & "</th>"
										End If
									End If
								End If
							End If
						End If
						strData = strData & "</tr></thead>"
						strData = strData & "<tbody>"
						For j = 1 To jsonObj.data.length-1
							strData = strData & "<tr>"
							strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA0) & "</td>"
							strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA1) & "</td>"
							strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA2) & "</td>"
							strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA3) & "</td>"
							If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Then
								strData = strData & "<td>" & ConvertNumDate(jsonObj.data.get(j).VA4) & "</td>"
							Else
								strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA4) & "</td>"
							End If
							strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA5) & "</td>"
							strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA6) & "</td>"
							If tFieldsLen > 7 Then
								strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA7) & "</td>"
								If tFieldsLen > 8 Then
									strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA8) & "</td>"
									If tFieldsLen > 9 Then
										strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA9) & "</td>"
										If tFieldsLen > 10 Then
											strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA10) & "</td>"
											If tFieldsLen > 11 Then
												strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA11) & "</td>"
												strData = strData & "<td>" & Trim(jsonObj.data.get(j).VA12) & "</td>"
											End If
										End If
									End If
								End If
							End If
							strData = strData & "</tr>"
						Next
						strData = strData & "</tbody>"
						strData = "<table class=""layui-table"">" & strData & "</table>"
					End If
				Set jsonOBJ = Nothing
			End If
		End If
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {padding: 10px;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding: 5px 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .width_80 {width:80px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.downTips {color:#900;line-height:30px;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.Export {color:#f60;font-size:18px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.Export i {font-size: 18px!important;position: relative;top:2px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	tmpHtml = "<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	tmpHtml = tmpHtml & "	<legend>模板名：" & tTemplate & "</legend>"
	tmpHtml = tmpHtml & "	<div class=""xlsData"" id=""xlsData"">" & strData & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""layer-hr-box Tips"">" & tDescr & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""layer-hr-box ExportBox"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""downTips"">请点击鼠标右键后选择“另存为”</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""Export""><i class=""hr-icon"">&#xf019;</i><a href=""" & tExcelFile & """>" & tExcelFile & "</a></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	Response.Write tmpHtml

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form"", ""laydate"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, laydate = layui.laydate;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#ExportPost"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var ItemName = $(""#ItemName"").val(), Template = $(""#Template"").val();" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "ExamItems/DownExcel.html"", {ItemName:ItemName, Template:Template}, function(result){" & vbCrlf
	tmpHtml = tmpHtml & "				if(result.Return){" & vbCrlf
	tmpHtml = tmpHtml & "					var downDoc = ""<div class='downBar'><h4>"" + result.reMessge + ""</h3><em><i class='hr-icon hr-icon-top'>&#xf019;</i>：<a href="" + result.fileUrl + "">"" + result.fileUrl + ""</a></em></div>"";" & vbCrlf
	tmpHtml = tmpHtml & "					layer.alert(downDoc, {icon:1,area:'500px',anim:1});" & vbCrlf
	tmpHtml = tmpHtml & "				}else{" & vbCrlf
	tmpHtml = tmpHtml & "					layer.alert(result.reMessge, {icon:2});" & vbCrlf
	tmpHtml = tmpHtml & "				}" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "			return false;" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub SetupField()
	Dim tmpID : tmpID = HR_Clng(Request("ItemID"))
	Dim tItemName, tTemplate, tFieldHead, tFieldLen, strArr, arrHead, arrLen
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tmpID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			tFieldHead = Trim(rsTmp("FieldHead"))
			tFieldLen = HR_Clng(rsTmp("FieldLen"))
		End If
	Set rsTmp = Nothing
	If Instr(tFieldHead, ",") = 0 Then tFieldHead = ""
	If tFieldHead <> "" Then
		arrHead = Split(tFieldHead, ",")
		If Ubound(arrHead) <> tFieldLen-1 Then Redim Preserve arrHead(tFieldLen-1)
		strArr = ""
		For i = 0 To Ubound(arrHead)
			If i > 0 Then strArr = strArr & ","
			strArr = strArr & "{""ItemID"":" & tmpID & ",""ID"":" & i + 1 & ", ""Title"":""" & Trim(arrHead(i)) & """}"
		Next
	Else
		Redim arrHead(tFieldLen-1)
		For i = 0 To Ubound(arrHead)
			If i > 0 Then strArr = strArr & ","
			strArr = strArr & "{""ItemID"":" & tmpID & ",""ID"":" & i + 1 & ", ""Title"":""""}"
		Next
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	'tmpHtml = tmpHtml & "		.layui-elem-field legend {font-size:18px;}" & vbCrlf
	'tmpHtml = tmpHtml & "		.layui-table thead tr {background-color: #eee;color: #000;text-align:center;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)

	tmpHtml = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	tmpHtml = tmpHtml & "	<fieldset class=""layui-elem-field layui-field-title""><legend>编辑 " & tItemName & " 字段标题</legend></fieldset>"
	'tmpHtml = tmpHtml & "<div class=""hr-shrink-x10"">" & vbCrlf
	tmpHtml = tmpHtml & "	<table class=""layui-table"" id=""EditTable"" lay-filter=""EditTable""></table>" & vbCrlf
	'tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	Response.Write tmpHtml

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table;" & vbCrlf
	strHtml = strHtml & "		var colsData = [{type:'numbers',unresize:true,align:'center',width:60,title:'序号'},{field:'Title',edit:'text',title:'字段标题'}];" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	strHtml = strHtml & "		table.render({" & vbCrlf
	strHtml = strHtml & "			elem: ""#EditTable"", limit:20, skin:""line"" ,cols: [colsData] ,data:[" & strArr & "]" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		table.on(""edit(EditTable)"", function(obj){" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "ExamItems/SaveField.html"",{ItemID:obj.data.ItemID, ID:obj.data.ID, Title:obj.value}, function(reData){ });" & vbCrlf
	strHtml = strHtml & "			table.reload(""EditTable"");" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub SaveField()
	Dim tResult, tmpJson, strTmp, tFieldHead, tFieldLen, arrHead, strArr
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tTitle : tTitle = Trim(ReplaceBadChar(Request("Title")))

	tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""字段标题修改失败！"",""ReStr"":""操作失败！""}"
	If tItemID > 0 Then
		Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				tFieldHead = Trim(rsTmp("FieldHead"))
				tFieldLen = HR_Clng(rsTmp("FieldLen"))
				If Instr(tFieldHead, ",") = 0 Then tFieldHead = ""
				If tFieldHead <> "" Then
					arrHead = Split(tFieldHead, ",")
					If Ubound(arrHead) <> tFieldLen-1 Then Redim Preserve arrHead(tFieldLen-1)
				Else
					Redim arrHead(tFieldLen-1)
				End If
				strArr = ""
				For i = 0 To Ubound(arrHead)
					If i > 0 Then strArr = strArr & ","
					If tmpID = i + 1 Then
						strArr = strArr & Trim(tTitle)
					Else
						strArr = strArr & Trim(arrHead(i))
					End If
				Next
				Conn.Execute("Update HR_Class Set FieldHead='" & strArr & "' Where ClassID=" & tItemID)
				tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""字段标题修改成功！"",""ReStr"":""操作成功！""}"
			End If
		Set rsTmp = Nothing
	End If
	Response.Write tmpJson
End Sub

Sub EditGrade()
	Dim tClassID : tClassID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tClassName
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tClassID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tClassName = rsTmp("ClassName")
		End If
	Set rsTmp = Nothing

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	Response.Write "	<legend>" & tClassName & " 等级管理</legend>" & vbCrlf
	Response.Write "	<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "ExamItems/GradeData.html',where:{ItemID:" & tClassID & ",ID:" & tmpID & "},limit:0,text:{none:'等级尚未添加'},id:'GradeTable'}"" lay-filter=""GradeTable"">" & vbCrlf
	Response.Write "		<thead><tr>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'ID',unresize:true,align:'center',width:60}"">序号</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'Grade',edit:'text'}"">等　级</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'Ratio',edit:'text',align:'center',width:70}"">分值</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'Unit',align:'center',width:70}"">单位</th>" & vbCrlf
	Response.Write "			<th lay-data=""{align:'center',unresize:true,width:90,toolbar: '#GradeBar'}"">操作</th>" & vbCrlf
	Response.Write "		</tr></thead>" & vbCrlf
	Response.Write "	</table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""GradeBar"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-danger"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf
	Response.Write "<div class=""hr-shrink-x10""></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-pop-fix""><button class=""layui-btn layui-btn-sm layui-btn-normal"" id=""AddBtn"" title=""添加等级"" data-cid=""" & tClassID & """ data-id=""" & tmpID & """><i class=""layui-icon"">&#xe654;</i>添加</button><button class=""layui-btn layui-btn-sm"" id=""refresh"" title=""刷新""><i class=""hr-icon"">&#xf343;</i></button></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "	layui.use([""form"", ""table"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var form = layui.form, table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element;" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "		table.on(""edit(GradeTable)"", function(obj){" & vbCrlf
	strHtml = strHtml & "			var loadMsg = layer.load(1,{shade:[0.2, ""#000""]});" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "ExamItems/SaveGrade.html"",{ID:obj.data.ID,Grade:obj.value,Field:obj.field}, function(reData){ layer.close(loadMsg); });" & vbCrlf
	'strHtml = strHtml & "			window.location.reload();" & vbCrlf		'刷新
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		table.on(""tool(GradeTable)"", function(obj){" & vbCrlf
	strHtml = strHtml & "			var data = obj.data;" & vbCrlf
	strHtml = strHtml & "			if(obj.event === ""del""){" & vbCrlf
	strHtml = strHtml & "				layer.confirm(""真的删除选中的等级吗？"", {icon:3, title:""删除提醒""}, function(index){" & vbCrlf
	strHtml = strHtml & "					var loadMsg = layer.load(1,{shade:[0.2, ""#000""]});" & vbCrlf
	strHtml = strHtml & "					$.getJSON(""" & ParmPath & "ExamItems/DelGrade.html"",{ID:data.ID, LevelID:data.LevelID}, function(reData){ layer.close(loadMsg);window.location.reload(); });" & vbCrlf
	strHtml = strHtml & "					obj.del();" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#AddBtn"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			var cid = $(this).data(""cid""), id = $(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "			layer.prompt({title:""添加等级""},function(value, index, elem){" & vbCrlf
	strHtml = strHtml & "				var loadMsg = layer.load(1,{shade:[0.2, ""#000""]});" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "ExamItems/SaveGrade.html"",{ItemID:cid,LevelID:id,Grade:value,Field:""AddNew""}, function(reData){ layer.close(loadMsg);window.location.reload(); });" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#refresh"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			window.location.reload();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	
	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub
Sub GradeData()
	Dim vCount, vMSG, tmpJson, rsGet, sqlGet
	Dim tClassID : tClassID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	tmpJson = ""
	sqlGet = "Select * From HR_ItemGrade Where ClassID=" & tClassID & " And LevelID=" & tmpID
	sqlGet = sqlGet & " Order By ID ASC"
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			i = 0
			Do While Not rsGet.EOF
				If i > 0 Then tmpJson = tmpJson & ","
				tmpJson = tmpJson & "{""ID"":" & rsGet("ID") & ",""LevelID"":" & rsGet("LevelID") & ",""ItemID"":" & rsGet("ClassID") & ""
				tmpJson = tmpJson & ",""Grade"":""" & Trim(rsGet("Grade")) & """,""Ratio"":""" & FormatNumber(rsGet("Ratio"), 1, -1) & """,""Unit"":""" & Trim(rsGet("Unit")) & """,""Intro"":""" & HR_HTMLEncode(nohtml(rsGet("Intro"))) & """}"
				rsGet.MoveNext
				i = i + 1
			Loop
		End If
		vCount = rsGet.Recordcount
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""暂无数据"",""count"":" & vCount & ",""data"":[" & tmpJson
	tmpJson = tmpJson & "],""limit"":""0"",""page"":""0""}"
	Response.Write tmpJson
End Sub
Sub SaveGrade()
	Dim tGrade, tResult, tmpJson, strTmp, rsSave
	Dim tField : tField = Trim(ReplaceBadChar(Request("Field")))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tLevelID : tLevelID = HR_Clng(Request("LevelID"))
	Dim tClassID : tClassID = HR_Clng(Request("ItemID"))
	tGrade = Trim(Request("Grade"))
	tResult = False
	If tmpID > 0 And tField = "Grade" Then
		Conn.Execute("Update HR_ItemGrade Set Grade='" & ReplaceBadChar(tGrade) & "' Where ID=" & tmpID)
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""等级 修改成功！"",""ReStr"":""操作成功！""}"
	ElseIf tmpID > 0 And tField = "Ratio" Then
		Conn.Execute("Update HR_ItemGrade Set Ratio=" & HR_CDbl(tGrade) & " Where ID=" & tmpID)
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""等级系数 修改成功！"",""ReStr"":""操作成功！""}"
	Else
		If tLevelID > 0 And tClassID > 0 And tField = "AddNew" Then
			Set rsSave = Server.CreateObject("ADODB.RecordSet")
				rsSave.Open("Select * From HR_ItemGrade"), Conn, 1 , 3
				rsSave.AddNew
				rsSave("ID") = GetNewID("HR_ItemGrade", "ID")
				rsSave("Grade") = ReplaceBadChar(tGrade)
				rsSave("TypeID") = GetTypeName("HR_Class", "ClassType", "ClassID", tClassID)
				rsSave("ClassID") = tClassID
				rsSave("LevelID") = tLevelID
				rsSave("Ratio") = 0
				rsSave("Unit") = GetTypeName("HR_Class", "Unit", "ClassID", tClassID)
				rsSave("LevelID") = tLevelID
				rsSave.Update
			Set rsSave = Nothing
			tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""等级 添加成功！"",""ReStr"":""操作成功！""}"
		Else
			tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""等级 添加失败！"",""ReStr"":""操作失败！""}"
		End If
	End If
	Call RecordFrontLog(tClassID, "编辑等级", "修改等级[级别ID：" & tLevelID & "]，FieldName：HR_ItemGrade", True, "Save")
	Call UpdateItemKPI(tClassID)	'更新项目KPI
	Response.Write tmpJson
End Sub

Sub DelGrade()
	Dim tmpJson
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tLevelID : tLevelID = HR_Clng(Request("LevelID"))
	Dim tmpClassID : tmpClassID = GetTypeName("HR_ItemGrade", "ClassID", "ID", tmpID)
	Conn.Execute("Delete From HR_ItemGrade Where ID=" & tmpID)
	Call UpdateItemKPI(tmpClassID)	'更新项目KPI
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""等级 删除成功！"",""ReStr"":""删除成功！""}"
	Response.Write tmpJson
End Sub

Sub UpdateItemKPI(sItemID)
	Dim rsSub, sqlSub, tClassName, tTemplate, tSheetName, strSub
	Dim rsUpdate
	If HR_CLng(sItemID) > 0 Then
		sqlSub = "Select * From HR_Class Where Child=0 And ClassID=" & HR_CLng(sItemID)
		Set rsSub = Conn.Execute(sqlSub)
			If Not(rsSub.BOF And rsSub.EOF) Then
				tClassName = rsSub("ClassName")
				tTemplate = Trim(rsSub("Template"))
				tSheetName = "HR_Sheet_" & sItemID
				If ChkTable(tSheetName) Then
					Set rsUpdate = Conn.Execute("Select VA1 From " & tSheetName & " Where scYear=" & DefYear & " Group By VA1")
						If Not(rsUpdate.BOF And rsUpdate.EOF) Then
							Do While Not rsUpdate.EOF
								Call ChkTeacherKPI(rsUpdate("VA1"))	'添加员工信息至业绩表
								Call UpdateTeacherKPI(sItemID, rsUpdate("VA1"), "")	'更新本项目员工统计数据
								Call UpdateTeacherTotalKPI(rsUpdate("VA1"))	'更新员工总计数据
								rsUpdate.MoveNext
							Loop
						Else
							strSub = strSub & "业绩考核项目" & tClassName & "中没有数据！<br>"
						End If
					Set rsUpdate = Nothing
				Else
					strSub = strSub & "未找到业绩考核项目" & tClassName & "数据表！<br>"
				End If
			Else
				strSub = strSub & "业绩考核项目[ID:" & sItemID & "]不存在！<br>"
			End If
		Set rsSub = Nothing
	End If
End Sub
%>