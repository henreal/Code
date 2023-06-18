<%
Sub Preview()
	Dim tmpHtml : SubButTxt = "添加" : ErrMsg = ""
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tItemName, tTemplate, tUnit, tSheetName, tFieldHead, lenField, tArrHead

	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			tUnit = Trim(rsTmp("Unit"))
			lenField = HR_Clng(rsTmp("FieldLen"))
			tFieldHead = Trim(rsTmp("FieldHead"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = ErrMsg & "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing
	If HR_IsNull(tTemplate) Then ErrMsg = ErrMsg & "业绩考核项目不存在！<br>"
	If ChkTable(tSheetName) = False Then ErrMsg = ErrMsg & "未找到数据表 " & tSheetName & "！<br>"
	If ErrMsg<>"" Then Response.Write GetErrBody(0) : Exit Sub

	If tFieldHead <> "" Then
		tFieldHead = FilterArrNull(tFieldHead, ",")
		tArrHead = Split(tFieldHead, ",")
		If Ubound(tArrHead) <> lenField Then Redim Preserve tArrHead(lenField)
	Else
		Redim tArrHead(lenField)
	End If
	Dim tStuType, tYGDM, tYGXM, tKSMC, tXMJP, tPRZC, tXZZW, tYGXB, arrField, tAttach, tArrAttach, tPassed
	Dim tmpPassed : tmpPassed = False
	
	Set rsTmp = Conn.Execute("Select * From " & tSheetName & " Where ID=" & tmpID)
		Redim arrField(lenField-1)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tStuType = Trim(rsTmp("StudentType"))
			tYGDM = Trim(rsTmp("VA1"))
			tmpPassed = HR_CBool(rsTmp("Passed"))
			For i = 0 To lenField-1
				arrField(i) = rsTmp("VA" & i)
			Next
			tAttach = Trim(rsTmp("Explain"))
			tPassed = GetShowBit(HR_CBool(rsTmp("Passed")), 1)
		End If
	Set rsTmp = Nothing
	Set rsTmp = Conn.Execute("Select Top 1 * From HR_Teacher Where YGDM='" & tYGDM & "'")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tYGXM = Trim(rsTmp("YGXM"))
			tKSMC = Trim(rsTmp("KSMC"))
			tXMJP = Trim(rsTmp("XMJP"))
			tPRZC = Trim(rsTmp("PRZC"))
			tXZZW = Trim(rsTmp("XZZW"))
			tYGXB = Trim(rsTmp("YGXB"))
		End If
	Set rsTmp = Nothing

	Call UpdateTeacherKPI(tItemID, tYGDM, tStuType)	'更新本项目员工统计数据
	Call UpdateTeacherTotalKPI(tYGDM)	'更新员工总计数据

	tAttach = FilterArrNull(tAttach, "|")		'取附件
	Dim tmpExtname, AttachNum : AttachNum = 0
	If HR_IsNull(tAttach) = False Then
		tArrAttach = Split(tAttach, "|")
		AttachNum = Ubound(tArrAttach) + 1
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {padding: 10px;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding: 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-color-true {color:#080;} .hr-color-false {color:#F30;}" & vbCrlf
	tmpHtml = tmpHtml & "		#PreviewBox {display: flex;align-items: center;flex-wrap: wrap;}" & vbCrlf
	tmpHtml = tmpHtml & "		#PreviewBox .prevImg {width: 92px;height: 92px;margin: 0 10px 10px 0;background-repeat: no-repeat;background-position: center;background-size: auto 92px;}" & vbCrlf

	tmpHtml = tmpHtml & "		#AttachBar {line-height:37px;display:flex;align-items:center;flex-wrap:wrap;width:100%;border:1px solid #ddd;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		#AttachBar em {min-height:60px; line-height:50px; cursor: pointer; padding:15px 0 0 15px;color:#39c;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		#AttachBar em i {font-size:46px;position:relative;top:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		#AttachBar em tt {display:none;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)

	strHtml = "<fieldset class=""layui-elem-field layui-field-title""><legend>" & tItemName & tStuType &" 课程信息</legend>" & vbCrlf
	strHtml = strHtml & "	<div class=""layui-form layer-hr-box"">" & vbCrlf
	strHtml = strHtml & "	<table class=""layui-table"">" & vbCrlf
	strHtml = strHtml & "		<colgroup><col width=""120""><col><col width=""120""><col></colgroup>" & vbCrlf
	strHtml = strHtml & "		<tbody>" & vbCrlf
	strHtml = strHtml & "			<tr><td style=""text-align:right;"">教师姓名：</td><td>" & arrField(2) & "</td><td style=""text-align:right;"">工　号：</td><td>" & arrField(1) & "</td></tr>" & vbCrlf
	strHtml = strHtml & "			<tr><td style=""text-align:right;"">科　室：</td><td>" & tKSMC & "</td><td style=""text-align:right;"">职　称：</td><td>" & tPRZC & "</td></tr>" & vbCrlf
	strHtml = strHtml & "			<tr><td style=""text-align:right;"">审核状态：</td><td>" & tPassed & "</td><td style=""text-align:right;"">附件数：</td><td>" & AttachNum & "</td></tr>" & vbCrlf
	If tTemplate = "TempTableB" Then
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">" & tArrHead(4) & "：</td><td>" & Trim(arrField(4)) & "</td><td style=""text-align:right;"">学　期：</td><td>" & arrField(5) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">" & tArrHead(3) & "：</td><td colspan=""3"">" & arrField(3) & " " & tUnit & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">" & tArrHead(5) & "：</td><td colspan=""3"">" & arrField(5) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">备　注：</td><td colspan=""3"">" & arrField(6) & "</td></tr>" & vbCrlf
	ElseIf tTemplate = "TempTableC" Then
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">日　期：</td><td>" & FormatDate(ConvertNumDate(arrField(4)), 4) & "</td><td style=""text-align:right;"">学期学年：</td><td>" & Trim(arrField(5)) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">数　值：</td><td colspan=""3"">" & arrField(3) & " " & tUnit & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">项目名称：</td><td colspan=""3"">" & arrField(6) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">备　注：</td><td colspan=""3"">" & arrField(7) & "</td></tr>" & vbCrlf
	ElseIf tTemplate = "TempTableD" Then
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">日　期：</td><td>" & FormatDate(ConvertNumDate(arrField(4)), 4) & "</td><td style=""text-align:right;"">学期学年：</td><td>" & Trim(arrField(5)) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">数　值：</td><td colspan=""3"">" & arrField(3) & " " & tUnit & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">项目名称：</td><td colspan=""3"">" & arrField(6) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">备　注：</td><td colspan=""3"">" & arrField(7) & "</td></tr>" & vbCrlf
	ElseIf tTemplate = "TempTableE" Then
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">" & tArrHead(4) & "：</td><td>" & FormatDate(ConvertNumDate(arrField(4)), 4) & "</td><td style=""text-align:right;"">学期学年：</td><td>" & Trim(arrField(5)) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">" & tArrHead(3) & "：</td><td>" & arrField(3) & " " & tUnit & "</td><td style=""text-align:right;"">" & tArrHead(6) & "：</td><td>" & arrField(6) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">" & tArrHead(7) & "：</td><td>" & arrField(7) & "</td><td style=""text-align:right;"">" & tArrHead(8) & "：</td><td>" & arrField(8) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">备　注：</td><td colspan=""3"">" & arrField(9) & "</td></tr>" & vbCrlf
	ElseIf tTemplate = "TempTableF" Then
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">" & tArrHead(3) & "：</td><td>" & Trim(arrField(3)) & " " & tUnit & "</td><td style=""text-align:right;"">" & tArrHead(4) & "：</td><td>" & Trim(arrField(4)) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">" & tArrHead(5) & "：</td><td colspan=""3"">" & arrField(5) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">" & tArrHead(6) & "：</td><td>" & arrField(6) & "</td><td style=""text-align:right;"">" & tArrHead(7) & "：</td><td colspan=""3"">" & arrField(7) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">备　注：</td><td colspan=""3"">" & arrField(8) & "</td></tr>" & vbCrlf
	ElseIf tTemplate = "TempTableG" Then
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">" & tArrHead(3) & "：</td><td>" & Trim(arrField(3)) & " " & tUnit & "</td><td style=""text-align:right;"">学期学年：</td><td>" & Trim(arrField(4)) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">项目名称：</td><td colspan=""3"">" & arrField(5) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">备　注：</td><td colspan=""3"">" & arrField(6) & "</td></tr>" & vbCrlf
	Else
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">日　期：</td><td>" & FormatDate(ConvertNumDate(arrField(4)), 4) & "</td><td style=""text-align:right;"">周　次：</td><td>" & arrField(5) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">星　期：</td><td>" & arrField(6) & "</td><td style=""text-align:right;"">节　次：</td><td>" & arrField(7) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">学　时：</td><td>" & arrField(3) & " " & tUnit & "</td><td style=""text-align:right;"">校(院)区：</td><td>" & arrField(11) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">课程名称：</td><td colspan=""3"">" & arrField(8) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">授课内容：</td><td colspan=""3"">" & arrField(9) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">授课对象：</td><td colspan=""3"">" & arrField(10) & "</td></tr>" & vbCrlf
		strHtml = strHtml & "			<tr><td style=""text-align:right;"">授课教室：</td><td colspan=""3"">" & arrField(12) & "</td></tr>" & vbCrlf
	End If
	'strHtml = strHtml & "			<tr><td style=""text-align:right;"">说　明：</td><td colspan=""3"">" & tExplain & "</td></tr>" & vbCrlf
	strHtml = strHtml & "		</tbody>" & vbCrlf
	strHtml = strHtml & "	</table>" & vbCrlf
	strHtml = strHtml & "	</div>" & vbCrlf
	If AttachNum > 0 Then
		strHtml = strHtml & "	<div class=""layui-form layer-hr-box"">" & vbCrlf
		strHtml = strHtml & "		<div class=""AttachTitle"">附件：<div>" & vbCrlf
		strHtml = strHtml & "		<div class=""layui-upload-list"" id=""AttachBar"">" & vbCrlf
		For i = 0 To Ubound(tArrAttach)
			tmpExtname = Right(Trim(tArrAttach(i)), Len(Trim(tArrAttach(i))) - inStr(Trim(tArrAttach(i)), "."))
			If HR_IsNull(tArrAttach(i)) = False Then
				If FoundInArr(strExtname, tmpExtname, ",") Then		'判断文件扩展名是否正确
					strHtml = strHtml & "<em class=""fileItem""><span title=""" & Trim(tArrAttach(i)) & """><i class=""hr-icon"">" & GetAttachIcon(tmpExtname) & "</i></span><tt>删除</tt></em>"
				End If
			End If
		Next
		strHtml = strHtml & "		</div>" & vbCrlf
		strHtml = strHtml & "	</div>" & vbCrlf
	End If

	strHtml = strHtml & "</fieldset>" & vbCrlf

	Response.Write ReplaceCommonLabel(strHtml)

	strHtml = "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "		$(""#AttachBar em span"").on(""click"",function(){" & vbCrlf		'预览附件
	strHtml = strHtml & "			parent.layer.open({type:2,content:""" & ParmPath & "Course/viewAttach.html?url="" + $(this).attr(""title""),title:[""预览附件"",""font-size:16""],area:[""80%"", ""86%""],maxmin:true});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#Affirm"").on(""click"",function(){" & vbCrlf		'确认提交
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "CourseProof/Affirm.html"",{ItemID:" & tItemID & ", ID:" & tmpID & "}, function(reData){" & vbCrlf
	strHtml = strHtml & "				layer.msg(reData.reMessge,{icon:1,btn:""关闭"",time:0});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#applyModify"").on(""click"",function(){" & vbCrlf		'申请修改
	strHtml = strHtml & "			parent.layer.open({type:2,id:""applyWin"",content:""" & ParmPath & "CourseProof/applyModify.html?ItemID=" & tItemID & "&ID=" & tmpID & """, title:[""申请修改"",""font-size:16""], area:[""630px"", ""350px""]});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
	
Sub winSelectTeacher()
	Dim tmpListTeacher, rsSearch
	Set rsSearch = Conn.Execute("Select Top 50 * From HR_Teacher Where YGDM<>''")
		If Not(rsSearch.BOF And rsSearch.EOF) Then
			Do While Not rsSearch.EOF
				tmpListTeacher = tmpListTeacher & "<em data-ygdm=""" & rsSearch("YGDM") & """ title=""" & rsSearch("KSMC") & """><span>" & rsSearch("YGXM") & "</span></em>"
				rsSearch.MoveNext
			Loop
		End If
	Set rsSearch = Nothing
	strHtml = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	'strHtml = strHtml & "	<form class=""layui-form layui-form-pane"" id=""SearchForm"" name=""SearchForm"" lay-filter=""SearchForm"" action="""">" & vbCrlf
	strHtml = strHtml & "	<div class=""layui-form-item"">" & vbCrlf
	strHtml = strHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">搜索教师：</label>" & vbCrlf
	strHtml = strHtml & "			<div class=""layui-input-inline""><input name=""soTeacher"" type=""text"" id=""soTeacher"" autocomplete=""off"" class=""layui-input""></div>" & vbCrlf
	strHtml = strHtml & "		</div>" & vbCrlf
	strHtml = strHtml & "		<div class=""layui-inline""><span class=""layui-btn soBtn"">搜索</span></div>" & vbCrlf
	strHtml = strHtml & "	</div>" & vbCrlf
	strHtml = strHtml & "	<div id=""ListTeacher"" class=""listBox"">" & tmpListTeacher & "</div>" & vbCrlf
	'strHtml = strHtml & "	</form>" & vbCrlf
	strHtml = strHtml & "</div>" & vbCrlf
	Response.Write strHtml
End Sub
Sub winTeacherData()
	Dim soKey : soKey = Trim(ReplaceBadChar(Request("soKey")))
	soKey = Replace(soKey, chr(9), "")
	Dim soDeptID : soDeptID = HR_Clng(Request("soksdm"))
	Dim rsSearch, sqlSearch, tmpListData
	sqlSearch = "Select Top 100 * From HR_Teacher Where YGDM<>''"
	If HR_IsNull(soKey) = False Then
		If HR_Clng(soKey) > 0 Then
			sqlSearch = sqlSearch & " And YGDM='" & soKey & "'"
		Else
			sqlSearch = sqlSearch & " And (YGXM like '%" & soKey & "%' Or XMJP='" & soKey & "')"
		End If
	End If
	If soDeptID > 0 Then sqlSearch = sqlSearch & " And KSDM=" & soDeptID

	Set rsSearch = Conn.Execute(sqlSearch)
		If Not(rsSearch.BOF And rsSearch.EOF) Then
			Do While Not rsSearch.EOF
				tmpListData = tmpListData & "<em data-ygdm=""" & rsSearch("YGDM") & """ title=""工号：" & rsSearch("YGDM") & "，科室：" & rsSearch("KSMC") & """><span>" & rsSearch("YGXM") & "</span></em>"
				rsSearch.MoveNext
			Loop
		End If
	Set rsSearch = Nothing
	Response.Write tmpListData
End Sub

Sub Passed()
	Server.ScriptTimeout=600		'10分钟
	Dim tmpJson, sqlGet, rsGet
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = Trim(ReplaceBadChar(Request("ID")))
	Dim PassType : PassType = HR_Clng(Request("type"))
	tmpID = FilterArrNull(tmpID, ",")

	Dim tItemName, tTemplate, tUnit, tSheetName
	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			tUnit = Trim(rsTmp("Unit"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = ErrMsg & "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing
	If ChkTable(tSheetName) = False Then ErrMsg = ErrMsg & "未找到数据表 " & tSheetName & "！<br>"
	If HR_IsNull(tmpID) Then ErrMsg = ErrMsg & "您还没有选择审核的业绩记录！<br>"
	If ErrMsg<>"" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}" : Exit Sub
	End If
	Dim SentMsg, tUpKPI : ErrMsg = ""
	sqlGet = "Select * From " & tSheetName & " Where ID in(" & tmpID & ")"
	If PassType = 1 Then		'单个审核
		Set rsGet = Server.CreateObject("ADODB.RecordSet")
			rsGet.Open("Select * From " & tSheetName & " Where ID=" & tmpID), Conn, 1, 3
			If Not(rsGet.BOF And rsGet.EOF) Then
				If HR_CBool(rsGet("Passed")) Then
					ErrMsg = rsGet("VA2") & "老师的本条课程业绩已经审核，不用重复审核！"
					If UserRank > 1 Then			'超管可取消审核
						rsGet("Passed") = HR_False
						rsGet.Update
						ErrMsg = rsGet("VA2") & "：您在考核项目：" & tItemName & "[序号：" & rsGet("VA0") & "]中的课程业绩已取消审核，您可以修改相关内容了！"
						Call SendMessage(0, tItemID, tmpID, rsGet("VA1"), "您的课程业绩 已取消审核，可修改相关内容", ErrMsg, "")
						ErrMsg = rsGet("VA2") & "老师的课程业绩已取消审核！"
					End If
				Else
					rsGet("Passed") = HR_True
					rsGet.Update
					ErrMsg = rsGet("VA2") & "：您在考核项目：" & tItemName & "[序号：" & rsGet("VA0") & "]中的课程业绩审核成功！"
					Call SendMessage(0, tItemID, tmpID, rsGet("VA1"), "您的课程业绩 审核通过！", ErrMsg, "")
					ErrMsg = rsGet("VA2") & "老师的课程业绩审核成功！"
				End If
				Call ChkTeacherKPI(rsGet("VA1"))	'添加员工信息至业绩表
				tUpKPI = UpdateTeacherKPI(tItemID, rsGet("VA1"), Trim(rsGet("StudentType")))	'更新本项目员工统计数据
				tUpKPI = UpdateTeacherTotalKPI(rsGet("VA1"))	'更新员工总计数据
			Else
				ErrMsg = "课程业绩没有找到！"
			End If
		Set rsGet = Nothing
		
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作完成！""}"
		Response.Write tmpJson : Exit Sub					'执行完毕后退出
	End If

	Set rsGet = Server.CreateObject("ADODB.RecordSet")		'批量审核
		rsGet.Open sqlGet, Conn, 1, 3
		If Not(rsGet.BOF And rsGet.EOF) Then
			m = 0
			Do While Not rsGet.EOF
				If HR_CBool(rsGet("Passed")) Then
					ErrMsg = rsGet("VA2") & "老师的本条课程业绩已经审核，不用重复审核！"
				Else
					rsGet("Passed") = HR_True
					If HR_Clng(rsGet("UserID")) = 0 Then rsGet("UserID") = UserID
					rsGet.Update
					ErrMsg = rsGet("VA1") & "：您在考核项目：" & tItemName & "[序号：" & rsGet("VA0") & "]中的课程业绩审核成功！<a href=""" & ParmPath & "Course.html?ItemID=" & tItemID & "&SearchWord=" & rsGet("VA1") & """>【查看】</a>"
					SentMsg = SendMessage(0, tItemID, tmpID, rsGet("VA1"), "您的课程业绩 审核通过", ErrMsg, 0)
					ErrMsg = rsGet("VA2") & "老师的课程业绩审核成功！"
				End If
				Call ChkTeacherKPI(rsGet("VA1"))	'添加员工信息至业绩表
				tUpKPI = UpdateTeacherKPI(tItemID, rsGet("VA1"), "")	'更新本项目员工统计数据
				tUpKPI = UpdateTeacherTotalKPI(rsGet("VA1"))	'更新员工总计数据
				rsGet.MoveNext
				m = m + 1
			Loop
		Else
			tmpJson = "没有找到课程记录！" & tmpID
		End If
	Set rsGet = Nothing

	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & HR_Clng(m) & " 条数据审核完成！"",""ReStr"":""操作完成！""}"
	Response.Write tmpJson
End Sub

Sub AntiPass()		'批量反审核

	Server.ScriptTimeout=600		'10分钟

	Dim tmpID : tmpID = Trim(ReplaceBadChar(Request("ID")))
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))

	tmpID = FilterArrNull(tmpID, ",")

	Dim tItemName, tSheetName
	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing

	Dim sqlGet, rsGet
	sqlGet = "Select * From " & tSheetName & " Where ID in(" & tmpID & ") And Passed=" & HR_True
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
	rsGet.Open sqlGet, Conn, 1, 3
		If Not(rsGet.BOF And rsGet.EOF) Then
			m = 0
			Do While Not rsGet.EOF
				If HR_Clng(rsGet("UserID")) = 0 Then rsGet("UserID") = UserID
				rsGet("Passed") = HR_False
				rsGet.Update
				ErrMsg = rsGet("VA1") & "：您在考核项目：" & tItemName & "[序号：" & rsGet("VA0") & "]中的课程业绩已取消审核！<a href=""" & ParmPath & "Course.html?ItemID=" & tItemID & "&SearchWord=" & rsGet("VA1") & "&ID=" & rsGet("ID") & """>【查看】</a>"
				'Call SendMessage(0, rsGet("VA1"), "您在“" & tItemName & "”中的课程业绩 已取消审核", ErrMsg, 0)
				ErrMsg = rsGet("VA2") & "老师的课程业绩已反审核！<br>"
				Call UpdateTeacherKPI(tItemID, rsGet("VA1"), "")	'更新本项目员工统计数据
				Call UpdateTeacherTotalKPI(rsGet("VA1"))	'更新员工总计数据
				rsGet.MoveNext
				m = m + 1
			Loop
			ErrMsg = "共有 " & HR_Clng(m) & " 条数据已反审核！其他的尚未审核，无须反审核！"
		Else
			ErrMsg = "您反审核的课程业绩尚未审核，无须反审核操作！"
		End If
	Set rsGet = Nothing
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作完成！""}"
End Sub

Sub Attach()
	Dim tmpHtml : SubButTxt = "添加" : ErrMsg = ""
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tItemName, tTemplate, tUnit, tSheetName, strPic, inputPic, picItem
	Dim isPassed : isPassed = False

	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			tUnit = Trim(rsTmp("Unit"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = ErrMsg & "业绩考核项目不存在！<br>"
			Response.Write GetErrBody(0) : Exit Sub
		End If
	Set rsTmp = Nothing
	If ChkTable(tSheetName) = False Then
		ErrMsg = ErrMsg & tItemName & " 数据表未建立，请联系管理员！<br />"
		Response.Write GetErrBody(2) : Exit Sub
	End If
	Set rsTmp = Conn.Execute("Select * From HR_Attach Where ClassID=" & tItemID & " And CourseID=" & tmpID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			m = 0
			picItem = "{""title"": """",""id"": " & tItemID & tmpID & ",""start"": 0,""data"": ["
			Do While Not rsTmp.EOF
				If m > 0 Then picItem = picItem & ","
				picItem = picItem & "{""alt"": """ & Trim(rsTmp("Title")) & """,""pid"": " & Trim(rsTmp("ID")) & ",""src"": """ & Trim(rsTmp("FilePath")) & """,""thumb"": """ & Trim(rsTmp("ThumbPic")) & """}"
				strPic = strPic & "<img layer-src=""" & rsTmp("FilePath") & """ alt=""" & rsTmp("Title") & """ src=""" & rsTmp("FilePath") & """ class=""layui-upload-img prevImg"">"
				inputPic = inputPic & "<div class=""layui-form-item""><label class=""layui-form-label"">图片" & rsTmp("ID") & ":</label><div class=""layui-input-block""><input type=""text"" name=""uploadPic"" value=""" & rsTmp("FilePath") & """ placeholder=""附件"" class=""layui-input""></div></div>"
				rsTmp.MoveNext
				m = m + 1
			Loop
			picItem = picItem & "]}"
		Else
		End If
	Set rsTmp = Nothing

	Set rsTmp = Conn.Execute("Select * From " & tSheetName & " Where ID=" & tmpID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			If HR_CBool(rsTmp("Passed")) Then isPassed = True
		End If
	Set rsTmp = Nothing

	tmpHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />" & vbCrlf
	tmpHtml = tmpHtml & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {padding: 10px;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding: 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-color-true {color:#080;} .hr-color-false {color:#F30;}" & vbCrlf
	tmpHtml = tmpHtml & "		#PreviewBox {display:flex;align-items:center;flex-wrap:wrap;}" & vbCrlf
	tmpHtml = tmpHtml & "		.prevImg {width: 92px; height: 92px; margin: 0 10px 10px 0;background-repeat:no-repeat;background-position: center;background-size:auto 92px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", "")
	strHeadHtml = Replace(strHeadHtml, "[@Head_style]", tmpHtml)

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		$(document).ready(function(){ });" & vbCrlf
	tmpHtml = tmpHtml & "		var picItem=eval(" & picItem & ");" & vbCrlf
	tmpHtml = tmpHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var layer = layui.layer, element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "			layer.config({skin:""layer-hr""});var loadInit = layer.load();" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Head_script]", tmpHtml)
	Response.Write strHeadHtml

	strHtml = "<fieldset class=""layui-elem-field layui-field-title""><legend>" & tItemName & " 课程附件1</legend>" & vbCrlf
	strHtml = strHtml & "</fieldset>" & vbCrlf
	strHtml = strHtml & "</fieldset>" & vbCrlf

	strHtml = strHtml & "<div class=""layer-hr-box"">" & vbCrlf
	strHtml = strHtml & "	<div class=""layui-upload"">" & vbCrlf
	If isPassed = False Then strHtml = strHtml & "		<button type=""button"" class=""layui-btn"" id=""UploadAttach"">图片附件上传</button>　注：单个文件不能超过2M，可多图片上传" & vbCrlf
	strHtml = strHtml & "		<blockquote class=""layui-elem-quote layui-quote-nm"" style=""margin-top: 10px;"">" & vbCrlf
	strHtml = strHtml & "		预览图：" & vbCrlf
	strHtml = strHtml & "		<div class=""layui-upload-list"" id=""PreviewBox"">" & strPic & "</div>" & vbCrlf
	strHtml = strHtml & "		</blockquote>" & vbCrlf
	strHtml = strHtml & "	</div>" & vbCrlf
	strHtml = strHtml & "</div>" & vbCrlf
	
	strHtml = strHtml & "<form class=""layui-form layui-form-pane"" id=""uploadForm"" name=""uploadForm"" action="""">" & vbCrlf
	strHtml = strHtml & "	<div class=""layer-hr-box"" id=""picBox"">" & inputPic & "</div>" & vbCrlf
	strHtml = strHtml & "	<div class=""layer-hr-box"">" & vbCrlf
	strHtml = strHtml & "		<input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""ItemID"" value=""" & tItemID & """><input type=""hidden"" name=""Modify"" value=""True"">" & vbCrlf
	If isPassed = False Then
		strHtml = strHtml & "		<div class=""layui-form-item"">" & vbCrlf
		strHtml = strHtml & "			<div class=""hr-btn-group""><button class=""layui-btn"" lay-submit lay-filter=""uploadPost"">保存</button><button type=""reset"" class=""layui-btn layui-btn-primary"">重置</button></div>" & vbCrlf
		strHtml = strHtml & "		</div>" & vbCrlf
	End If

	strHtml = strHtml & "	</div>" & vbCrlf
	strHtml = strHtml & "</form>" & vbCrlf

	strHtml = strHtml & "</fieldset>" & vbCrlf
	Response.Write strHtml

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""upload"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form, upload = layui.upload;" & vbCrlf
	strHtml = strHtml & "		upload.render({" & vbCrlf
	strHtml = strHtml & "			elem: '#UploadAttach',url: '/Manage/UploadFile.htm?UploadDir=Picture', accept:'file'" & vbCrlf
	strHtml = strHtml & "			,multiple: true,before: function(obj){" & vbCrlf		'//预读本地文件示例，不支持ie8
	strHtml = strHtml & "				obj.preview(function(index, file, result){" & vbCrlf
	strHtml = strHtml & "					$('#PreviewBox').append('<img layer-src=""'+ result + '"" alt=""'+ file.name +'"" src=""'+ result + '"" class=""layui-upload-img prevImg"">')" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "			,done: function(res, index){" & vbCrlf		'//上传完毕
	strHtml = strHtml & "				var extName = /\.[^\.]+$/.exec(res.data.src);console.log(extName);" & vbCrlf
	strHtml = strHtml & "				$(""#picBox"").append(""<div class=\""layui-form-item\""><label class=\""layui-form-label\"">图片"" + index + "":</label><div class=\""layui-input-block\""><input type=\""text\"" name=\""uploadPic\"" value=\"""" + res.data.src + ""\"" placeholder=\""附件\"" class=\""layui-input\"" /></div></div>"")" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "			,error: function (index, upload){console.log(index);}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		form.on(""submit(uploadPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Course/SaveAttach.html"", $(""#uploadForm"").serialize(), function(result){" & vbCrlf
	strHtml = strHtml & "				var reData = eval(""("" + result + "")"");var reMSG = reData.reMessge;" & vbCrlf
	strHtml = strHtml & "				layer.alert(reMSG, {icon:1,title: ""修改结果""},function(layero, index){parent.layer.closeAll();form.render();});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	'strHtml = strHtml & "		$(""#PreviewBox"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "			layer.photos({" & vbCrlf
	strHtml = strHtml & "				photos: '#PreviewBox',anim: 5" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	'strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@Foot_script]", strHtml)
	Response.Write strFootHtml
End Sub

Sub viewAttach()
	Dim getStr, tUrl : tUrl = Trim(ReplaceBadUrl(Request("url")))
	Dim tmpExtname, tmpHtml
	If HR_IsNull(tUrl) = False Then tmpExtname = Right(tUrl, Len(tUrl) - inStr(tUrl, "."))
	If LCase(tmpExtname) = "txt" Then getStr = ReadFromFile(tUrl, "GB2312", 0)

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body, .mediaPdf, #viewWord {width:100%;height:100%;} .mediaPdf {box-sizing: border-box;margin:0;padding:0;overflow: hidden}" & vbCrlf
	tmpHtml = tmpHtml & "		.dispText {margin:20px;padding:20px;box-sizing: border-box;border:1px solid #777;font-size:16px;color:#000;line-height:180%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.dispPic {margin:20px;padding:10px;box-sizing: border-box;background:#ddd;border:1px solid #777;}" & vbCrlf
	tmpHtml = tmpHtml & "		.dispPic img {border:1px solid #fff;width:100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.viewExcel th {min-width:100px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.media.js?v=0.99""></script>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Dim strXls, arrRow, strData, arrField, j
	Select Case LCase(tmpExtname)
		Case "txt" Response.Write "<div class=""dispText"">" & HR_HTMLEncode(getStr) & "</div>" & vbCrlf
		Case "jpg", "jpeg", "png", "bmp", "gif" Response.Write "<div class=""dispPic""><img src=""" & tUrl & """></div>" & vbCrlf
		Case "pdf" Response.Write "<a class=""mediaPdf"" href=""" & tUrl & """></a>" & vbCrlf
		'Case "doc", "docx" Response.Write "<frame name=""viewWord"" id=""viewWord"" title=""预览Word"" src=""http://www.xdocin.com/xdoc?_func=to&_format=html&_cache=1&_xdoc=" & apiHost & tUrl & """></frame>" & vbCrlf
		Case "xls", "xlsx"
			strXls = GetHttpPage(apiHost & "/API/ReadExcel.htm?type=2&xlsFile=" & tUrl, 1)
			If HR_IsNull(strXls) = False Then
				arrRow = Split(strXls, "@@")
				If Ubound(arrRow) > 0 Then
					strData = "<thead><tr>"
					arrField = Split(arrRow(0), "||")
					For i = 0 To Ubound(arrField)
						strData = strData & "<th>" & Trim(arrField(i)) & "</th>"
					Next
					strData = strData & "</tr></thead>"
					strData = strData & "<tbody>"
					For i = 1 To Ubound(arrRow)
						arrField = Split(arrRow(i), "||")
						strData = strData & "<tr>"
						For j = 0 To Ubound(arrField)
							strData = strData & "<td>" & Trim(arrField(j)) & "</td>"
						Next
						strData = strData & "</tr>" & vbCrlf
					Next
					strData = strData & "</tbody>" & vbCrlf
					strData = "<table class=""layui-table viewExcel"">" & strData & "</table>" & vbCrlf
				End If
			End If
			Response.Write "		<div class=""xlsData"" id=""xlsData"">" & strData & "</div>" & vbCrlf
	End Select

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	If tmpExtname = "doc" Or tmpExtname = "docx" Or tmpExtname = "rar" Or tmpExtname = "zip" Then strHtml = strHtml & "	window.open(""" & tUrl & """);" & vbCrlf
	strHtml = strHtml & "	layui.use([""layer"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		layer = layui.layer, element = layui.element;" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	If tmpExtname = "doc" Or tmpExtname = "docx" Or tmpExtname = "rar" Or tmpExtname = "zip" Then strHtml = strHtml & "		var index = parent.layer.getFrameIndex(window.name);parent.layer.close(index);" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	If tmpExtname = "pdf" Then strHtml = strHtml & "	$("".mediaPdf"").media({width:""100%"", height:""100%""});" & vbCrlf
	'strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub Delete()
	Dim tItemName, tTemplate, tUnit, tSheetName, delErrNum
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			tUnit = Trim(rsTmp("Unit"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = tItemID & "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing
	If ChkTable(tSheetName) = False Then
		ErrMsg = ErrMsg & tItemName & " 数据表未建立，请联系管理员！<br />"
	End If
	If ErrMsg<>"" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}" : Exit Sub
	End If

	Dim tmpJson, rsDel, sqlDel, strDel, arrDel, iDel, tmpErr, tYGDM, tStuType, tUpKPI, iDelNum
	strDel = Trim(ReplaceBadChar(Request("ID")))
	strDel = DelRightComma(strDel)
	arrDel = Split(strDel, ",")
	iDel = 0 : delErrNum = 0
	For iDelNum = 0 To Ubound(arrDel)
		sqlDel = "Select * From " & tSheetName & " Where ID=" & HR_Clng(arrDel(iDelNum))
		'If UserRank > 1 Then sqlDel = "Select * From " & tSheetName & " Where ID=" & HR_Clng(arrDel(iDelNum))				'超管可直接删除已经审核的记录
		Set rsDel = Server.CreateObject("ADODB.RecordSet")
			rsDel.Open(sqlDel), Conn, 1, 3
			If Not(rsDel.BOF And rsDel.EOF) Then
				If HR_CBool(rsDel("Passed")) Then	'不能删除已审记录
					delErrNum = delErrNum + 1
				Else
					tYGDM = rsDel("VA1"): tStuType = Trim(rsDel("StudentType"))
					rsDel.Delete
					iDel = iDel + 1
					tUpKPI = UpdateTeacherKPI(tItemID, tYGDM, "")	'更新本项目员工统计数据
					tUpKPI = UpdateTeacherTotalKPI(tYGDM)	'更新员工总计数据
				End If
				rsDel.Close
			End If
		Set rsDel = Nothing
	Next
	If delErrNum > 0 Then tmpErr = "<br><ul><li>其中 " & delErrNum & " 条记录已审核，无法删除！</li></ul>"
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & iDel & "/" & Ubound(arrDel) + 1 & " 条课程记录删除成功！" & tmpErr & """,""ReStr"":""操作成功！""}"
	Response.Write tmpJson
End Sub

Sub levelData()
	Dim tmpJson, rsGet, reData, tClassID, tClassName, j
	tClassID = HR_Clng(Request("item")) : tClassName = GetTypeName("HR_Class", "ClassName", "ClassID", tClassID)
	Set rsGet = Conn.Execute("Select * From HR_ItemModel Where ClassID=" & tClassID)
		If Not(rsGet.BOF And rsGet.EOF) Then
			reData = ""
			i = 0
			Do While Not rsGet.EOF
				If i > 0 Then reData = reData & ","
				reData = reData & "{""LevelID"":" & rsGet("ID") & ",""LevelName"":""" & rsGet("FieldName") & """,""Grade"":["
				Set rsTmp = Conn.Execute("Select * From HR_ItemGrade Where LevelID=" & rsGet("ID"))
					If Not(rsTmp.BOF And rsTmp.EOF) Then
						j = 0
						Do While Not rsTmp.EOF
							If j > 0 Then reData = reData & ","
							reData = reData & "{""GradeID"":" & rsTmp("ID") & ",""Grade"":""" & rsTmp("Grade") & """}"
							rsTmp.MoveNext
							j = j + 1
						Loop
					End If
				Set rsTmp = Nothing
				reData = reData & "]}"
				rsGet.MoveNext
				i = i + 1
			Loop
		End If
	Set rsGet = Nothing
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""级别数据获取成功！"",""ReStr"":""操作成功！"",""itemID"":" & tClassID & ",""itemName"":""" & tClassName & """,""data"":[" & reData & "]}"
	Response.Write tmpJson
End Sub

Sub CampusData()		'校(院)区JSON
	Dim tmpJson, rsGet, reData, j
	Dim arrCampus : arrCampus = Split(XmlText("Common", "Campus", ""), "|")
		reData = ""
		For i = 0 To Ubound(arrCampus)
			If i > 0 Then reData = reData & ","
			reData = reData & "{""CampusID"":" & i + 1 & ",""Campus"":""" & arrCampus(i) & """,""Items"":["
				Set rsTmp = Conn.Execute("Select * From HR_Period Where Campus='" & arrCampus(i) & "'")
					If Not(rsTmp.BOF And rsTmp.EOF) Then
						j = 0
						Do While Not rsTmp.EOF
							If j > 0 Then reData = reData & ","
							reData = reData & "{""PeriodID"":" & rsTmp("PeriodID") & ",""Period"":""" & rsTmp("Period") & """}"
							rsTmp.MoveNext
							j = j + 1
						Loop
					End If
				Set rsTmp = Nothing
			reData = reData & "]}"
		Next
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""校(院)区数据获取成功！"",""ReStr"":""操作成功！"",""data"":[" & reData & "]}"
	Response.Write tmpJson
End Sub

Sub ExcelTemp()
	Dim tClassID : tClassID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tClassName, tExcelFile, tFileName, tTemplate, strData, j
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tClassID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tClassName = rsTmp("ClassName")
			tTemplate = Trim(rsTmp("Template"))
			tFileName = Trim(rsTmp("ExcelFile"))
		End If
	Set rsTmp = Nothing

	Dim hrDate : hrDate = False		'判断VA4是否为日期
	If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Then hrDate = True

	Dim ModelFile : ModelFile = strGetTypeName("HR_DataModel", "ExcelFile", "ModelName", tTemplate)
	If HR_IsNull(tFileName) Then
		ErrMsg = "" & tClassName & " 项目的模板文件未设置！<br>"
		Response.Write GetErrBody(2) : Exit Sub
	Else
		tExcelFile = "/Upload/ExcelTemp/" & tFileName
	End If
	If fso.FileExists(Server.MapPath(tExcelFile)) = False Then
		ErrMsg = "没有找到 " & tClassName & " 项目的模板文件！<br>"
		Response.Write GetErrBody(1) : Exit Sub
	End If
	Dim arrRow, arrField, strTmp : strTmp = GetHttpPage(apiHost & "/API/ReadExcel.htm?type=1&xlsFile=" & tExcelFile, 1)
	If strTmp <> "" Then
		arrRow = Split(strTmp, "@@")
		If Ubound(arrRow) > 0 Then
			strData = "<thead><tr>"
			arrField = Split(arrRow(0), "||")
			For i = 0 To Ubound(arrField)
				strData = strData & "<th>" & Trim(arrField(i)) & "</th>"
			Next
			strData = strData & "</tr></thead>"
			strData = strData & "<tbody>"
			For i = 1 To Ubound(arrRow)
				arrField = Split(arrRow(i), "||")
				strData = strData & "<tr>"
				For j = 0 To Ubound(arrField)
					If j=4 And hrDate Then
						strData = strData & "<td>" & FormatDate(ConvertNumDate(arrField(j)), 2) & "</td>"
					Else
						strData = strData & "<td>" & Trim(arrField(j)) & "</td>"
					End If
				Next
				strData = strData & "</tr>" & vbCrlf
			Next
			strData = strData & "</tbody>" & vbCrlf
			strData = "<table class=""layui-table tplView"">" & strData & "</table>" & vbCrlf
		End If
	End If
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.downTips {color:#900;line-height:30px;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.Export {color:#f60;font-size:18px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.Export i {font-size: 18px!important;position: relative;top:2px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.tplView th, .tplView td {white-space:nowrap;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	Response.Write "	<legend>" & tClassName & " Excel模板管理</legend>" & vbCrlf
	If UserRank = 2 Then
		Response.Write "		<div class=""hr-shrink-x10"">" & vbCrlf
		Response.Write "			<div><button class=""layui-btn layui-btn-sm layui-btn-normal"" id=""AddBtn"" title=""修改Excel模板文件"" data-cid=""" & tClassID & """><i class=""layui-icon"">&#xe642;</i>设置</button></div>" & vbCrlf
		Response.Write "		</div>" & vbCrlf
	End If
	Response.Write "		<div class=""xlsData"" id=""xlsData"">" & strData & "</div>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf
	Response.Write "<div class=""layer-hr-box ExportBox"">" & vbCrlf
	Response.Write "	<div class=""downTips"">请点击鼠标右键后选择“另存为”</div>" & vbCrlf
	Response.Write "	<div class=""Export""><i class=""hr-icon"">&#xf019;</i><a href=""" & tExcelFile & """ id=""ExcelFile"">" & tExcelFile & "</a></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "	layui.use([""form"", ""table"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var form = layui.form, table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element; layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "		$(""#AddBtn"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			var cid = $(this).data(""cid"");" & vbCrlf
	strHtml = strHtml & "			layer.prompt({title:""修改Excel模板文件"", value:""" & tFileName & """}, function(value, index, elem){" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "KaoItem/SaveExcelName.html"",{ItemID:cid, FileName:value, Field:""AddNew""}, function(reData){" & vbCrlf
	strHtml = strHtml & "					if(reData.Return){window.location.reload();}else{" & vbCrlf
	strHtml = strHtml & "						layer.alert(reData.reMessge, function(index){layer.close(index); return false;});" & vbCrlf
	strHtml = strHtml & "					}" & vbCrlf
	strHtml = strHtml & "				}); return false;" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub
Sub SaveExcelName()
	Dim tClassID : tClassID = HR_Clng(Request("ItemID"))
	Dim tmpJson, tExcelFile, tFileName : tFileName = Trim(ReplaceBadUrl(Request("FileName")))
	If tFileName <> "" Then tExcelFile = "/Upload/ExcelTemp/" & tFileName
	If fso.FileExists(Server.MapPath(tExcelFile)) Then
		Conn.Execute("Update HR_Class Set ExcelFile='" & tFileName & "' Where ClassID=" & tClassID)
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""Excel模板文件修改成功！"",""ReStr"":""修改成功！""}"
	Else
		tmpJson = "{""Return"":false,""Err"":400,""reMessge"":""Excel模板文件不存在！"",""ReStr"":""修改失败！""}"
	End If
	Response.Write tmpJson
End Sub

Sub ViewItem()
	Dim tmpID : tmpID = HR_Clng(Request("ItemID"))
	Dim arrItemType : arrItemType = Split(XmlText("Common", "ItemType", ""), "|")

	Dim rsShow, tItemName, tTypeID, tType, tParent, tStudentType, tRatio, tHtml
	Set rsShow = Conn.Execute("Select * From HR_Class Where ClassID=" & tmpID )
		If rsShow.BOF And rsShow.EOF Then
			tHtml = "<div class=""layui-row"" style=""text-align:center;padding:20px 0"">"
			tHtml = tHtml & "<a class=""layui-btn layui-btn-lg layui-btn-danger""><i class=""layui-icon"">&#xe69c;</i> 您要查看的考核项目信息【ID：" & tmpID & "】不存在！</a></div>"
			Response.Write tHtml
			Exit Sub
		Else
			tItemName = Trim(rsShow("ClassName"))
			tTypeID = HR_Clng(rsShow("ClassType"))
			tParent = GetTypeName("HR_Class", "ClassName", "ClassID", HR_Clng(rsShow("ParentID")))
			If tTypeID > 0 Then tType = arrItemType(tTypeID - 1)
			tStudentType = Trim(rsShow("StudentType"))
			tStudentType = Replace(tStudentType, ",", "　")
			tRatio = Trim(rsShow("Ratio"))
			tRatio = Replace(tRatio, ",", "　")

			tHtml = "<fieldset class=""layui-elem-field layui-field-title""><legend>考核项目 " & tItemName & " 预览</legend>"
			tHtml = tHtml & "<div class=""layui-form layer-hr-box""><table class=""layui-table"">"
			tHtml = tHtml & "<colgroup><col width=""120""><col><col width=""120""><col></colgroup>"
			tHtml = tHtml & "<tbody>"

			tHtml = tHtml & "<tr><td style=""text-align:right;"">考核项目：</td><td colspan=""3"">" & tItemName & "</td></tr>" & vbCrlf
			tHtml = tHtml & "<tr><td style=""text-align:right;"">序　　号：</td><td>" & HR_Clng(rsShow("ClassID")) & "</td>"
			tHtml = tHtml & "<td style=""text-align:right;"">类　　别：</td><td>" & Trim(tType) & "</td></tr>" & vbCrlf
			tHtml = tHtml & "<tr><td style=""text-align:right;"">上级项目：</td><td>" & tParent & "</td>"
			tHtml = tHtml & "<td style=""text-align:right;"">计量单位：</td><td>" & Trim(rsShow("Unit")) & "</td></tr>"
			tHtml = tHtml & "<tr><td style=""text-align:right;"">数据模板：</td><td>" & Trim(rsShow("Template")) & "</td>"
			tHtml = tHtml & "<td style=""text-align:right;"">数据表名：</td><td>" & Trim(rsShow("SheetName")) & "</td></tr>"
			If tStudentType <> "" Then tHtml = tHtml & "<tr><td style=""text-align:right;"">学生类别：</td><td colspan=""3"">" & tStudentType & "</td></tr>"
			tHtml = tHtml & "<tr><td style=""text-align:right;"">考核系数：</td><td colspan=""3"">" & tRatio & "</td></tr>"
			tHtml = tHtml & "<tr><td style=""text-align:right;"">项目说明：</td><td colspan=""3"">" & Trim(rsShow("Tips")) & "</td></tr>"
			tHtml = tHtml & "</tbody>"
			tHtml = tHtml & "</table></div>" & vbCrlf
			tHtml = tHtml & "</fieldset>" & vbCrlf
		End If
	Set rsShow = Nothing

	Response.Write tHtml
End Sub

Sub SwitchLock()	'锁定开关
	Server.ScriptTimeout=1800		'30分钟
	Dim tmpJson, rsUpdate, tClassID : tClassID = HR_Clng(Request("ItemID"))
	Dim tmpLocked : tmpLocked = Trim(Request("Locked"))

	Dim tClassName, tTemplate, tSheetName, tUpKPI
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tClassID)
		If Not(rsTmp.BOF And rsTmp.EOF) And tClassID > 0 Then
			tClassName = rsTmp("ClassName")
			tTemplate = Trim(rsTmp("Template"))
			tSheetName = Trim(rsTmp("SheetName"))
			If ChkTable(tSheetName) Then
				If tmpLocked Then
					ErrMsg = tClassName & " 中的所有未审的课程业绩已审核成功！"
					Conn.Execute("Update " & tSheetName & " Set Passed=" & HR_True & " Where Passed=" & HR_False )
					Call RecordFrontLog(tClassID, "一键审核", UserName & " 进行了" & tClassName & "[ID：" & UserID & "]一键审核操作", True, "Setup")
				Else
					ErrMsg = tClassName & " 中的所有课程业绩已解除锁定，允许修改！"
					Conn.Execute("Update " & tSheetName & " Set Passed=" & HR_False & " Where Passed=" & HR_True )
					Call RecordFrontLog(tClassID, "一键反审核", UserName & " 进行了" & tClassName & "[ID：" & UserID & "]一键反审核操作", True, "Setup")
				End If
				Set rsUpdate = Conn.Execute("Select VA1 From " & tSheetName & " Where VA1>0 Group By VA1")
					If Not(rsUpdate.BOF And rsUpdate.EOF) Then
						Do While Not rsUpdate.EOF
							Call ChkTeacherKPI(rsUpdate("VA1"))
							tUpKPI = UpdateTeacherKPI(tClassID, rsUpdate("VA1"), "")	'更新本项目员工统计数据
							tUpKPI = UpdateTeacherTotalKPI(rsUpdate("VA1"))	'更新员工总计数据
							rsUpdate.MoveNext
						Loop
					End If
				Set rsUpdate = Nothing
			End If
			tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """,""ReStr"":""设置成功！""}"
		Else
			tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""考核项目不存在或已删除！"",""ReStr"":""操作失败！""}"
		End If
	Set rsTmp = Nothing
	
	Response.Write tmpJson
End Sub


%>