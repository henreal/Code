<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./incKPI.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim arrStudentType : arrStudentType = Split(XmlText("Common", "StudentType", ""), "|")

Dim Page_Title : Page_Title = "课程业绩提交"
Dim SubButTxt : SubButTxt = "所有课程"

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))				'Get head template code.
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))				'Get foot template code.
Dim strNavPath : strNavPath = ReplaceCommonLabel(getFrameNav(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "AllData" Call GetDataList()
	Case "BatchPass" Call BatchPass()
	Case "BatchDel" Call BatchDel()
	Case "ItemUp" Call ItemUp()
	Case "ItemUpData" Call ItemUpData()
	Case "applyModify" Call applyModify()
	Case "SaveApply" Call SaveApply()

	Case "SwitchLock" Call SwitchLock()		'解锁
	Case "Affirm" Call Affirm()				'员工确认课程提交
	Case "oneAffirm" Call oneAffirm()
	Case "backModify" Call backModify()		'退回修改
	Case "backSave" Call backSave()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tClassID : tClassID = HR_Clng(Request("ItemID"))
	Dim tClassName, tSheetName, tStuType, tTemplate, tFieldLen, tFieldHead, arrHead
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tClassID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tClassName = Trim(rsTmp("ClassName"))
			tSheetName = Trim(rsTmp("SheetName"))
			tStuType = Trim(rsTmp("StudentType"))
			tTemplate = Trim(rsTmp("Template"))
			tFieldLen = HR_Clng(rsTmp("FieldLen"))
			tFieldHead = Trim(rsTmp("FieldHead"))
		End If
	Set rsTmp = Nothing

	If tFieldHead <> "" Then
		arrHead = Split(tFieldHead, ",")
		If Ubound(arrHead) <> tFieldLen-1 Then Redim Preserve arrHead(tFieldLen-1)
	Else
		Redim arrHead(tFieldLen)
	End If

	strHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />" & vbCrlf
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.iframe-nav .navBtn .navLayer {font-size: 16px;}" & vbCrlf

	strHtml = strHtml & "		.tplBtn .layui-btn-sm i {font-size: 14px!important;}" & vbCrlf
	strHtml = strHtml & "		.tplBtn ..layui-btn-sm {height: 25px; line-height: 25px; padding: 0 8px;}" & vbCrlf
	strHtml = strHtml & "		.soBox .searchBtn {vertical-align:top} .soBox .layui-inline {margin-bottom:1px;} .soBox .layui-form-select dl {top: 31px;}" & vbCrlf
	strHtml = strHtml & "		.soBox .layui-input {height: 30px;}" & vbCrlf
	strHtml = strHtml & "		.soBox .layui-btn {height: 30px;line-height: 30px;padding: 0 12px;}" & vbCrlf
	strHtml = strHtml & "		.soBox .layui-form-select dl dd {padding: 0 5px;line-height: 30px;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", Page_Title)
	strHeadHtml = Replace(strHeadHtml, "[@Head_style]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "			var layer = layui.layer, element = layui.element;" & vbCrlf
	strHtml = strHtml & "			layer.config({skin:""layer-hr""});var loadInit = layer.load();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf

	strHeadHtml = Replace(strHeadHtml, "[@Head_script]", strHtml)
	Response.Write strHeadHtml
	'strHtml = "<a href=""" & ParmPath & "Course.html?ItemID=" & tClassID & """>" & tClassName & "</a><a><cite>数据核验</cite></a>"
	'strNavPath = Replace(strNavPath, "[@Module_Path]", strHtml)
	'Response.Write strNavPath

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title"" style=""margin-top:1px;""><legend>" & tClassName & "　数据核验</legend></fieldset>" & vbCrlf
	Response.Write "	<div class=""layui-form soBox""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input class=""layui-input"" name=""soTeacher"" value="""" id=""soTeacher"" placeholder=""员工姓名/工号"" autocomplete=""off"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><select name=""sortType"" id=""sortType""><option value="""">选择排序方式</option><option value=""importTimeUP"">上传时间正序↑</option><option value=""importTimeDown"">上传时间倒序↓</option>"
	Response.Write "<option value=""xhUP"">序号正序↑</option><option value=""xhDown"">序号倒序↓</option></select></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn"">" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-normal"" data-type=""reload"" id=""SearchBtn""><i class=""layui-icon"">&#xe615;</i>搜索</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-normal"" data-type=""refresh"" id=""refresh"" title=""刷新""><i class=""hr-icon"">&#xebbb;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn"" data-type=""pass"" id=""BatchPass"" title=""一键提交""><i class=""hr-icon"">&#xebc5;</i>提交</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-normal"" data-type=""delete"" id=""BatchDel"" title=""批量删除""><i class=""hr-icon"">&#xea64;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "CourseProof/AllData.html?ItemID=" & tClassID & "',height:'full-120',page:true,limit:20,limits:[10,15,20,30,50,100],text:{none:'您暂时还没有需要核验的课程业绩！'},id:'TableList'}"" lay-filter=""TableList"">"
	Response.Write "		<thead><tr>" & vbCrlf
	Response.Write "			<th lay-data=""{fixed:'left',type:'checkbox',unresize:true,align:'center',width:60}""></th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'VA0',align:'center', width:70,sort:true}"">序号</th>" & vbCrlf
	If tStuType <> "" Then Response.Write "			<th lay-data=""{field:'StudentType', width:80}"">类别</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'VA1',unresize:true, width:80}"">工号</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'VA2',width:100}"">" & arrHead(2) & "</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'KSMC',width:100}"">科室</th>" & vbCrlf
	For i = 4 To Ubound(arrHead)
		Response.Write "			<th lay-data=""{field:'VA" & i & "',minWidth:100}"">" & arrHead(i) & "</th>" & vbCrlf
	Next
	Response.Write "			<th lay-data=""{field:'AppendTime',align:'center',width:160}"">上传时间</th>" & vbCrlf
	Response.Write "			<th lay-data=""{fixed:'right',align:'center',unresize:true,width:100, toolbar: '#barTable'}"">操作</th>" & vbCrlf
	Response.Write "		</tr></thead>" & vbCrlf
	Response.Write "	</table>" & vbCrlf

	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group tplBtn"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-danger"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm"" lay-event=""attach"" title=""提交""><i class=""hr-icon"">&#xebc5;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table, form = layui.form;" & vbCrlf
	strHtml = strHtml & "		element = layui.element;" & vbCrlf
	strHtml = strHtml & "		$("".searchBtn button"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var btnEvent = $(this).data(""type"");" & vbCrlf
	strHtml = strHtml & "			var checkStatus = table.checkStatus(""TableList""), arrID=[];" & vbCrlf
	strHtml = strHtml & "			for(var i=0;i<checkStatus.data.length;i++){" & vbCrlf
	strHtml = strHtml & "				arrID.push(checkStatus.data[i].ID);" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf

	strHtml = strHtml & "			if(btnEvent==""reload""){" & vbCrlf
	strHtml = strHtml & "				var soTeacher = $(""#soTeacher"").val(), soType = $(""#sortType"").val();" & vbCrlf		'员工搜索、排序
	strHtml = strHtml & "				table.reload(""TableList"",{" & vbCrlf
	strHtml = strHtml & "					url:'" & ParmPath & "CourseProof/AllData.html', where:{ItemID:" & tClassID & ",soWord:soTeacher, soType:soType}" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}else if(btnEvent==""refresh""){" & vbCrlf
	strHtml = strHtml & "				table.reload(""TableList"",{url:'" & ParmPath & "CourseProof/AllData.html',where:{ItemID:" & tClassID & ",soWord:"""", soType:""""}});" & vbCrlf
	strHtml = strHtml & "			}else if(btnEvent==""pass""){" & vbCrlf
	strHtml = strHtml & "				layer.confirm(""您将提交所有的课程记录，是否继续？"",{icon:0, title:""重要提示""},function(index){" & vbCrlf
	'strHtml = strHtml & "				if(checkStatus.data.length==0){layer.tips(""请选择您要提交的课程！"",""#BatchPass"",{tips: [3, ""#4A5""]});return false;}" & vbCrlf
	strHtml = strHtml & "					$.getJSON(""" & ParmPath & "CourseProof/BatchPass.html"",{ItemID:" & tClassID & ",ID:arrID.join()}, function(strForm){" & vbCrlf
	strHtml = strHtml & "						layer.msg(strForm.reMessge,{icon:1,time:0,btn:""关闭""},function(){table.reload(""TableList"");});" & vbCrlf
	strHtml = strHtml & "					});return false;" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}else if(btnEvent==""delete""){" & vbCrlf
	strHtml = strHtml & "				if(checkStatus.data.length==0){layer.tips(""请选择要删除的课程记录！"",""#BatchDel"",{tips: [3, ""#F30""]});return false;}" & vbCrlf
	strHtml = strHtml & "				layer.confirm(""确认要删除选中的“"" + checkStatus.data.length + ""”条课程记录？<br />删除后无法恢复"",{icon:0, title:""重要提示""},function(index){" & vbCrlf
	strHtml = strHtml & "					$.getJSON(""" & ParmPath & "CourseProof/BatchDel.html"",{ItemID:" & tClassID & ",ID:arrID.join()}, function(strForm){" & vbCrlf
	strHtml = strHtml & "						layer.msg(strForm.reMessge,{icon:1,time:0,btn:""关闭""},function(){table.reload(""TableList"");});" & vbCrlf
	strHtml = strHtml & "					});return false;" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf

	strHtml = strHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	strHtml = strHtml & "			var data = obj.data;" & vbCrlf
	strHtml = strHtml & "			if(obj.event === ""del""){" & vbCrlf
	strHtml = strHtml & "				layer.confirm(""确认要删除选中的课程记录？<br />删除后无法恢复"",{icon:0, title:""重要提示""},function(index){" & vbCrlf
	strHtml = strHtml & "					$.getJSON(""" & ParmPath & "CourseProof/BatchDel.html"",{ItemID:" & tClassID & ",ID:data.ID}, function(strForm){" & vbCrlf
	strHtml = strHtml & "						layer.msg(strForm.reMessge,{icon:1,time:0,btn:""关闭""},function(){table.reload(""TableList"");});" & vbCrlf
	strHtml = strHtml & "					});return false;" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf

	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@Foot_script]", strHtml)
	Response.Write strFootHtml
End Sub

Sub GetDataList()
	Dim tClassID : tClassID = HR_Clng(Request("ItemID"))
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim soType : soType = Trim(ReplaceBadChar(Request("soType")))

	Dim tClassName, tSheetName, tStuType, tTemplate, tUnit
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tClassID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tClassName = Trim(rsTmp("ClassName"))
			tSheetName = Trim(rsTmp("SheetName"))
			tStuType = Trim(rsTmp("StudentType"))
			tTemplate = Trim(rsTmp("Template"))
			tUnit = rsTmp("Unit")
		End If
	Set rsTmp = Nothing

	If Not(ChkTable(tSheetName)) Then		'检查表是否存在
		tmpJson = "{""code"":500, ""msg"":""数据表 " & tSheetName & " 不存在！"",""count"":0,""data"":[]}"
		Response.Write tmpJson : Exit Sub
	End If
	Dim vCount, vMSG, tmpJson, tmpData, rsGet, sqlGet, isErr : isErr = False
	sqlGet = "Select * From " & tSheetName & " Where ItemID=" & tClassID & " And State=0"
	sqlGet = sqlGet & " And UserID=" & UserID		'仅可查看自己上传的

	If soWord <> "" Then
		If HR_Clng(soWord) > 0 Then
			sqlGet = sqlGet & " And VA1='" & soWord & "'"		'搜索员工工号
		Else
			sqlGet = sqlGet & " And VA2 like '%" & soWord & "%'"		'搜索员工姓名
		End If
	End If

	If soType = "xhUP" Then
		sqlGet = sqlGet & " Order By VA0 ASC"
	ElseIf soType = "xhDown" Then
		sqlGet = sqlGet & " Order By VA0 DESC"
	ElseIf soType = "importTimeUP" Then
		sqlGet = sqlGet & " Order By AppendTime ASC"
	Else
		sqlGet = sqlGet & " Order By AppendTime DESC"
	End If
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			i = 0 :CurrentPage = 1
			If HR_Clng(Trim(Request("page"))) > 0 Then CurrentPage = HR_Clng(Trim(Request("page")))
			MaxPerPage = HR_Clng(Trim(Request("limit")))
			If MaxPerPage <= 0 Then MaxPerPage = 20
			TotalPut = rsGet.Recordcount
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
						rsGet.Move (CurrentPage - 1) * MaxPerPage
					Else
						CurrentPage = 1
					End If
				End If
			End If

			Dim tVA7, tmpTime
			Do While Not rsGet.EOF
				If i > 0 Then tmpData = tmpData & ","
				If tTemplate = "TempTableA" Then
					tVA7 = Trim(rsGet("VA7"))
					tmpTime = GetPeriodTime(Trim(rsGet("VA11")), tVA7, 0)		'计算节次时间
				End If

				tmpData = tmpData & "{""ID"":" & rsGet("ID") & ",""CourseID"":""" & HR_Clng(rsGet("ItemID")) & """,""Course"":""" & Trim(tClassName) & """,""StudentType"":""" & Trim(rsGet("StudentType")) & """,""KSMC"":""" & Trim(rsGet("KSMC")) & """"
				tmpData = tmpData & ",""Time"":""" & Trim(tmpTime) & """"
				For m = 3 To rsGet.Fields.Count - 2
					If m = 7 And (tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE") Then
						tmpData = tmpData & ",""" & rsGet.Fields(m).name & """:""" & FormatDate(ConvertNumDate(rsGet("VA" & m-3 & "")), 2) & """"
					Else
						tmpData = tmpData & ",""" & rsGet.Fields(m).name & """:""" & HR_HTMLEncode(rsGet.Fields(m).value) & """"
					End If
				Next
				tmpData = tmpData & ",""Unit"":""" & tUnit & """,""AppendTime"":""" & FormatDate(rsGet("AppendTime"), 1) & """}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
		vCount = rsGet.Recordcount
	Set rsGet = Nothing
	tmpJson = "{""code"":0, ""msg"":""课程查询成功！ " & sqlGet & """,""count"":" & vCount & ",""data"":[" & tmpData
	tmpJson = tmpJson & "]}"
	Response.Write tmpJson
End Sub

Sub ItemUp()
	Server.ScriptTimeout = 900
	Dim timeStart : timeStart = Timer
	Dim rsGet, sqlGet, tStrItem, tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim chk1, tSheetName : tSheetName = "HR_Sheet_" & tItemID
	If ChkTable(tSheetName) Then
		sqlGet = "Select VA1 As YGDM From " & tSheetName & " Where ItemID=" & tItemID & " Group By VA1"
		Set rsGet = Server.CreateObject("ADODB.RecordSet")
			rsGet.Open sqlGet, Conn, 1, 1
			If Not(rsGet.BOF And rsGet.EOF) Then
				Do While Not rsGet.EOF
					chk1 = ChkTeacherKPI(rsGet("YGDM"))
					rsGet.MoveNext
				Loop
				tStrItem = "更新员工：" & rsGet.Recordcount & "条"
			End If
		Set rsGet = Nothing
	Else
		tStrItem = tSheetName & "表不存在"
	End If
	Response.Write tStrItem & "，时间：" & Timer - timeStart & " 秒"
End Sub

Sub ItemUpData()
	Server.ScriptTimeout = 900
	Dim timeStart : timeStart = Timer
	Dim rsGet, sqlGet, tStrItem, tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim chk1, tSheetName : tSheetName = "HR_Sheet_" & tItemID
	Dim iCount, tPage : tPage = HR_Clng(Request("Page"))

	If ChkTable(tSheetName) Then
		sqlGet = "Select VA1,StudentType From " & tSheetName & ""
		Set rsGet = Server.CreateObject("ADODB.RecordSet")
			rsGet.Open sqlGet, Conn, 1, 1
			If Not(rsGet.BOF And rsGet.EOF) Then
				iCount = 0 : CurrentPage = 1
				If tPage > 0 Then CurrentPage = tPage
				MaxPerPage = 500 : TotalPut = rsGet.Recordcount		'每100条更新一次
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
							rsGet.Move (CurrentPage - 1) * MaxPerPage
						Else
							CurrentPage = 1
						End If
					End If
				End If
				Do While Not rsGet.EOF
					If HR_Clng(rsGet("VA1")) > 0 Then
						chk1 = ChkTeacherKPI(rsGet("VA1"))
						chk1 = UpdateTeacherKPI(tItemID, rsGet("VA1"), Trim(rsGet("StudentType")))
						chk1 = UpdateTeacherTotalKPI(rsGet("VA1"))
					End If
					rsGet.MoveNext
					iCount = iCount + 1
					If iCount >= MaxPerPage Then Exit Do
				Loop
				tStrItem = "更新员工：" & iCount & "条"
			End If
		Set rsGet = Nothing
	Else
		tStrItem = tSheetName & "表不存在"
	End If
	ErrMsg = "{""Return"":true,""Err"":0,""reMessge"":""" & tStrItem & "！耗时：" & Timer - timeStart & " 秒"",""ReStr"":""操作成功！"",""Total"":" & TotalPut & ",""Page"":" & CurrentPage & "}"
	Response.Write ErrMsg
End Sub

Sub BatchPass()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tItemName, tTemplate, tSheetName
	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = tItemID & "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing
	If ChkTable(tSheetName) = False Then ErrMsg = ErrMsg & tItemName & " 数据表未建立，请联系管理员！<br />"

	If ErrMsg<>"" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}" : Exit Sub
	End If

	Dim PassNum
	Set rsTmp = Conn.Execute("Select Count(ID) From " & tSheetName & " Where State=1 And UserID=" & UserID)
		PassNum = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing
	Conn.Execute("Update " & tSheetName & " Set State=3,Passed=1 Where State=1 And UserID=" & UserID)
	ErrMsg = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & PassNum & " 条记录提交并审核成功！"",""ReStr"":""操作成功！""}"
	Response.Write ErrMsg
End Sub

Sub BatchDel()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim iDel, iArrNum, rsDel, sqlDel, tArrID, strArrID : strArrID = Trim(ReplaceBadChar(Request("ID")))
	strArrID = DelRightComma(strArrID)

	Dim tItemName, tTemplate, tSheetName, tUpKPI, tYGDM, tStuType
	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = tItemID & "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing
	If ChkTable(tSheetName) = False Then ErrMsg = ErrMsg & tItemName & " 数据表未建立，请联系管理员！<br />"
	If HR_IsNull(strArrID) Then ErrMsg = ErrMsg & " 您没有选择核验记录！<br />"

	If ErrMsg<>"" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}" : Exit Sub
	End If

	If strArrID <> "" Then
		tArrID = Split(strArrID, ",") : iDel = 0
		For iArrNum = 0 To Ubound(tArrID)
			sqlDel = "Select VA1,StudentType From " & tSheetName & " Where ID=" & HR_Clng(tArrID(iArrNum))
			If UserRank < 2 Then sqlDel = sqlDel & " And (Passed=0 Or ISNULL(Passed,0)=0) And UserID=" & UserID	'仅删除本人上传且未审
			Set rsDel = Server.CreateObject("ADODB.RecordSet")
				rsDel.Open(sqlDel), Conn, 1, 3
				If Not(rsDel.BOF And rsDel.EOF) Then
					tYGDM = rsDel(0) : tStuType = Trim(rsDel(1))
					rsDel.Delete
					iDel = iDel + 1
					tUpKPI = UpdateTeacherKPI(tItemID, tYGDM, tStuType)	'更新本项目员工统计数据
					tUpKPI = UpdateTeacherTotalKPI(tYGDM)	'更新员工总计数据
				End If
			Set rsDel = Nothing
		Next
		ErrMsg = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & iDel & "/" & Ubound(tArrID) + 1 & " 条课程记录删除成功！"",""ReStr"":""操作成功！""}"
		Response.Write ErrMsg
	End If
End Sub

Function GetPeriodTime(fCampus, fPeriod, fType)
	Dim strFun, rsFun, fArr, strArr, stTime, enTime
	fCampus = Trim(fCampus) : fPeriod = Trim(fPeriod)
	If fCampus <> "" And fPeriod <> "" Then
		If HR_Clng(fPeriod) > 0 Then
			Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fPeriod))
				If Not(rsFun.BOF And rsFun.EOF) Then
					strFun = Trim(rsFun("StartTime")) & " - " & Trim(rsFun("EndTime"))
				End If
			Set rsFun = Nothing
		ElseIf Instr(fPeriod, "-") Then
			fArr = Split(fPeriod, "-")
			If Ubound(fArr) = 1 Then
				Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fArr(0)))
					If Not(rsFun.BOF And rsFun.EOF) Then
						stTime = Trim(rsFun("StartTime"))
					End If
				Set rsFun = Nothing
				Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fArr(1)))
					If Not(rsFun.BOF And rsFun.EOF) Then
						enTime = Trim(rsFun("EndTime"))
					End If
				Set rsFun = Nothing
				strFun = Trim(stTime) & " - " & Trim(enTime)
			End If
		End If
	End If
	GetPeriodTime = strFun
End Function


Sub applyModify()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))

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

	strHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />" & vbCrlf
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.iframe-nav .navBtn .navLayer {font-size: 16px;}" & vbCrlf

	strHtml = strHtml & "		.tplBtn .layui-btn-sm i {font-size: 14px!important;}" & vbCrlf
	strHtml = strHtml & "		.tplBtn ..layui-btn-sm {height: 25px; line-height: 25px; padding: 0 8px;}" & vbCrlf
	strHtml = strHtml & "		.soBox .searchBtn {vertical-align:top} .soBox .layui-inline {margin-bottom:1px;} .soBox .layui-form-select dl {top: 31px;}" & vbCrlf
	strHtml = strHtml & "		.soBox .layui-input {height: 30px;}" & vbCrlf
	strHtml = strHtml & "		.soBox .layui-btn {height: 30px;line-height: 30px;padding: 0 12px;}" & vbCrlf
	strHtml = strHtml & "		.soBox .layui-form-select dl dd {padding: 0 5px;line-height: 30px;}" & vbCrlf

	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", Page_Title)
	strHeadHtml = Replace(strHeadHtml, "[@Head_style]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "			var layer = layui.layer, element = layui.element;" & vbCrlf
	strHtml = strHtml & "			layer.config({skin:""layer-hr""});var loadInit = layer.load();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf

	strHeadHtml = Replace(strHeadHtml, "[@Head_script]", strHtml)
	Response.Write strHeadHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title"" style=""margin-top:1px;""><legend>申请修改 " & tItemName & " 业绩记录</legend></fieldset>" & vbCrlf
	Response.Write "	<form class=""layui-form layui-form-pane"" id=""ApplyForm"" name=""ApplyForm"" lay-filter=""ApplyForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layui-form-item layui-form-text"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">申请理由：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><textarea name=""Explain"" id=""Explain"" placeholder=""备注"" class=""layui-textarea""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<input name=""ItemID"" type=""hidden"" value=""" & tItemID & """><input name=""ID"" type=""hidden"" value=""" & tmpID & """>" & vbCrlf
	Response.Write "	<input name=""ygdm"" type=""hidden"" value=""" & UserYGDM & """><input name=""userid"" type=""hidden"" value=""" & UserID & """>" & vbCrlf
	Response.Write "	<div class=""searchBtn"">" & vbCrlf
	Response.Write "		<button class=""layui-btn"" type=""button"" id=""ApplyPost"" title=""提交申请""><i class=""hr-icon"">&#xebc5;</i>提交申请</button>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table, form = layui.form;" & vbCrlf
	strHtml = strHtml & "		element = layui.element;" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "		$(""#ApplyPost"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var strExplain = $(""#Explain"").val();" & vbCrlf
	strHtml = strHtml & "			if(strExplain ==""""){layer.msg(""您没有填写申请的理由！"");return false;}" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "CourseProof/SaveApply.html"",$(""#ApplyForm"").serialize(), function(strForm){" & vbCrlf
	strHtml = strHtml & "				layer.msg(strForm.reMessge,{icon:6,time:0,btn:""关闭""},function(){" & vbCrlf
	'strHtml = strHtml & "				var index1 = parent.layer.getFrameIndex(window.name);" & vbCrlf
	'strHtml = strHtml & "				parent.layer.close(index1);" & vbCrlf		'关闭自身，在iframe页面
	strHtml = strHtml & "				parent.layer.closeAll();" & vbCrlf
	strHtml = strHtml & "				return false;" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@Foot_script]", strHtml)
	Response.Write strFootHtml
End Sub
Sub SaveApply()
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
	If UserRank > 1 Then
		ErrMsg = ErrMsg & "您有这么高的管理级别，不用申请了，想干嘛都行！<br>"
	End If
	Dim tmpStuType, tYGXM, tYGDM, tPXXH, sendUserID
	Set rsTmp = Conn.Execute("Select * From " & tSheetName & " Where ID=" & tID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tmpStuType = Trim(rsTmp("StudentType"))
			tYGXM = Trim(rsTmp("VA2"))
			tYGDM = HR_Clng(rsTmp("VA1"))
			tPXXH = HR_Clng(rsTmp("VA0"))
			sendUserID = HR_Clng(rsTmp("UserID"))
		Else
			ErrMsg = ErrMsg & tItemName & "课程业绩不存在！【ID：" & tID & "】<br>"
		End If
	Set rsTmp = Nothing
	If ErrMsg <> "" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}"
		Exit Sub
	End If

	Dim SendMsgManager, tArrSender, SentMsg
	If UserYGDM <> "" And UserID=0 Then
		ErrMsg = UserYGXM & "申请修改课程业绩，考核项目：" & tItemName & "，序号：" & tPXXH & "，教师" & UserYGDM & "[工号 " & UserYGXM & "]，时间：" & FormatDate(Now(), 1)
		ErrMsg = ErrMsg & " <span class=""ShowCourse"" data-ItemID=""" & tItemID & """ data-id=""" & tID & """ data-sender=""" & UserYGDM & """>【查看】</span>"
		If HR_IsNull(tExplain) = False Then
			ErrMsg = ErrMsg & "<br>申请理由：" & tExplain
			ErrMsg = ErrMsg & "<br><span class=""BackApply"" data-ItemID=""" & tItemID & """ data-id=""" & tID & """ data-sender=""" & UserYGDM & """>【退回】</span>"
			ErrMsg = ErrMsg & "<span class=""Transfer"" data-ItemID=""" & tItemID & """ data-id=""" & tID & """ data-sender=""" & UserYGDM & """>【转交超管】</span>"
		End If
		If sendUserID > 0 Then
			SentMsg = SendMessage(1, sendUserID, UserYGXM & "申请修改课程业绩", ErrMsg, 0)
		Else
			If HR_IsNull(tmpStuType) = False Then
				SendMsgManager = GetManagerID(tmpStuType, 0)
				tArrSender = Split(SendMsgManager, ",")
			Else
				tArrSender = arrManager
			End If
			For i = 0 To Ubound(tArrSender)
				SentMsg = SendMessage(1, tArrSender(i), UserYGXM & "申请修改课程业绩", ErrMsg, 0)
			Next
		End If
	ElseIf UserID > 0 Then		'向超管发送申请
		Set rsTmp = Conn.Execute("Select * From HR_User Where ManageRank>1")
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				Do While Not rsTmp.EOF
					ErrMsg = "管理员" & UserName & "申请修改课程业绩，考核项目：" & tItemName & "，序号：" & tPXXH & "，教师" & tYGXM & "[工号 " & tYGDM & "]，时间：" & FormatDate(Now(), 1)
					ErrMsg = ErrMsg & " <a href=""" & ParmPath & "Course.html?ItemID=" & tItemID & "&SearchWord=" & tYGDM & """>【查看】</a><br>申请理由：" & tExplain
					SentMsg = SendMessage(1, rsTmp("UserID"), "管理员" & UserName & "申请修改课程业绩", ErrMsg, 0)
					rsTmp.MoveNext
				Loop
			End If
		Set rsTmp = Nothing
	End If
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""修改课程业绩申请提交成功！<br />"",""ReStr"":""操作成功！""}"
End Sub

Sub SwitchLock()			'锁定或解锁
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim Locked : Locked = HR_CBool(Request("Locked"))
	SubButTxt = "取消锁定" : ErrMsg = ""

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

	If Locked Then SubButTxt = "锁定"
	If ChkTable(tSheetName) = False Then ErrMsg = ErrMsg & "未找到数据表 " & tSheetName & "！<br>"
	If ErrMsg<>"" Then Response.Write GetErrBody(0) : Exit Sub

	Dim Count1 : Count1 = 0
	sqlTmp = "Select VA1 From " & tSheetName & " GROUP BY VA1"
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		Count1 = rsTmp.Recordcount	'计算生成总数
	Set rsTmp = Nothing


	strHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />" & vbCrlf
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.ImportTips {text-align:center;line-height:50px;}" & vbCrlf
	strHtml = strHtml & "		.ImportTips b {color:#f30}" & vbCrlf
	strHtml = strHtml & "		.layui-btn-disabled {background:#ddd;color:#aaa}" & vbCrlf
	
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", Page_Title)
	strHeadHtml = Replace(strHeadHtml, "[@Head_style]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "			var layer = layui.layer, element = layui.element;" & vbCrlf
	strHtml = strHtml & "			layer.config({skin:""layer-hr""});var loadInit = layer.load();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf

	strHeadHtml = Replace(strHeadHtml, "[@Head_script]", strHtml)
	Response.Write strHeadHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field site-demo-button"" style=""margin-top:1px;"">" & vbCrlf
	Response.Write "		<legend>" & SubButTxt & " " & tItemName & " 业绩记录</legend>" & vbCrlf
	Response.Write "		<div class=""hr-shrink-x10"">" & vbCrlf
	Response.Write "			<div><button class=""layui-btn layui-btn-sm"" id=""LockUpdate"" title=""更新""><i class=""hr-icon hr-icon-top"">&#xeeaa;</i>开始</button></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""hr-shrink-x10"">" & vbCrlf
	Response.Write "			<div class=""layui-progress layui-progress-big"" lay-showpercent=""true"" lay-filter=""demo"">" & vbCrlf
	Response.Write "				<div class=""layui-progress-bar layui-bg-red"" lay-percent=""0%""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""ImportTips"" id=""ImportTips"">" & SubButTxt & "操作可能会持续约 " & Clng((Count1*0.33)/60) & " 分钟　<b>请点击“开始”按钮</b></div>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf

	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table, form = layui.form;" & vbCrlf
	strHtml = strHtml & "		element = layui.element; layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "		$(""#LockUpdate"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			$(this).addClass(""layui-btn-disabled"");$(this).html(""<i class='hr-icon hr-icon-top'>&#xeeb7;</i>更新中…"");" & vbCrlf
	strHtml = strHtml & "			layer.load(1);" & vbCrlf

	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	function updateItemKPI(iNum, itemid, islock){" & vbCrlf			'更新栏目KPI
	strHtml = strHtml & "		$.getJSON(""" & ParmPath & "Course/SwitchLock.html"",{ItemID:itemid, Locked:islock}, function(reData){" & vbCrlf

	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	}" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@Foot_script]", strHtml)
	Response.Write strFootHtml

End Sub

Sub Affirm()			'老师确认课程业绩正确
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim arrTmpID, tmpID : tmpID = Trim(ReplaceBadChar(Request("ID")))
	tmpID = FilterArrNull(tmpID, ",")

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
	Dim tmpStuType, tYGXM, tYGDM, tPXXH, sendUserID

	If ChkTable(tSheetName) = False Then		'检查数据表是否存在
		ErrMsg = ErrMsg & "数据表 " & tSheetName & " 不存在！<br>"
	End If
	If ErrMsg <> "" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}"
		Exit Sub
	End If

	Dim SentMsg, iRs
	If UserYGDM <> "" And UserID=0 Then
		If HR_IsNull(tmpID) Then
			ErrMsg = "您还没有选择课程业绩！<br>"
		Else
			arrTmpID = Split(tmpID, ",")
			iRs = 0
			For m = 0 To Ubound(arrTmpID)
				Set rsTmp = Server.CreateObject("ADODB.RecordSet")
					rsTmp.Open("Select * From " & tSheetName & " Where Passed=" & HR_True & " And ID=" & HR_Clng(arrTmpID(m))), Conn, 1, 3
					If Not(rsTmp.BOF And rsTmp.EOF) Then
						tmpStuType = Trim(rsTmp("StudentType"))
						tYGXM = Trim(rsTmp("VA2"))
						tYGDM = HR_Clng(rsTmp("VA1"))
						tPXXH = HR_Clng(rsTmp("VA0"))
						sendUserID = HR_Clng(rsTmp("UserID"))
						rsTmp("State") = 1			'确认课程业绩（教师个人）
						rsTmp.Update
						iRs = iRs + 1
					Else
						ErrMsg = tItemName & "课程业绩不存在！【ID：" & arrTmpID(m) & "】<br>"
						Exit For
					End If
				Set rsTmp = Nothing


				ErrMsg = UserYGXM & "老师已确认本条课程业绩内容，考核项目：" & tItemName & "，序号：" & tPXXH & "，教师" & UserYGDM & "[工号 " & UserYGXM & "]，时间：" & FormatDate(Now(), 1)
				ErrMsg = ErrMsg & " <span class=""ShowCourse"" data-ItemID=""" & tItemID & """ data-id=""" & arrTmpID(m) & """ data-sender=""" & UserYGDM & """>【查看】</span>"
				
				'If sendUserID > 0 Then		'发送给上传者【去掉注释可恢复给管理员发送信息】
				'	SentMsg = SendMessage(1, sendUserID, UserYGXM & "老师已确认课程业绩", ErrMsg, 0)
				'Else		'发送给所有管理员
				'	If HR_IsNull(tmpStuType) = False Then
				'		SendMsgManager = GetManagerID(tmpStuType, 0)
				'		tArrSender = Split(SendMsgManager, ",")
				'	Else
				'		tArrSender = arrManager
				'	End If
				'	For i = 0 To Ubound(tArrSender)
				'		SentMsg = SendMessage(1, tArrSender(i), UserYGXM & "老师已确认课程业绩", ErrMsg, 0)
				'	Next
				'End If
			Next
			ErrMsg = "您选择中的课程业绩有" & iRs & "条已确认提交！<br>但不包括未审核的数据！"
		End If
	Else
		ErrMsg = "管理员 " & UserName & " ，请用工号登陆后再试！<br>"
	End If
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作成功！""}"
End Sub
Sub oneAffirm()		'一键确认所有项目中的业绩数据
	If HR_Clng(UserYGDM) > 0 And UserID=0 Then
		Dim rsClass, tSheetName
		Set rsClass = Conn.Execute("Select * From HR_Class")
			If Not(rsClass.BOF And rsClass.EOF) Then
				Do While Not rsClass.EOF
					tSheetName = "HR_Sheet_" & rsClass("ClassID")
					If ChkTable(tSheetName) Then
						Conn.Execute("Update " & tSheetName & " Set State=1 Where State=0 And Passed=" & HR_True & " And VA1=" & HR_Clng(UserYGDM))
					End If
					rsClass.MoveNext
				Loop
				ErrMsg = "您所有考核项目的课程业绩已确认提交！<br>但不包括未审核的数据！"
			Else
				ErrMsg = "没有考核项目！<br>"
			End If
		Set rsClass = Nothing
		
	Else
		ErrMsg = "管理员 " & UserName & " ，请用工号登陆后再试！<br>"
	End If
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作成功！""}"
End Sub

Sub backModify()	'退回修改
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tYGDM


	If tItemID = 0 Or tmpID = 0 Then
		ErrMsg = "您没有选择课程业绩！【ID：" & tmpID & "/" & tItemID & "】<br>"
	End If
	If HR_IsNull(ErrMsg) = False Then Response.Write GetErrBody(0) : Exit Sub

	strHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />" & vbCrlf
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	'strHtml = strHtml & "		.layui-btn-disabled {background:#ddd;color:#aaa}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", Page_Title)
	strHeadHtml = Replace(strHeadHtml, "[@Head_style]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "			var layer = layui.layer, element = layui.element;" & vbCrlf
	strHtml = strHtml & "			layer.config({skin:""layer-hr""});var loadInit = layer.load();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf

	strHeadHtml = Replace(strHeadHtml, "[@Head_script]", strHtml)
	Response.Write strHeadHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title"" style=""margin-top:1px;"">" & vbCrlf
	Response.Write "		<legend>退回课程业绩</legend>" & vbCrlf

	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "	<form class=""layui-form layui-form-pane"" id=""backForm"" name=""backForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layui-form-item layui-form-text"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">退回理由：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><textarea name=""Explain"" id=""Explain"" placeholder=""备注"" class=""layui-textarea""></textarea></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<input name=""ItemID"" type=""hidden"" value=""" & tItemID & """><input name=""ID"" type=""hidden"" value=""" & tmpID & """>" & vbCrlf
	Response.Write "		<input name=""ygdm"" type=""hidden"" value=""" & tYGDM & """><input name=""userid"" type=""hidden"" value=""" & UserID & """>" & vbCrlf
	Response.Write "		<div class=""searchBtn"">" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-sm"" type=""button"" id=""backPost"" title=""发送""><i class=""hr-icon hr-icon-top"">&#xec58;</i>发送</button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var form = layui.form;" & vbCrlf
	strHtml = strHtml & "		element = layui.element; layer.closeAll(""loading"");" & vbCrlf

	strHtml = strHtml & "		$(""#backPost"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			if($(""#Explain"").val()==""""){layer.alert(""退回理由没有填写"",{btn:""关闭"",icon:2});return false;}" & vbCrlf
	strHtml = strHtml & "			$(this).addClass(""layui-btn-disabled"");" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "CourseProof/backSave.html"", $(""#backForm"").serialize(), function(reData){" & vbCrlf
	strHtml = strHtml & "				var icon=1;if(!reData.Return){icon=2};" & vbCrlf
	strHtml = strHtml & "				layer.alert(reData.reMessge, {btn:""关闭"",icon:icon}, function(){" & vbCrlf
	strHtml = strHtml & "					if(reData.Return){parent.layer.closeAll();parent.location.reload();}else{layer.closeAll();$(""#backPost"").removeClass(""layui-btn-disabled"");}" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@Foot_script]", strHtml)
	Response.Write strFootHtml

End Sub
Sub backSave()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim arrID, tmpID : tmpID = Trim(ReplaceBadChar(Request("ID")))
	tmpID = FilterArrNull(tmpID, ",")
	Dim tYGDM : tYGDM = HR_Clng(Request("ygdm"))

	Dim tItemName, tTemplate, lenField, tFieldHead, tSheetName
	sqlTmp = "Select * From HR_Class Where ClassID=" & tItemID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			tSheetName = "HR_Sheet_" & tItemID
		Else
			ErrMsg = "业绩考核项目" & tItemName & "不存在！<br>"
		End If
	Set rsTmp = Nothing
	If ChkTable(tSheetName) = False Then ErrMsg = ErrMsg & "数据表 " & tSheetName & " 不存在！<br>"	'检查数据表是否存在
	If HR_IsNull(tmpID) Then ErrMsg = ErrMsg & "您没有选择业绩数据！<br>"	'检查数据表是否存在

	If HR_IsNull(ErrMsg) = False Then Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """}" : Exit Sub

	Dim tmpStuType, tYGXM, tPXXH, sendUserID, tTitle, tContent
	arrID = Split(tmpID, ",")
	Dim backMsg, n, forID, iFor : iFor = 0
	For n = 0 To Ubound(arrID)
		forID = HR_Clng(arrID(n))
		tmpStuType = "" : tYGXM = "" : tYGDM = 0 : tPXXH = 0 : sendUserID = 0
		Set rsTmp = Server.CreateObject("ADODB.RecordSet")
			rsTmp.Open("Select * From " & tSheetName & " Where ID=" & forID), Conn, 1, 3
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				tmpStuType = Trim(rsTmp("StudentType"))
				tYGXM = Trim(rsTmp("VA2"))
				tYGDM = HR_Clng(rsTmp("VA1"))
				tPXXH = HR_Clng(rsTmp("VA0"))
				sendUserID = HR_Clng(rsTmp("UserID"))
				If HR_Clng(rsTmp("UserID")) = 0 And UserID>0 Then rsTmp("UserID") = UserID		'当管理员未分配时
				rsTmp("State") = 0				'退回（教师个人）
				rsTmp("Retreat") = 1			'变更退回状态
				If HR_CBool(rsTmp("Passed")) Then		'若已经审核则改为未审，同时更新业绩分
					rsTmp("Passed") = HR_False
					Call UpdateTeacherKPI(tItemID, rsTmp("VA1"), "")	'更新本项目员工统计数据
					Call UpdateTeacherTotalKPI(rsTmp("VA1"))			'更新员工总计数据
				End If
				rsTmp.Update
				iFor = iFor + 1
				
			
				backMsg = "【退回】您在" & tItemName & "中的课程业绩有误，请修改！[序号：" & tPXXH & "]<br>" & Request("Explain") & ""
				Call SendMessage(2, tItemID, forID, tYGDM, "您在" & tItemName & "中的课程业绩有误，请修改！", backMsg, "")		'发送站内消息
				'发送消息到企业微信提醒！【所有管理员】

				tTitle = "【退回】您在" & tItemName & "的课程已被退回！"
				tContent = HR_HtmlDecode(Trim(Request("Explain"))) : tContent = Replace(nohtml(tContent), " ", "") : tContent = Replace(nohtml(tContent), "&nbsp;", "") : tContent = GetSubStr(tContent, 110, True)
				tContent = "发送时间：" & FormatDate(Now(), 10) & "<br>" & tContent
				tContent = Replace(tContent, "</p><p>", "<br>")
				backMsg = SentWechatMSG_QYCard(tYGDM, tTitle, SiteUrl & "/Touch/Course/View.html?ItemID=" & tItemID & "&ID=" & forID & "", tContent)

			End If
		Set rsTmp = Nothing
	Next
	ErrMsg = "共" & iFor & "/" & Ubound(arrID) + 1 & "退回消息已经发送！"
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作成功！""}"
End Sub
%>