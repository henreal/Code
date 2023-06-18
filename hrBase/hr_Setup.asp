<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./incKPI.asp"-->
<!--#include file="./hr_SetupInc.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim arrCampus : arrCampus = Split(XmlText("Common", "Campus", ""), "|")
SiteTitle = "系统管理"
Dim SubButTxt : SubButTxt = "参数"

Dim sc4Json, jsonobj

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))				'Get head template code.
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))				'Get foot template code.
Dim strNavPath : strNavPath = ReplaceCommonLabel(getFrameNav(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

If UserRank < 1 Then
	'ErrMsg = "您没有 系统管理 的权限！"
	'Response.Write GetErrBody(0) : Response.End
End If

Select Case Action
	Case "Index" Call MainBody()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveForm" Call SaveForm()
	Case "PeriodData" Call PeriodData()
	Case "SavePeriod" Call SavePeriod()
	Case "SetupBody" Call SetupBody()
	Case "DelPeriod" Call DelPeriod()
	Case "Period" Call Period()

	Case "EditCampus" Call EditCampus()
	Case "CampusData" Call CampusData()
	Case "SaveCampus" Call SaveCampus()
	Case "DeleteCampus" Call DeleteCampus()

	Case "Course" Call CourseBody()
	Case "SaveCourse" Call SaveCourse()
	Case "EditCourse" Call EditCourse()
	Case "DeleteCourse" Call DeleteCourse()

	Case "TeachClass" Call TeachClassBody()
	Case "SaveTeachClass" Call SaveTeachClass()
	Case "DelTeachClass" Call DelTeachClass()

	Case "ClassRoom" Call ClassRoomBody()
	Case "SaveClassRoom" Call SaveClassRoom()
	Case "DelClassRoom" Call DelClassRoom()

	Case "ImportGrade" Call ImportGrade()		'等级导入
	Case "ImportSave" Call ImportSave()

	Case "SetupSwitch" Call SetupSwitch()
	Case "SaveSwitch" Call SaveSwitch()
	Case "SaveSwitchImport" Call SaveSwitchImport()
	Case "SaveYear" Call SaveYear()
	Case "BackData" Call BackBody()
	Case "getBackupData" Call BackupData()		'执行备份
	Case "GetBackdataJson" Call GetBackdataJson()	'取已备份数据列表
	Case Else Response.Write GetErrBody(0)
End Select

Sub Period()
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .width_100 {width:100px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-label {padding:9px;} .hr-pop-fix {position: absolute;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layer-hr-box {border:1px solid #ffc107;box-sizing: border-box;margin:0;padding:10px 10px 0}" & vbCrlf
	tmpHtml = tmpHtml & "		body{overflow-y: scroll;}" & vbCrlf
	tmpHtml = tmpHtml & "		.filter-box {border:1px solid #eee;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.filter-tit {border-bottom:1px solid #eee;padding:10px;background-color:#f3f3f3;font-weight: bold;}" & vbCrlf
	tmpHtml = tmpHtml & "		.filter-list {box-sizing: border-box;padding:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	strHtml = strHtml & getFrameNav(1)
	tmpHtml = "<a href=""" & ParmPath & "Setup/Index.html"">" & SiteTitle & "</a><a><cite>校区节次管理</cite></a>"
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layer-hr-box"">" & vbCrlf
	Response.Write "	<form class=""layui-form soBox"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layui-inline"">校(院)区：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><select name=""Campus"" lay-verify=""required"">"
	For i = 0 To Ubound(arrCampus)
		Response.Write "<option value=""" & arrCampus(i) & """"
		If i=0 Then Response.Write " selected"
		Response.Write ">" & arrCampus(i) & "</option>"
	Next
	Response.Write "</select>"
	Response.Write "</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline"">　节次：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline width_100""><input type=""number"" name=""Period"" value="""" placeholder=""请输入节次"" lay-verify=""number"" autocomplete=""off"" class=""layui-input""></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline"">　时间：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input type=""text"" name=""Time"" id=""StartTime"" value="""" placeholder=""时间"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><button class=""layui-btn"" lay-submit lay-filter=""SubPost"">添加节次</button></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><button class=""layui-btn layui-btn-normal"" id=""AddCampus"">添加校(院)区</button></div>" & vbCrlf
	Response.Write "		<input type=""hidden"" name=""ID"" value=""""><input type=""hidden"" name=""Modify"" value=""True"">" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form filter-box""><div class=""filter-tit"">筛选：</div><div class=""layui-row filter-list"">"
	For i = 0 To Ubound(arrCampus)
		Response.Write "<em class=""layui-col-xs5 layui-col-sm4 layui-col-md3""><input type=""checkbox"" name=""Campus"" lay-skin=""primary"" class=""CheckFilter"" title=""" & arrCampus(i) & """ value=""" & arrCampus(i) & """></em>"
	Next
	Response.Write "</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "Setup/PeriodData.html',limit:20,page:true,text:{none:'暂时未添加节次'},id:'TableList'}"" lay-filter=""TableList"">" & vbCrlf
	Response.Write "	<thead><tr>" & vbCrlf
	Response.Write "		<th lay-data=""{field:'Campus'}"">校(院)区</th>" & vbCrlf
	Response.Write "		<th lay-data=""{field:'Period',unresize:true,align:'center',sort: true,width:80}"">节次</th>" & vbCrlf
	Response.Write "		<th lay-data=""{field:'StartTime',unresize:true,width:80}"">开始时间</th>" & vbCrlf
	Response.Write "		<th lay-data=""{field:'EndTime',unresize:true,width:80}"">结束时间</th>" & vbCrlf
	Response.Write "		<th lay-data=""{align:'center',unresize:true,width:110, toolbar: '#barTable'}"">操作</th>" & vbCrlf
	Response.Write "		</tr></thead>" & vbCrlf
	Response.Write "	</table>" & vbCrlf
	Response.Write "<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "	<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "		<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	Response.Write "		<a class=""layui-btn layui-btn-sm layui-btn-danger"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</script>" & vbCrlf

	Dim xlsFile, strTmp, arr0
	xlsFile = "/Upload/Data_A3_2.xls"
	strTmp = GetHttpPage(apiHost & "/Manage/ReadExcel.htm?xlsFile=" & xlsFile, 1)

	Response.Write "</div>" & vbCrlf

	strHtml = "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""laydate"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table, form = layui.form, laydate = layui.laydate;" & vbCrlf
	strHtml = strHtml & "		element = layui.element;" & vbCrlf

	strHtml = strHtml & "		form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Setup/SavePeriod.html"", $(""#EditForm"").serialize(), function(result){" & vbCrlf
	strHtml = strHtml & "				var reData = eval(""("" + result + "")""), icon=2;" & vbCrlf
	strHtml = strHtml & "				if(reData.Return){icon=1;}" & vbCrlf
	strHtml = strHtml & "				layer.alert(reData.reMessge, {icon:icon},function(layero, index){" & vbCrlf
	strHtml = strHtml & "					if(reData.Return){layer.closeAll();form.render();table.reload(""TableList"");}" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	strHtml = strHtml & "			var data = obj.data;" & vbCrlf
	strHtml = strHtml & "			if(obj.event === ""del""){" & vbCrlf
	strHtml = strHtml & "				layer.confirm('真的删除选中的节次吗？', {icon:3,title: ""删除提醒""}, function(index){" & vbCrlf
	strHtml = strHtml & "					$.getJSON(""" & ParmPath & "Setup/DelPeriod.html"",{ID:data.PeriodID}, function(reData){" & vbCrlf
	strHtml = strHtml & "						if(reData.Return){;" & vbCrlf
	strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:1,title: ""系统提示""},function(layero, index){layer.closeAll();table.reload(""TableList"");});" & vbCrlf
	strHtml = strHtml & "						}else{" & vbCrlf
	strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:2,title: ""系统提示""});" & vbCrlf
	strHtml = strHtml & "						}" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""edit""){" & vbCrlf
	strHtml = strHtml & "				var loadTips = layer.load(1);" & vbCrlf
	strHtml = strHtml & "				layer.open({type:1,id:""popBody"",title:[""编辑节次"",""font-size:16""],area:[""700px"", ""360px""]});" & vbCrlf
	strHtml = strHtml & "				$.get(""" & ParmPath & "Setup/Edit.html"",{ID:data.PeriodID,Eve:obj.event}, function(strForm){" & vbCrlf
	strHtml = strHtml & "					$(""#popBody"").html(strForm);form.render();" & vbCrlf
	strHtml = strHtml & "					laydate.render({" & vbCrlf
	strHtml = strHtml & "						elem: ""#PeriodTime"",type: ""time"",range: ""-"",theme: '#060',format: ""HH:mm""" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "				$(""#popBody"").niceScroll();" & vbCrlf
	strHtml = strHtml & "				layer.close(loadTips);" & vbCrlf
	strHtml = strHtml & "				form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "					$.post(""" & ParmPath & "Setup/SavePeriod.html"",PostData.field, function(result){" & vbCrlf
	strHtml = strHtml & "						var reData = eval(""("" + result + "")"");" & vbCrlf
	strHtml = strHtml & "						if(reData.Return){" & vbCrlf
	strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:1},function(layero, index){layer.closeAll();table.reload(""TableList"");});" & vbCrlf
	strHtml = strHtml & "						}else{" & vbCrlf
	strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:2});" & vbCrlf
	strHtml = strHtml & "						}" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "					return false;" & vbCrlf
	strHtml = strHtml & "					" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		form.on(""checkbox"", function(data){" & vbCrlf
	strHtml = strHtml & "			var strArrID = """";" & vbCrlf
	strHtml = strHtml & "			$("".CheckFilter"").each(function(){" & vbCrlf
	strHtml = strHtml & "				if($(this).is("":checked""))strArrID += $(this).attr(""value"")+"","";" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			table.reload(""TableList"", {" & vbCrlf
	strHtml = strHtml & "				url:""" & ParmPath & "Setup/PeriodData.html"",where:{arr:strArrID}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		laydate.render({" & vbCrlf
	strHtml = strHtml & "			elem: ""#StartTime"",type: ""time"",range: ""-"",theme: '#060',format: ""HH:mm""" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#AddCampus"").on(""click"", function(index){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:2,id:""popWin"",content:""" & ParmPath & "Setup/EditCampus.html"",title:[""校(院)区管理"",""font-size:16""],area:[""700px"", ""550px""]});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "	});" & vbCrlf
	'strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub

Sub InitScriptControl()
	Set sc4Json = Server.CreateObject("MSScriptControl.ScriptControl")
		sc4Json.Language = "JavaScript"
		sc4Json.AddCode "var itemTemp=null;function getJSArray(arr, index){itemTemp=arr[index];}"
End Sub

Function getJSONObject(strJSON)
	sc4Json.AddCode "var jsonObject = " & strJSON
	Set getJSONObject = sc4Json.CodeObject.jsonObject
End Function

Sub getJSArrayItem(objDest, objJSArray, index)
	On Error Resume Next
	sc4Json.Run "getJSArray",objJSArray, index
	Set objDest = sc4Json.CodeObject.itemTemp
	If Err.number=0 Then
		Exit Sub
	End If
	objDest = sc4Json.CodeObject.itemTemp
End Sub



Sub MainBody()

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	
	tmpHtml = "<a href=""" & ParmPath & "Setup/Index.html"">" & SiteTitle & "</a><a><cite>首页</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones"">" & vbCrlf
	Response.Write "	<div class=""layui-tab layui-tab-brief"" lay-filter=""SetupTab"">" & vbCrlf
	Response.Write "		<ul class=""layui-tab-title"">" & vbCrlf
	Response.Write "			<li class=""layui-this"" name=""Base"">基本参数</li>" & vbCrlf
	Response.Write "			<li name=""Rank"">管理级别</li>" & vbCrlf
	Response.Write "			<li name=""AddArea"">校区设置</li>" & vbCrlf
	Response.Write "		</ul>" & vbCrlf
	Response.Write "		<div class=""layui-tab-content tabBox"">" & vbCrlf
	Response.Write "			<div class=""layui-tab-item layui-show"">基本参数</div>" & vbCrlf
	Response.Write "			<div class=""layui-tab-item"">管理级别</div>" & vbCrlf
	Response.Write "			<div class=""layui-tab-item"">校区设置"
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf

	strHtml = strHtml & "	layui.use([""table"", ""form"", ""laydate"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table, laydate = layui.laydate;" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf

	strHtml = strHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	strHtml = strHtml & "			var data = obj.data;" & vbCrlf
	strHtml = strHtml & "			if(obj.event === ""del""){" & vbCrlf

	strHtml = strHtml & "			}else if(obj.event === ""edit""){" & vbCrlf


	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "	});" & vbCrlf
	'strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub

Sub SetupBody()
	Dim tabIndex : tabIndex = HR_Clng(Request("index"))
	strHtml = "<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "Department/AllData.html',id:'TableList'}"" lay-filter=""TableList"">" & vbCrlf
	strHtml = strHtml & "	<thead><tr>" & vbCrlf
	strHtml = strHtml & "		<th lay-data=""{fixed:'left',type:'checkbox'}""></th>" & vbCrlf
	strHtml = strHtml & "		<th lay-data=""{field:'KSDM',unresize:true, width:80}"">科室代码</th>" & vbCrlf
	strHtml = strHtml & "		<th lay-data=""{field:'KSMC',width:180}"">科室名称</th>" & vbCrlf
	strHtml = strHtml & "		</tr></thead>" & vbCrlf
	strHtml = strHtml & "</table>" & vbCrlf
	strHtml = strHtml & "" & vbCrlf
	Response.Write strHtml
End Sub

Sub PeriodData()

	Dim vCount, vMSG, tmpJson, tmpTime, rsGet, sqlGet, tmpData
	Dim strCampus, arrCampus : strCampus = Trim(ReplaceBadChar(Request("arr")))
	strCampus = FilterArrNull(strCampus, ",")
	If strCampus<> "" Then
		arrCampus = Split(strCampus, ",")
		sqlGet = "Where "
		For i = 0 To Ubound(arrCampus)
			If i>0 Then sqlGet = sqlGet & " Or"
			sqlGet = sqlGet & " Campus='" & arrCampus(i) & "'"
		Next
	End If

	sqlGet = "Select * From HR_Period " & sqlGet & " Order By PeriodID ASC"
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			i = 0
			CurrentPage = 1
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
			Do While Not rsGet.EOF
				If i > 0 Then tmpData = tmpData & ","
				tmpData = tmpData & "{""PeriodID"":" & rsGet("PeriodID") & ""
				tmpData = tmpData & ",""Campus"":""" & ReplaceAPIStr(rsGet("Campus")) & """,""Period"":""" & HR_Clng(rsGet("Period")) & """"
				tmpData = tmpData & ",""StartTime"":""" & Trim(rsGet("StartTime")) & """,""EndTime"":""" & Trim(rsGet("EndTime")) & """"
				tmpData = tmpData & "}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
		vCount = rsGet.Recordcount
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""暂无数据"",""count"":" & vCount & ",""data"":[" & tmpData
	tmpJson = tmpJson & "],""limit"":""0"",""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub

Sub EditBody()
	Dim tmpHtml, tmpID, rsEdit, tmpTime, tCampus, tPeriod, tExplain
	tmpID = HR_Clng(Request("ID"))
	Set rsEdit = Conn.Execute("Select * From HR_Period Where PeriodID=" & tmpID )
		If Not(rsEdit.BOF And rsEdit.EOF) Then
			tCampus = Trim(rsEdit("Campus"))
			tPeriod = HR_Clng(rsEdit("Period"))
			tmpTime = Trim(rsEdit("StartTime")) & " - " & Trim(rsEdit("EndTime"))
			tExplain = Trim(rsEdit("Explain"))
		End If
	Set rsEdit = Nothing
	tmpHtml = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	tmpHtml = tmpHtml & "	<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<label class=""layui-form-label"">校(院)区：</label>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-inline""><select name=""Campus"" lay-verify=""required"">"
	For i = 0 To Ubound(arrCampus)
		tmpHtml = tmpHtml & "<option value=""" & arrCampus(i) & """"
		If tCampus = arrCampus(i) Then tmpHtml = tmpHtml & " selected"
		tmpHtml = tmpHtml & ">" & arrCampus(i) & "</option>"
	Next
	tmpHtml = tmpHtml & "</select></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "		<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-inline""><label class=""layui-form-label"">节　次：</label>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-input-inline""><input type=""number"" name=""Period"" value=""" & tPeriod & """ placeholder=""请输入节次"" lay-verify=""number"" autocomplete=""off"" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-inline""><label class=""layui-form-label"">时　间：</label>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-input-inline""><input type=""text"" name=""Time"" id=""PeriodTime"" value=""" & tmpTime & """ placeholder=""时间"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "		<input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""Modify"" value=""True"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""hr-pop-fix"">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-inline""><button class=""layui-btn layui-btn-sm"" lay-submit lay-filter=""SubPost""><i class=""hr-icon"">&#xf0c7;</i>保存</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</form>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	Response.Write tmpHtml
End Sub
Sub SavePeriod()
	Dim tmpID : tmpID = HR_Clng(Request("ID")) : ErrMsg = ""
	Dim sqlAdd, rsAdd, tmpJson, StartTime, EndTime, arrTime : SubButTxt = "修改"
	Dim tmpTime : tmpTime = Trim(Request("Time"))
	If Instr(tmpTime, "-") > 0 Then
		arrTime = Split(tmpTime, "-")
		StartTime = arrTime(0) : EndTime = arrTime(1)
	Else
		StartTime = "" : EndTime = ""
	End If

	sqlAdd = "Select * From HR_Period Where PeriodID=" & tmpID
	Set rsAdd = Server.CreateObject("ADODB.RecordSet")
		rsAdd.Open(sqlAdd), Conn, 1, 3
		If rsAdd.BOF And rsAdd.EOF Then
			rsAdd.AddNew
			rsAdd("PeriodID") = GetNewID("HR_Period", "PeriodID")
			SubButTxt = "添加"
		End If
		rsAdd("Campus") = Trim(Request.Form("Campus"))
		rsAdd("Period") = HR_Clng(Request.Form("Period"))
		rsAdd("StartTime") = Trim(StartTime)
		rsAdd("EndTime") = Trim(EndTime)
		rsAdd("Explain") = Trim(Request.Form("Explain"))
		rsAdd.Update
		rsAdd.Close
	Set rsAdd = Nothing

	If ErrMsg <> "" Then
		tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}"
	Else
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""节次 " & Request.Form("Campus") & " " & SubButTxt & "成功！"",""ReStr"":""操作成功！""}"
	End If
	Response.Write tmpJson
End Sub

Sub DelPeriod()
	If UserRank <> 2 Then
		ErrMsg = "{""Return"":false,""Err"":400,""reMessge"":""您没有删除节次权限"",""ReStr"":[]}"
		Response.Write ErrMsg : Exit Sub
	End If

	Dim tmpJson, rsDel, sqlDel, strDel, arrDel, iDel, tmpErr
	strDel = Trim(ReplaceBadChar(Request("ID")))
	strDel = DelRightComma(strDel)
	arrDel = Split(strDel, ",")
	iDel = 0
	For i = 0 To Ubound(arrDel)
		Set rsDel = Server.CreateObject("ADODB.RecordSet")
			rsDel.Open("Select * From HR_Period Where PeriodID=" & HR_Clng(arrDel(i))), Conn, 1, 3
			If Not(rsDel.BOF And rsDel.EOF) Then
				rsDel.Delete
				iDel = iDel + 1
				rsDel.Close
			End If
		Set rsDel = Nothing
	Next
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & iDel & "/" & Ubound(arrDel) + 1 & " 条记录删除成功！" & tmpErr & """,""ReStr"":""操作成功！""}"
	Response.Write tmpJson
End Sub

Sub EditCampus()
	Dim tmpJson
	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table thead tr {background-color: #eee;color: #000;text-align:center;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")

	tmpJson = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	tmpJson = tmpJson & "<div class=""hr-pop-fix""><button class=""layui-btn layui-btn-sm layui-btn-normal"" id=""AddBtn"" title=""添加校(院)区""><i class=""layui-icon"">&#xe654;</i>添加</button><button class=""layui-btn layui-btn-sm"" id=""refresh"" title=""刷新""><i class=""hr-icon"">&#xf343;</i></button></div>" & vbCrlf
	tmpJson = tmpJson & "<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "Setup/CampusData.html',limit:0,text:{none:'校(院)区尚未添加'},id:'EditTable'}"" lay-filter=""EditTable"">" & vbCrlf
	tmpJson = tmpJson & "	<thead><tr>" & vbCrlf
	tmpJson = tmpJson & "		<th lay-data=""{type:'numbers',unresize:true,align:'center',width:60}"">序号</th>" & vbCrlf
	tmpJson = tmpJson & "		<th lay-data=""{field:'Campus',edit:'text'}"">校(院)区</th>" & vbCrlf
	tmpJson = tmpJson & "		<th lay-data=""{align:'center',unresize:true,width:90,toolbar: '#EditBar'}"">操作</th>" & vbCrlf
	tmpJson = tmpJson & "	</tr></thead>" & vbCrlf
	tmpJson = tmpJson & "</table>" & vbCrlf
	tmpJson = tmpJson & "<script type=""text/html"" id=""EditBar"">" & vbCrlf
	tmpJson = tmpJson & "	<div class=""layui-btn-group"">" & vbCrlf
	tmpJson = tmpJson & "		<a class=""layui-btn layui-btn-sm layui-btn-danger"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	tmpJson = tmpJson & "	</div>" & vbCrlf
	tmpJson = tmpJson & "</script>" & vbCrlf
	tmpJson = tmpJson & "</div>" & vbCrlf
	tmpJson = tmpJson & "<div class=""hr-place-h50""></div>" & vbCrlf
	strHtml = strHtml & tmpJson

	tmpJson = "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpJson = tmpJson & "<script type=""text/javascript"">" & vbCrlf
	tmpJson = tmpJson & "	$(document).ready(function(){});" & vbCrlf
	tmpJson = tmpJson & "	layui.use([""table"", ""laytpl"", ""form"", ""element""], function(){" & vbCrlf
	tmpJson = tmpJson & "		var table = layui.table, laytpl = layui.laytpl;" & vbCrlf
	tmpJson = tmpJson & "		element = layui.element, form = layui.form;" & vbCrlf
	tmpJson = tmpJson & "		$(""#AddBtn"").on(""click"", function(index){" & vbCrlf
	tmpJson = tmpJson & "			layer.prompt({title:""请输入校(院)区""},function(value, index, elem){" & vbCrlf
	tmpJson = tmpJson & "				$.getJSON(""" & ParmPath & "Setup/SaveCampus.html"",{Campus:value}, function(reData){});" & vbCrlf
	tmpJson = tmpJson & "				table.reload(""EditTable"");layer.close(index);" & vbCrlf
	tmpJson = tmpJson & "			});" & vbCrlf
	tmpJson = tmpJson & "			return false;" & vbCrlf
	tmpJson = tmpJson & "		});" & vbCrlf
	tmpJson = tmpJson & "		$(""#refresh"").on(""click"", function(index){window.location.reload();});" & vbCrlf
	tmpJson = tmpJson & "		table.on(""tool(EditTable)"", function(obj){" & vbCrlf
	tmpJson = tmpJson & "			var data = obj.data;" & vbCrlf
	tmpJson = tmpJson & "			if(obj.event === ""del""){" & vbCrlf
	tmpJson = tmpJson & "				layer.confirm('真的删除该校(院)区吗？<br />删除后无法恢复！', {icon:3,title: ""删除提醒""}, function(index){" & vbCrlf
	tmpJson = tmpJson & "					$.getJSON(""" & ParmPath & "Setup/DeleteCampus.html"",{Campus:data.Campus}, function(reData){ });" & vbCrlf
	tmpJson = tmpJson & "					window.location.reload();layer.close(index);" & vbCrlf
	tmpJson = tmpJson & "				});" & vbCrlf
	tmpJson = tmpJson & "			}" & vbCrlf
	tmpJson = tmpJson & "		});" & vbCrlf
	tmpJson = tmpJson & "		table.on(""edit(EditTable)"", function(obj){" & vbCrlf
	tmpJson = tmpJson & "			$.getJSON(""" & ParmPath & "Setup/SaveCampus.html"",{ID:obj.data.ID,Campus:obj.value}, function(reData){ });" & vbCrlf
	tmpJson = tmpJson & "			table.reload(""EditTable"");" & vbCrlf
	tmpJson = tmpJson & "		});" & vbCrlf
	tmpJson = tmpJson & "		layer.closeAll(""loading"");" & vbCrlf
	tmpJson = tmpJson & "	});" & vbCrlf
	tmpJson = tmpJson & "</script>" & vbCrlf

	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpJson)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub
Sub CampusData()
	Dim vCount, vMSG, tmpJson
	tmpJson = ""
	For i = 0 To Ubound(arrCampus)
		If i > 0 Then tmpJson = tmpJson & ","
		tmpJson = tmpJson & "{""ID"":" & i + 1 & ""
		tmpJson = tmpJson & ",""Campus"":""" & Trim(arrCampus(i)) & """"
		tmpJson = tmpJson & "}"
	Next
	vCount = Ubound(arrCampus) + 1

	tmpJson = "{""code"":0,""msg"":""暂无数据"",""count"":" & vCount & ",""data"":[" & tmpJson
	tmpJson = tmpJson & "],""limit"":""0"",""page"":""0""}"
	Response.Write tmpJson
End Sub
Sub SaveCampus()
	Dim tmpCampus, tCampus, tResult, tmpJson, strTmp
	Dim tmpID :tmpID = HR_Clng(Request("ID"))
	arrCampus = Split(XmlText("Common", "Campus", ""), "|")

	tCampus = Trim(ReplaceBadChar(Request("Campus")))
	tmpCampus = XmlText("Common", "Campus", "")
	tResult = False
	If tmpID > 0 Then
		For i = 0 To Ubound(arrCampus)
			If tmpID = i + 1 And Trim(tCampus) <> "" Then
				strTmp = strTmp & Trim(tCampus) & "|"
			Else
				strTmp = strTmp & Trim(arrCampus(i)) & "|"
			End If
		Next
		strTmp = FilterArrNull(strTmp, "|")
		tResult = UpdateXmlText("Common", "Campus", strTmp)
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""校(院)区 修改成功！"",""ReStr"":""操作成功！""}"
	Else
		If Instr(tmpCampus, "|") > 0 And tCampus <> "" Then
			tmpCampus = tmpCampus & "|" & tCampus
			tResult = UpdateXmlText("Common", "Campus", tmpCampus)
			tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""校(院)区 添加成功！"",""ReStr"":""操作成功！""}"
		Else
			tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""校(院)区 添加失败！"",""ReStr"":""操作失败！""}"
		End If
	End If
	Response.Write tmpJson
End Sub
Sub DeleteCampus()
	Dim tResult, tmpJson, tmpCampus : strTmp = ""
	Dim tCampus : tCampus = Trim(ReplaceBadChar(Request("Campus")))
	arrCampus = Split(XmlText("Common", "Campus", ""), "|")
	tResult = False

	For i = 0 To Ubound(arrCampus)
		If tCampus <> Trim(arrCampus(i)) Then strTmp = strTmp & Trim(arrCampus(i)) & "|"
	Next
	strTmp = FilterArrNull(strTmp, "|")
	If strTmp <> "" Then
		tResult = UpdateXmlText("Common", "Campus", strTmp)
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""校(院)区 删除成功！"",""ReStr"":""操作成功！""}"
	Else
		tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""校(院)区 删除失败！"",""ReStr"":""操作失败！""}"
	End If
	Response.Write tmpJson
End Sub
%>