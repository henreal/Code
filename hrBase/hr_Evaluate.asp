<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
SiteTitle = "课堂教学质量评价管理"
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "jsonList" Call GetJsonList()
	Case "List" Call List()
	Case "ListData" Call ListData()
	Case "Details" Call Details()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveEdit" Call SaveEdit()
	Case "Delete" Call Delete()
	Case "CourseOption" Call CourseOption()
	Case "GetCourseData" Call GetCourseData()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim layUrl : layUrl = ParmPath & "Evaluate/jsonList.html"

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-tablebtn .layui-btn {height:28px;line-height:28px;padding:0 10px;font-size:1.1rem;}" & vbCrlf		'表头工具集
	tmpHtml = tmpHtml & "		.sumbar b {color:#F60;padding:0 2px} .sumbar b.sumDebit{color:#080}" & vbCrlf		'表头汇总
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)

	tmpHtml = "<a href=""" & ParmPath & "Evaluate/Index.html"">" & SiteTitle & "</a><a><cite>列表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form name=""SearchForm"" id=""SearchForm"" class=""layui-form soBox"" action="""" method=""get""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input type=""text"" class=""layui-input"" name=""soWord"" value=""" & soWord & """ id=""soWord"" placeholder=""搜索评价人"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-normal"" data-event=""reload""><i class=""hr-icon hr-icon-top"">&#xeba1;</i>搜索</button></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn"">" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn hr-btn_deon"" data-event=""addnew"" name=""addnew"" title=""添加评价""><i class=""hr-icon"">&#xecfb;</i></button>" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn hr-btn_darkgreen"" data-event=""refresh"" name=""refresh"" title=""刷新本页""><i class=""hr-icon"">&#xf021;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""toolBtn"">" & vbCrlf	'表头模板
	Response.Write "		<div class=""hr-rows tpltools"">" & vbCrlf
	Response.Write "			<div class=""sumbar"">共<b class=""Count"">0</b>条记录</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf		'行工具
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""eval"" title=""查看评价""><i class=""hr-icon"">&#xefe2;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element"", ""form""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table; element = layui.element, form=layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",height:'full-115',page:true,limit:30,skin:'',limits:[10,15,20,30,50,100,200,300],toolbar: '#toolBtn'" & vbCrlf
	tmpHtml = tmpHtml & "			,text:{none:'暂时没有数据'},cols: [[" & vbCrlf				'设置表头
	tmpHtml = tmpHtml & "				{type:'numbers',title:'序号',width:60,align:'right'}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherCode',title:'工号',width:80}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Teacher',title:'评价人',width:105,event:'details',style:'cursor: pointer;'}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherDepart',title:'科室',minWidth:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherZC',title:'职称',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'CountEvalu',title:'次数',sort:true,align:'right',width:95,templet:function(res){ return res.CountEvalu + '次' } }" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'EduYear',title:'学年',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{title:'操作',align:'center',unresize:true,width:100, toolbar:'#barTable'}" & vbCrlf
	tmpHtml = tmpHtml & "			]],url:""" & layUrl & """,parseData: function(res){" & vbCrlf
	tmpHtml = tmpHtml & "				$("".Count"").text(res.count);" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".searchBtn button"").on(""click"", function(){" & vbCrlf	'搜索等按钮click事件
	tmpHtml = tmpHtml & "			var btnEvent = $(this).data(""event"");" & vbCrlf
	tmpHtml = tmpHtml & "			switch(btnEvent){" & vbCrlf
	tmpHtml = tmpHtml & "				case ""reload"":" & vbCrlf
	tmpHtml = tmpHtml & "					var arrForm = $(""#SearchForm"").serializeArray(), postStr={};" & vbCrlf
	tmpHtml = tmpHtml & "					$.each(arrForm, function(key, val){ postStr[this.name]=this.value; });" & vbCrlf		'表单序列转json
	tmpHtml = tmpHtml & "					table.reload(""layList"", {" & vbCrlf
	tmpHtml = tmpHtml & "						url:""" & layUrl & """,where: postStr" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""addnew"":" & vbCrlf
	tmpHtml = tmpHtml & "					layer.open({type:2, content:'" & ParmPath & "Evaluate/AddNew.html',title:[""添加课堂教学评价"",""font-size:16""],area:[""850px"",""90%""],moveOut:true,maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""refresh"":location.reload(); break;" & vbCrlf
	tmpHtml = tmpHtml & "			};" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf		'监听工具栏
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""eval""){" & vbCrlf
	tmpHtml = tmpHtml & "				location.href=""" & ParmPath & "Evaluate/List.html?ygdm="" + data.TeacherCode;" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub GetJsonList()				'
	Dim tmpJson, rsGet, sqlGet, tIntro
	Dim tLimit : tLimit = HR_Clng(Request("limit"))
	Dim tPage : tPage = HR_Clng(Request("page"))
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))

	sqlGet = "Select Count(0) As Count, a.ParticipantID, a.Participant"
	sqlGet = sqlGet & " ,(Select KSMC From HR_Teacher Where YGDM=a.ParticipantID) As KSMC"
	sqlGet = sqlGet & " ,(Select PRZC From HR_Teacher Where YGDM=a.ParticipantID) As PRZC"
	sqlGet = sqlGet & " From HR_Evaluate a Where a.ClassTime<'" & DefYear & "-06-30 23:59:59' And a.ClassTime>'" & DefYear - 1 & "-07-01 00:00:00'"
	If HR_Clng(soWord) > 1000 Then
		sqlGet = sqlGet & " And a.ParticipantID=" & soWord
	ElseIf HR_IsNull(soWord) = False Then
		sqlGet = sqlGet & " And a.Participant like '%" & soWord &"%'"
	End If
	sqlGet = sqlGet & " Group By a.ParticipantID, a.Participant"
	sqlGet = sqlGet & " Order By a.ParticipantID ASC"

	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			i = 0 : CurrentPage = 1 : MaxPerPage = tLimit
			If tPage > 0 Then CurrentPage = tPage
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
			Dim tSheetName, tTemplate, tItemID, tItemName
			Do While Not rsGet.EOF
				If i > 0 Then tmpJson = tmpJson & ","
				tmpJson = tmpJson & "{""Teacher"":""" & Trim(rsGet("Participant")) & """,""TeacherCode"":""" & HR_CLng(rsGet("ParticipantID")) & """,""TeacherDepart"":""" & Trim(rsGet("KSMC")) & """,""TeacherZC"":""" & Trim(rsGet("PRZC")) & """"
				tmpJson = tmpJson & ",""EduYear"":""" & DefYear-1 & "-" & DefYear & """,""CountEvalu"":" & HR_CLng(rsGet("Count")) & "}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""查询成功"",""count"":" & HR_Clng(TotalPut) & ",""data"":[" & tmpJson & "],""limit"":""" & MaxPerPage & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub

Sub List()
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim tYGDM : tYGDM = HR_Clng(Trim(Request("ygdm")))
	Dim layUrl : layUrl = ParmPath & "Evaluate/ListData.html?ygdm=" & tYGDM
	Dim tYGXM, tKSMC, tPRZC, tXZZW
	Set rs = Conn.Execute("Select * From HR_Teacher Where YGDM='" & tYGDM & "'")
		If Not(rs.BOF And rs.EOF) Then
			tYGXM = Trim(rs("YGXM"))
			tKSMC = Trim(rs("KSMC"))
			tPRZC = Trim(rs("PRZC"))
			tXZZW = Trim(rs("XZZW"))
		End If
	Set rs = Nothing
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-tablebtn .layui-btn {height:28px;line-height:28px;padding:0 10px;font-size:1.1rem;}" & vbCrlf		'表头工具集
	tmpHtml = tmpHtml & "		.sumbar b {color:#F60;padding:0 2px} .sumbar b.sumDebit{color:#080}" & vbCrlf		'表头汇总
	tmpHtml = tmpHtml & "		.t-title {width:100px;text-align:right;background-color:#eee}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)

	tmpHtml = "<a href=""" & ParmPath & "Evaluate/Index.html"">" & SiteTitle & "</a><a><cite>查看课程评价</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title""><legend><a name="""">评价人信息</a></legend></fieldset>" & vbCrlf
	Response.Write "	<div class=""evalTable"">" & vbCrlf
	Response.Write "		<table class=""layui-table"">" & vbCrlf
	Response.Write "			<tbody><tr><td class=""t-title"">评价人人：</td><td>" & tYGXM & "</td><td class=""t-title"">工号：</td><td colspan=""3"">" & tYGDM & "</td></tr>" & vbCrlf
	Response.Write "				<tr><td class=""t-title"">科室：</td><td>" & tKSMC & "</td><td class=""t-title"">职称：</td><td>" & tPRZC & "</td><td class=""t-title"">职务：</td><td>" & tXZZW & "</td></tr>" & vbCrlf
	Response.Write "			</tbody>" & vbCrlf
	Response.Write "		</table>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title""><legend><a name="""">所有评价</a></legend></fieldset>" & vbCrlf
	Response.Write "	<form name=""SearchForm"" id=""SearchForm"" class=""layui-form soBox"" action="""" method=""get""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input type=""text"" class=""layui-input"" name=""soWord"" value=""" & soWord & """ id=""soWord"" placeholder=""搜索授课人"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-normal"" data-event=""reload""><i class=""hr-icon hr-icon-top"">&#xeba1;</i>搜索</button></div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""toolBtn"">" & vbCrlf	'表头模板
	Response.Write "		<div class=""hr-rows tpltools"">" & vbCrlf
	Response.Write "			<div class=""layui-btn-group hr-tablebtn"">" & vbCrlf
	Response.Write "				<button type=""button"" class=""layui-btn hr-btn_deon"" lay-event=""addNew"" title=""新增评价""><i class=""hr-icon"">&#xecfb;</i></button>" & vbCrlf
	Response.Write "				<button type=""button"" class=""layui-btn hr-btn_fuch"" lay-event=""batchDel"" title=""批量删除""><i class=""hr-icon"">&#xea64;</i></button>" & vbCrlf
	Response.Write "				<button type=""button"" class=""layui-btn hr-btn_skyblue"" lay-event=""reload"" title=""重载数据""><i class=""hr-icon"">&#xf01e;</i></button>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""sumbar"">总计：<b class=""Count"">0</b>次</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf		'行工具
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm hr-btn_deon"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""details"" title=""查看评价表""><i class=""hr-icon"">&#xea59;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element"", ""form""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table; element = layui.element, form=layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",height:'full-115',page:true,limit:30,skin:'',limits:[10,15,20,30,50,100,200,300],toolbar: '#toolBtn'" & vbCrlf
	tmpHtml = tmpHtml & "			,text:{none:'暂时没有数据'},totalRow:true,cols: [[" & vbCrlf				'设置表头
	tmpHtml = tmpHtml & "				{type:'checkbox',unresize:true,align:'center',width:50}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Teacher',title:'授课人',width:95}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherCode',title:'工号',width:80}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherDepart',title:'科室',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Course',title:'课程名称',width:160}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Student',title:'授课对象',minWidth:100,event:'details',style:'cursor: pointer;'}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'CourseDate',title:'授课时间',width:125}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'AppraiTime',title:'评价时间',width:125}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TotalScore',title:'评分',sort:true,align:'center',width:95}" & vbCrlf
	tmpHtml = tmpHtml & "				,{title:'操作',align:'center',unresize:true,width:100, toolbar:'#barTable'}" & vbCrlf
	tmpHtml = tmpHtml & "			]],url:""" & layUrl & """,parseData: function(res){" & vbCrlf
	tmpHtml = tmpHtml & "				$("".Count"").text(res.count);" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf		'监听工具栏
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""edit""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""viewWin"", content:""" & ParmPath & "Evaluate/Edit.html?ID="" + data.ID,title:[""编辑评价"",""font-size:16""],area:[""760px"", ""92%""],moveOut:true });" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""details""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""viewWin"", content:""" & ParmPath & "Evaluate/Details.html?ID="" + data.ID,title:[""查看评价详情"",""font-size:16""],area:[""760px"", ""92%""],moveOut:true });" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".searchBtn button"").on(""click"", function(){" & vbCrlf	'搜索等按钮click事件
	tmpHtml = tmpHtml & "			var btnEvent = $(this).data(""event"");" & vbCrlf
	tmpHtml = tmpHtml & "			switch(btnEvent){" & vbCrlf
	tmpHtml = tmpHtml & "				case ""reload"":" & vbCrlf
	tmpHtml = tmpHtml & "					var arrForm = $(""#SearchForm"").serializeArray(), postStr={};" & vbCrlf
	tmpHtml = tmpHtml & "					$.each(arrForm, function(key, val){ postStr[this.name]=this.value; });" & vbCrlf		'表单序列转json
	tmpHtml = tmpHtml & "					table.reload(""layList"", {" & vbCrlf
	tmpHtml = tmpHtml & "						url:""" & layUrl & """,where: postStr" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""refresh"":location.reload(); break;" & vbCrlf
	tmpHtml = tmpHtml & "			};" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""toolbar(TableList)"", function(obj){" & vbCrlf		'监听表头工具
	tmpHtml = tmpHtml & "			var data = table.checkStatus(obj.config.id).data;" & vbCrlf
	tmpHtml = tmpHtml & "			switch(obj.event){" & vbCrlf
	tmpHtml = tmpHtml & "				case ""addNew"":" & vbCrlf
	tmpHtml = tmpHtml & "					layer.open({type:2, content:'" & ParmPath & "Evaluate/AddNew.html?ygdm=" & tYGDM & "',title:[""添加课堂教学评价"",""font-size:16""],area:[""850px"",""90%""],moveOut:true,maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""batchDel"":" & vbCrlf
	tmpHtml = tmpHtml & "					if(data.length==0){layer.tips(""请选择您要删除的评价！"","".laytable-cell-checkbox"",{tips: [1, ""#F60""]});return false;}" & vbCrlf
	tmpHtml = tmpHtml & "					var arrID=[]; for(var i=0;i<data.length;i++){ arrID.push(data[i].ID); }" & vbCrlf
	tmpHtml = tmpHtml & "					layer.confirm(""确认要删除选中的 "" + data.length + "" 条数据？<br />删除后将无法恢复。"",{icon:3, title:[""删除警告"",""background-color:#f30""]},function(index){" & vbCrlf
	tmpHtml = tmpHtml & "						$.getJSON(""" & ParmPath & "Evaluate/Delete.html"",{ID:arrID.join()}, function(reJson){" & vbCrlf
	tmpHtml = tmpHtml & "							layer.msg(reJson.msg,{title:""删除结果"",btn:""关闭"",time:0},function(){ table.reload(""layList""); });" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""refresh"":table.reload(""layList"");break;" & vbCrlf
	tmpHtml = tmpHtml & "			};" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub ListData()				'
	Dim tmpJson, rsGet, sqlGet, tIntro, tSum
	Dim tLimit : tLimit = HR_Clng(Request("limit"))
	Dim tPage : tPage = HR_Clng(Request("page"))
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tYGDM : tYGDM = HR_Clng(Trim(Request("ygdm")))

	sqlGet = "Select a.*,b.KSMC,b.PRZC"
	sqlGet = sqlGet & " ,(Select KSMC From HR_Teacher Where YGDM=a.TeacherID) As TeacherKS"
	sqlGet = sqlGet & " ,(Select PRZC From HR_Teacher Where YGDM=a.TeacherID) As TeacherZC"
	sqlGet = sqlGet & " From HR_Evaluate a Left Join HR_Teacher b on a.ParticipantID=b.YGDM Where a.ParticipantID=" & tYGDM
	If HR_IsNull(soWord) = False Then sqlGet = sqlGet & " And a.Teacher like '%" & soWord &"%'"
	sqlGet = sqlGet & " And a.ClassTime<'" & DefYear & "-06-30 23:59:59' And a.ClassTime>'" & DefYear - 1 & "-07-01 00:00:00'"
	sqlGet = sqlGet & " Order By a.ParticipantID ASC, CreateTime DESC"
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			i = 0 : CurrentPage = 1 : MaxPerPage = tLimit
			If tPage > 0 Then CurrentPage = tPage
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
			Dim tSheetName, tTemplate, tItemName, tCourse, tStudent, tCourseDate, tmpDate
			Do While Not rsGet.EOF
				tItemName = GetTypeName("HR_Class", "ClassName", "ClassID", HR_CLng(rsGet("ItemID")))
				tSheetName = "HR_Sheet_" & rsGet("ItemID")		'数据表名
				tTemplate = GetTypeName("HR_Class", "Template", "ClassID", HR_CLng(rsGet("ItemID")))
				If i > 0 Then tmpJson = tmpJson & ","
				tmpJson = tmpJson & "{""ID"":" & HR_CLng(rsGet("ID")) & ",""Teacher"":""" & Trim(rsGet("Teacher")) & """,""TeacherCode"":""" & HR_CLng(rsGet("TeacherID")) & """,""TeacherDepart"":""" & Trim(rsGet("TeacherKS")) & """,""TeacherZC"":""" & Trim(rsGet("TeacherZC")) & """"
				tmpJson = tmpJson & ",""ItemID"":""" & rsGet("ItemID") & """,""ItemName"":""" & tItemName & """,""CourseID"":""" & HR_CLng(rsGet("CourseID")) & """,""Course"":""" & Trim(rsGet("Course")) & """,""Student"":""" & Trim(rsGet("StuClass")) & """"
				tmpJson = tmpJson & ",""Appraiser"":""" & Trim(rsGet("Participant")) & """,""AppraiserCode"":""" & Trim(rsGet("ParticipantID")) & """,""AppraiserKS"":""" & Trim(rsGet("KSMC")) & """,""AppraiserZC"":""" & Trim(rsGet("PRZC")) & """,""Campus"":""" & FilterHtmlToText(rsGet("Campus")) & """"
				tmpJson = tmpJson & ",""Contents"":""" & Trim(rsGet("Contents")) & """,""AppraiTime"":""" & FormatDate(rsGet("CreateTime"), 10) & """,""Merit"":""" & FilterHtmlToText(rsGet("Merit")) & """"
				tmpJson = tmpJson & ",""CourseDate"":""" & Trim(rsGet("ClassTime")) & """,""TotalScore"":" & HR_CLng(rsGet("TotalScore")) & ",""CountEvalu"":" & HR_CLng(rsGet("TeacherKS")) & "}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""查询成功"",""Sum"":" & HR_Clng(tSum) & ",""count"":" & HR_Clng(TotalPut) & ",""data"":[" & tmpJson & "],""limit"":""" & MaxPerPage & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub
Sub Details()
	Dim tmpID : tmpID = HR_Clng(Trim(Request("ID")))
	Dim sqlGet, rsGet
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.t-title {}" & vbCrlf
	tmpHtml = tmpHtml & "		.w1 {width:100px;text-align:right;background-color:#eee}" & vbCrlf
	tmpHtml = tmpHtml & "		.w2 {width:130px;text-align:left;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title""><legend><a name="""">评价详情</a></legend></fieldset>" & vbCrlf
	Response.Write "	<div class=""evalTable"">" & vbCrlf
	Response.Write "		<table class=""layui-table"">" & vbCrlf

	sqlGet = "Select a.*"
	sqlGet = sqlGet & " From HR_Evaluate a Left Join HR_Teacher b on a.ParticipantID=b.YGDM Where a.ID=" & tmpID
	Set rsGet = Conn.Execute(sqlGet)
		If Not(rsGet.BOF And rsGet.EOF) Then
			Response.Write "			<tbody><tr><td class=""t-title"" colspan=""4"">授课信息</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title w1"">授课人：</td><td>" & Trim(rsGet("Teacher")) & "</td><td class=""t-title w1"">授课时间：</td><td class=""w2"">" & FormatDate(rsGet("ClassTime"), 4) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title w1"">授课对象：</td><td>" & Trim(rsGet("StuClass")) & "</td><td class=""t-title w1"">课程名称：</td><td class=""w2"">" & Trim(rsGet("Course")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title w1"">授课内容：</td><td>" & Trim(rsGet("Contents")) & "</td><td class=""t-title w1"">开课学院：</td><td class=""w2"">" & Trim(rsGet("Campus")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""4"">一、教学态度与基本技能</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""3"">1、要求脱稿讲授，语言准确流畅，逻辑性强，富感染力，语速、语调适宜、抑扬顿挫。</td><td class=""w2"">" & HR_CLng(rsGet("Score1")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""3"">2、精神饱满，教态大方，仪表端正。</td><td class=""w2"">" & HR_CLng(rsGet("Score2")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""3"">3、PPT设计科学，板书工整，教案讲稿规范。</td><td class=""w2"">" & HR_CLng(rsGet("Score3")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""4"">二、教学设计与方法</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""3"">1、运用先进教学理念、方法进行教学，三维目标明确，学情清楚，因材施教，循循善诱。</td><td class=""w2"">" & HR_CLng(rsGet("Score4")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""3"">2、教学设计科学，新课导入、知识教授、总结巩固、课外自主学习等教学环节设计合理。</td><td class=""w2"">" & HR_CLng(rsGet("Score5")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""3"">3、广泛使用多媒体、互联网等现代化教学手段进行辅助教学。</td><td class=""w2"">" & HR_CLng(rsGet("Score6")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""4"">三、教学内容</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""3"">符合教学大纲（或课程标准）要求，授课内容正确，重点难点突出，深度与广度适宜，联系实际，例证恰当，适当关注学科进展。</td><td class=""w2"">" & HR_CLng(rsGet("Score7")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""4"">四、教学效果</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""3"">1、课堂驾驭能力强，师生互动性、课堂纪律、学习气氛好。</td><td class=""w2"">" & HR_CLng(rsGet("Score8")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""3"">2、完成教学任务，实现教学目的，学生反馈教学效果好。</td><td class=""w2"">" & HR_CLng(rsGet("Score9")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""4"">五、整体评价</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""3"">整体评价</td><td class=""w2"">" & HR_CLng(rsGet("Score10")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""4"">六、优点</td></tr>" & vbCrlf
			Response.Write "				<tr><td colspan=""4"">" & Trim(rsGet("Merit")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"" colspan=""4"">七、问题与建议</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title w1"">1、意识形态</td><td colspan=""3"">" & Trim(rsGet("Suggest0")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title w1"">2、教学</td><td colspan=""3"">" & Trim(rsGet("Suggest1")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title w1"">3、学风</td><td colspan=""3"">" & Trim(rsGet("Suggest2")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title w1"">4、硬件</td><td colspan=""3"">" & Trim(rsGet("Suggest3")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title w1"">评价人：</td><td>" & Trim(rsGet("Participant")) & "</td><td class=""t-title w1"">评价时间：</td><td>" & FormatDate(rsGet("CreateTime"), 4) & "</td></tr>" & vbCrlf
			Response.Write "			</tbody>" & vbCrlf
		End If
	Set rsGet = Nothing
	Response.Write "		</table>" & vbCrlf
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

Sub EditBody()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tYGDM : tYGDM = Trim(Request("ygdm"))
	Dim tYGXM : tYGXM = strGetTypeName("HR_Teacher", "ygxm", "ygdm", tYGDM)
	Dim isModify : isModify = HR_CBool(Request("Modify"))
	Dim tTeacher, tTeacherID, tParticipant, tParticipantID
	Dim tItemID, tItemName, tCourseID, tSheetName, arrScore(10), tTotalScore, tMerit, arrSuggest(3)
	Dim tCourse, tTitle, tCampus, tStuClass, tContents, tClassTime, tCreateTime, tPassed
	Set rs = Conn.Execute("Select * From HR_Evaluate Where ID=" & tmpID)
		If Not(rs.BOF And rs.EOF) Then
			isModify = True
			tParticipant = Trim(rs("Participant"))
			tParticipantID = HR_CLng(rs("ParticipantID"))
			tTeacher = Trim(rs("Teacher"))
			tTeacherID = HR_CLng(rs("TeacherID"))

			tItemID = HR_CLng(rs("ItemID"))
			tCourseID = HR_CLng(rs("CourseID"))
			tCourse = Trim(rs("Course"))
			tClassTime = FormatDate(rs("ClassTime"), 2)
			tTitle = Trim(rs("Title"))
			tCampus = Trim(rs("Campus"))
			tStuClass = Trim(rs("StuClass"))
			tContents = Trim(rs("Contents"))
			tTotalScore = HR_CDbl(rs("TotalScore"))
			For i = 0 To 9
				arrScore(i) = HR_CLng(rs("Score" & i + 1))
			Next
			For i = 0 To 3
				arrSuggest(i) = Trim(rs("Suggest" & i))
			Next
			tMerit = Trim(rs("Merit"))
			tCreateTime = FormatDate(rs("CreateTime"), 2)
			tPassed = HR_CBool(rs("Passed"))
		Else
			tParticipantID = tYGDM
			tParticipant = tYGXM
		End If
	Set rs = Nothing
	tSheetName = "HR_Sheet_" & tItemID
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .morebtn {padding:3px 0!important;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .tips {padding-left:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-slider {flex-grow:1;} .slider{box-sizing:border-box;padding:1px 5px 0 5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.form-title {border-top:1px solid #eee;border-bottom:1px solid #eee;margin-bottom:10px;} .form-title h3 {padding:0 0 0 10px;line-height:2}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .inputleft {width:75%;text-align:left;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form class=""layui-form"" id=""EditForm"" name=""EditForm"" lay-filter=""EditForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layui-form-item""><label class=""layui-form-label"">选择评价人：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""text"" name=""Participant"" id=""ygxm"" value=""" & tParticipant & """ lay-verify=""required"" autocomplete=""on"" title=""查找评价人"" class=""layui-input txt1"" readonly></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux morebtn""><span class=""layui-btn layui-btn-sm getBtn"" data-code=""ygdm"" data-name=""ygxm"">查找</span><span class=""tips"">请输入关键字搜索教师，必填项</span></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">工　　号：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""text"" name=""ParticipantID"" id=""ygdm"" lay-verify=""required"" value=""" & tParticipantID & """ class=""layui-input txt1"" readonly></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf

	Response.Write "		<div class=""layui-form-item""><label class=""layui-form-label"">选择授课人：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""text"" name=""Teacher"" id=""teacher"" value=""" & tTeacher & """ lay-verify=""required"" autocomplete=""on"" title=""查找授课人"" class=""layui-input txt1"" readonly></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux morebtn""><span class=""layui-btn layui-btn-sm getBtn"" data-code=""teacherid"" data-name=""teacher"">查找</span><span class=""tips"">请输入关键字搜索教师，必填项</span></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">工　　号：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""text"" name=""TeacherID"" id=""teacherid"" lay-verify=""required"" value=""" & tTeacherID & """ class=""layui-input txt1"" readonly></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">选择项目：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><select name=""ItemID"" id=""ItemID"" lay-filter=""ItemID""><option value="""">请选择项目</option>" & GetItemOption(1, tItemID, True) & "</select></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">选择课程：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline"" id=""CourseSelect""><select name=""CourseID"" id=""CourseID"" lay-filter=""CourseOption""><option value="""">请选择课程</option>" & GetItemCourseOption(tItemID, tCourseID, tTeacherID, "") & "</select></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">授课时间：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""ClassTime"" id=""ClassTime"" lay-verify=""date"" value=""" & tClassTime & """ class=""layui-input dataitem"" readonly></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	'Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">开课学院：</label>" & vbCrlf
	'Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""Campus"" id=""Campus"" lay-verify=""required"" value=""" & tCampus & """ class=""layui-input""></div>" & vbCrlf
	'Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">课程名称：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""Course"" id=""Course"" lay-verify=""required"" value=""" & tCourse & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">授课内容：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""Contents"" id=""Contents"" lay-verify=""required"" value=""" & tContents & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">授课对象：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:45%""><input type=""text"" name=""StuClass"" id=""StuClass"" lay-verify=""required"" value=""" & tStuClass & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""form-title""><h3>一、教学态度与基本技能</h3></div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">1、要求脱稿讲授，语言准确流畅，逻辑性强，富感染力，语速、语调适宜、抑扬顿挫。（10分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""Score1"" id=""Score1"" lay-verify=""number"" min=""0"" max=""10"" value=""" & arrScore(0) & """ class=""layui-input count""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">2、精神饱满，教态大方，仪表端正。（10分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""Score2"" id=""Score2"" lay-verify=""number"" min=""0"" max=""10"" value=""" & arrScore(1) & """ class=""layui-input count""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">3、PPT设计科学，板书工整，教案讲稿规范。（10分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""Score3"" id=""Score3"" lay-verify=""number"" min=""0"" max=""10"" value=""" & arrScore(2) & """ class=""layui-input count""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""form-title""><h3>二、教学设计与方法</h3></div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">1、运用先进教学理念、方法进行教学，三维目标明确，学情清楚，因材施教，循循善诱。（10分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""Score4"" id=""Score4"" lay-verify=""number"" min=""0"" max=""10"" value=""" & arrScore(3) & """ class=""layui-input count""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">2、教学设计科学，新课导入、知识教授、总结巩固、课外自主学习等教学环节设计合理。（10分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""Score5"" id=""Score5"" lay-verify=""number"" min=""0"" max=""10"" value=""" & arrScore(4) & """ class=""layui-input count""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">3、广泛使用多媒体、互联网等现代化教学手段进行辅助教学。（10分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""Score6"" id=""Score6"" lay-verify=""number"" min=""0"" max=""10"" value=""" & arrScore(5) & """ class=""layui-input count""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""form-title""><h3>三、教学内容</h3></div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">符合教学大纲（或课程标准）要求，授课内容正确，重点难点突出，深度与广度适宜，联系实际，例证恰当，适当关注学科进展。（10分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""Score7"" id=""Score7"" lay-verify=""number"" min=""0"" max=""10"" value=""" & arrScore(6) & """ class=""layui-input count""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""form-title""><h3>四、教学效果</h3></div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">1、课堂驾驭能力强，师生互动性、课堂纪律、学习气氛好。（10分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""Score8"" id=""Score8"" lay-verify=""number"" min=""0"" max=""10"" value=""" & arrScore(7) & """ class=""layui-input count""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">1、完成教学任务，实现教学目的，学生反馈教学效果好。（10分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""Score9"" id=""Score9"" lay-verify=""number"" min=""0"" max=""10"" value=""" & arrScore(8) & """ class=""layui-input count""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""form-title""><h3>五、整体评价</h3></div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">整体评价（10分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""Score10"" id=""Score10"" lay-verify=""number"" min=""0"" max=""10"" value=""" & arrScore(9) & """ class=""layui-input count""></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label inputleft"">总评得分（100分）</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""number"" name=""TotalScore"" id=""TotalScore"" min=""0"" max=""100"" value=""" & tTotalScore & """ class=""layui-input"" readonly></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux"">分</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""form-title""><h3>六、优点</h3></div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item layui-form-text""><label class=""layui-form-label"">优点：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><textarea name=""Merit"" placeholder=""请输入内容"" class=""layui-textarea"">" & tMerit & "</textarea></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""form-title""><h3>七、问题与建议</h3></div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item layui-form-text""><label class=""layui-form-label"">1、意识形态：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><textarea name=""Suggest0"" placeholder=""请输入内容"" class=""layui-textarea"">" & arrSuggest(0) & "</textarea></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item layui-form-text""><label class=""layui-form-label"">2、教学：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><textarea name=""Suggest1"" placeholder=""请输入内容"" class=""layui-textarea"">" & arrSuggest(1) & "</textarea></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item layui-form-text""><label class=""layui-form-label"">3、学风：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><textarea name=""Suggest2"" placeholder=""请输入内容"" class=""layui-textarea"">" & arrSuggest(2) & "</textarea></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item layui-form-text""><label class=""layui-form-label"">4、硬件：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><textarea name=""Suggest3"" placeholder=""请输入内容"" class=""layui-textarea"">" & arrSuggest(3) & "</textarea></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf

	If isModify Then Response.Write "	<input name=""Modify"" type=""hidden"" value=""True""><input name=""ID"" type=""hidden"" value=""" & tmpID & """>" & vbCrlf
	Response.Write "		<div class=""hr-pop-fix"">" & vbCrlf
	Response.Write "			<div class=""hr-grids hr-btn-group"">" & vbCrlf
	Response.Write "				<em><button type=""button"" class=""layui-btn hr-btn_deon"" id=""EditPost"" data-event=""EditPost"" lay-submit title=""保存""><i class=""hr-icon"">&#xf0c7;</i>保存</button></em>" & vbCrlf
	Response.Write "				<em><button type=""button"" class=""layui-btn layui-btn-primary"" id=""refresh"" data-event=""refresh"" title=""刷新""><i class=""hr-icon"">&#xf343;</i></button></em>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-place-h50""></div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""form"", ""table"", ""element"", ""laydate"", ""slider""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table=layui.table, element=layui.element, form=layui.form, laydate=layui.laydate, slider=layui.slider;" & vbCrlf

	tmpHtml = tmpHtml & "		lay("".dataitem"").each(function(){" & vbCrlf			'更改日期时取星期
	tmpHtml = tmpHtml & "			laydate.render({elem: this, format: 'yyyy-MM-dd',done:function(value, date, endDate){" & vbCrlf
	'tmpHtml = tmpHtml & "				var today = new Array('日','一','二','三','四','五','六'), day = new Date(value);" & vbCrlf
	'tmpHtml = tmpHtml & "				var week = today[day.getDay()];$(""#VA6"").val(week);" & vbCrlf
	tmpHtml = tmpHtml & "				}" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		$("".getBtn"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var elcode=$(this).data(""code""), elname=$(this).data(""name"");" & vbCrlf		'返回员工代码及名称时的对象
	tmpHtml = tmpHtml & "			var openurl=""" & InstallDir & "Desktop/Contacts/Float.html?Type=3"";" & vbCrlf
	tmpHtml = tmpHtml & "			parent.layer.open({type:2,id:""getWin"",content:openurl, title:[""选择教师"",""font-size:16""],area:[""500px"", ""80%""],scrollbar:false,success:function(layero, index){" & vbCrlf
	tmpHtml = tmpHtml & "				var objIframe = $(layero).find('iframe')[0].contentWindow.document.body;" & vbCrlf
	tmpHtml = tmpHtml & "				var obj1 = $(objIframe).contents().find(""#listGroup"");" & vbCrlf
	tmpHtml = tmpHtml & "				obj1.attr(""value"",window.name);obj1.attr(""code"", elcode); obj1.attr(""name"", elname);" & vbCrlf		'回车搜索
	tmpHtml = tmpHtml & "				}" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		form.on(""select(ItemID)"", function(data){" & vbCrlf			'监听项目下拉，联动课程
	tmpHtml = tmpHtml & "			var ygdm = $(""#teacherid"").val(), course = $(""#CourseID"").val();" & vbCrlf
	tmpHtml = tmpHtml & "			if(!ygdm){layer.tips(""请选择申请人！"" + ygdm,""#teacher"");return false;}" & vbCrlf
	tmpHtml = tmpHtml & "			$.post(""" & ParmPath & "Evaluate/CourseOption.html"", {ItemID:data.value, ygdm:ygdm, CourseID:course}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#CourseSelect"").html(rsStr);form.render(""select"");" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		form.on(""select(CourseOption)"", function(data){" & vbCrlf		'监听课程下拉，更新课程数据
	tmpHtml = tmpHtml & "			var ygdm = $(""#teacherid"").val(), itemid = $(""#ItemID"").val();" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "Evaluate/GetCourseData.html"", {ItemID:itemid, ygdm:ygdm, CourseID:data.value}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#ClassTime"").val(rsStr.VA4);$(""#Course"").val(rsStr.VA8);$(""#Contents"").val(rsStr.VA9);" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#StuClass"").val(rsStr.VA10); $(""#Campus"").val(rsStr.VA11);" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".count"").bind(""input propertychange"", function(){" & vbCrlf					'计算总分
	tmpHtml = tmpHtml & "			var total = 0;" & vbCrlf
	tmpHtml = tmpHtml & "			$("".count"").each(function(){" & vbCrlf
	tmpHtml = tmpHtml & "				total = total + parseInt($(this).val());" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#TotalScore"").val(total);" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#EditPost"").on(""click"", function(){" & vbCrlf					'提交保存
	tmpHtml = tmpHtml & "			$.post(""" & ParmPath & "Evaluate/SaveEdit.html"",$(""#EditForm"").serialize(), function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.alert(reData.errmsg,{icon:reData.icon,btn:""关闭""},function(){" & vbCrlf
	tmpHtml = tmpHtml & "					if(!reData.err){" & vbCrlf
	tmpHtml = tmpHtml & "						var index1 = parent.layer.getFrameIndex(window.name);" & vbCrlf
	tmpHtml = tmpHtml & "						parent.layui.table.reload(""layList"");" & vbCrlf		'重构数据列表
	tmpHtml = tmpHtml & "						parent.layer.close(index1);" & vbCrlf		'关闭[在iframe页面]
	tmpHtml = tmpHtml & "					}else{layer.close(layer.index)}" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			},""json"");" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = Replace(getPageFoot(1), "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub SaveEdit()
	Dim tmpJson, rsSave, sqlSave
	Dim tmpID : tmpID = HR_Clng(Trim(Request("ID")))
	Dim isModify : isModify = HR_CBool(Request("Modify"))
	Dim tParticipant : tParticipant = Trim(Request("Participant"))
	Dim tParticipantID : tParticipantID = HR_Clng(Trim(Request("ParticipantID")))
	Dim tTeacher : tTeacher = Trim(Request("Teacher"))
	Dim tTeacherID : tTeacherID = HR_Clng(Trim(Request("TeacherID")))
	Dim tItemID : tItemID = HR_Clng(Trim(Request("ItemID")))
	Dim tCourseID : tCourseID = HR_Clng(Trim(Request("CourseID")))
	Dim tNow : tNow = FormatDate(Now(), 1)

	If tParticipantID = 0 Then ErrMsg = "请选择评价人！"
	If tTeacherID = 0 Then ErrMsg = "请选择授课老师！"
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""err"":true,""errcode"":500,""icon"":2,""errmsg"":""" & ErrMsg & """}" : Exit Sub

	sqlSave = "Select * From HR_Evaluate Where ID>0 And ID=" & tmpID
	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open(sqlSave), Conn, 1, 3
		If rsSave.BOF And rsSave.EOF Then
			rsSave.AddNew
			rsSave("ID") = GetNewID("HR_Evaluate", "ID")
			rsSave("CreateTime") = tNow
			rsSave("Title") = "课堂教学质量评价"
			rsSave("Passed") = 0
		End If

		rsSave("Participant") = tParticipant
		rsSave("ParticipantID") = tParticipantID
		rsSave("Teacher") = tTeacher
		rsSave("TeacherID") = tTeacherID
		rsSave("ItemID") = tItemID
		rsSave("CourseID") = tCourseID
		rsSave("ClassTime") = SaveDate(Request("ClassTime"))
		rsSave("Campus") = Trim(Request("Campus"))
		rsSave("Course") = Trim(Request("Course"))
		rsSave("Contents") = Trim(Request("Contents"))
		rsSave("StuClass") = Trim(Request("StuClass"))
		rsSave("Score1") = HR_Clng(Request("Score1"))
		rsSave("Score2") = HR_Clng(Request("Score2"))
		rsSave("Score3") = HR_Clng(Request("Score3"))
		rsSave("Score4") = HR_Clng(Request("Score4"))
		rsSave("Score5") = HR_Clng(Request("Score5"))
		rsSave("Score6") = HR_Clng(Request("Score6"))
		rsSave("Score7") = HR_Clng(Request("Score7"))
		rsSave("Score8") = HR_Clng(Request("Score8"))
		rsSave("Score9") = HR_Clng(Request("Score9"))
		rsSave("Score10") = HR_Clng(Request("Score10"))
		rsSave("TotalScore") = HR_Clng(Request("TotalScore"))
		rsSave("Merit") = Trim(Request("Merit"))
		rsSave("Suggest0") = Trim(Request("Suggest0"))
		rsSave("Suggest1") = Trim(Request("Suggest1"))
		rsSave("Suggest2") = Trim(Request("Suggest2"))
		rsSave("Suggest3") = Trim(Request("Suggest3"))
		rsSave.Update
	Set rsSave = Nothing
	Response.Write "{""err"":false,""errcode"":0,""icon"":1,""errmsg"":""评价保存成功！""}"

End Sub
Sub CourseOption()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tCourseID : tCourseID = HR_Clng(Request("CourseID"))
	Dim tYGDM : tYGDM = HR_Clng(Request("ygdm"))
	tmpHtml = GetItemCourseOption(tItemID, tCourseID, tYGDM, "")
	If HR_IsNull(tmpHtml) Then
		tmpHtml = "<select name=""CourseID"" id=""CourseID""><option value="""">暂无课程</option></select>"
	Else
		tmpHtml = "<select name=""CourseID"" id=""CourseID"" lay-filter=""CourseOption""><option value="""">请选择课程</option>" & tmpHtml & "</select>"
	End If
	Response.Write tmpHtml
End Sub
Sub GetCourseData()		'取指定课程数据
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tCourseID : tCourseID = HR_Clng(Request("CourseID"))
	Dim tYGDM : tYGDM = HR_Clng(Request("ygdm"))
	Dim tSheetName : tSheetName = "HR_Sheet_" & tItemID
	Dim tmpData
	If ChkTable(tSheetName) Then
		sql = "Select * From " & tSheetName & " Where scYear=" & DefYear & " And ID=" & tCourseID
		Set rs = Conn.Execute(sql)
			If Not(rs.BOF And rs.EOF) Then
				tmpData = "{""VA1"":" & HR_CLng(rs("VA1")) & ",""VA2"":""" & Trim(rs("VA2")) & """,""VA3"":" & HR_CDbl(rs("VA3")) & ",""VA4"":""" & FormatDate(ConvertNumDate(rs("VA4")), 2) & ""","
				tmpData = tmpData & """VA5"":""" & Trim(rs("VA5")) & """,""VA6"":""" & Trim(rs("VA6")) & """,""VA7"":""" & Trim(rs("VA7")) & """,""VA8"":""" & Trim(rs("VA8")) & """,""VA9"":""" & Trim(rs("VA9")) & ""","
				tmpData = tmpData & """VA10"":""" & Trim(rs("VA10")) & """,""VA11"":""" & Trim(rs("VA11")) & """,""VA12"":""" & Trim(rs("VA12")) & """}"
			End If
		Set rs = Nothing
	End If
	Response.Write tmpData
End Sub
%>