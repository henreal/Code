<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
SiteTitle = "mini-CEX plus记录管理"
Dim arrTeachJob : arrTeachJob = Array("主任医师","副主任医师","主治医师","完成住培的住院医师")
Dim arrStuMajor : arrStuMajor = Array("实习医师","住培/硕士研究生","住院医师","专培/博士研究生")
Dim arrEvalAdd : arrEvalAdd = Array("病房","门诊","急诊","ICU","临床技能中心","其他")
Dim arrPatientType : arrPatientType = Array("新接触患者","已接触患者")
Dim arrComplexity : arrComplexity = Array("低","中","高")
Dim arrFocus : arrFocus = Array("医疗问诊","体格检查","临床操作","医疗咨询及宣教","临床思维与治疗")
Dim arrEvaluate1 : arrEvaluate1 = Array("正确称呼患者","自我介绍","向患者说明目的","尽可能让患者自己陈述，适时给患者支持、鼓励","耐心倾听患者陈述","与患者有适当的眼神、言语、肢体的交流","适时引导患者，以充分获取正确资料","问诊逻辑清晰、条理清楚","采用易懂语言","重点突出，信息收集完整","必要的记录")
Dim arrEvaluate2 : arrEvaluate2 = Array("准备必需的体检用物","注意保护患者的隐私，必要时，请其他人员在旁","清洁双手","按病情需要进行检查，顺序合理，及时处理患者在体检中出现的不适","检查手法规范","检查内容全面")
Dim arrEvaluate3 : arrEvaluate3 = Array("了解适应证及相关解剖知识","取得患者同意（口头或书面）","操作前准备","适当的止痛或镇静","操作能力","无菌技术","适时寻求帮助","术后处理")
Dim arrEvaluate4 : arrEvaluate4 = Array("能对病史与体检内容进行整合、分析","能解释相关的检查结果","临床分析具有逻辑性","有一定的诊断、鉴别诊断能力","治疗方案合理可行")
Dim arrEvaluate5 : arrEvaluate5 = Array("解释检查或处置的基本理由","各种治疗方案的利弊比较","患者用药指导","生活方式及注意事项的宣教")
Dim arrEvaluate6 : arrEvaluate6 = Array("仪表端正，态度和蔼，口齿清楚","尊重患者与家属，具有同情心","获得患者与家属的信任","注意患者的舒适度，适时正确处理患者出现的不适","适当解释患者及家属提出的问题")
Dim arrEvaluate7 : arrEvaluate7 = Array("对患者及家属态度","时间控制得当，过程简洁精炼","有整合资料与判断能力","能按优先顺序进行正确处理","整体效率高")

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
	Case "AllData" Call AllDataList()
	Case "jsonAllData" Call jsonAllData()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim layUrl : layUrl = ParmPath & "EvaluateCEX/jsonList.html"

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-tablebtn .layui-btn {height:28px;line-height:28px;padding:0 10px;font-size:1.1rem;}" & vbCrlf		'表头工具集
	tmpHtml = tmpHtml & "		.sumbar b {color:#F60;padding:0 2px} .sumbar b.sumDebit{color:#080}" & vbCrlf		'表头汇总
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)

	tmpHtml = "<a href=""" & ParmPath & "EvaluateCEX/Index.html"">" & SiteTitle & "</a><a><cite>列表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form name=""SearchForm"" id=""SearchForm"" class=""layui-form soBox"" action="""" method=""get""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input type=""text"" class=""layui-input"" name=""soWord"" value=""" & soWord & """ id=""soWord"" placeholder=""搜索测评人"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-normal"" data-event=""reload""><i class=""hr-icon hr-icon-top"">&#xeba1;</i>搜索</button></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn"">" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn hr-btn_deon"" data-event=""addnew"" name=""addnew"" title=""添加评价""><i class=""hr-icon"">&#xecfb;</i></button>" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn hr-btn_darkgreen"" data-event=""refresh"" name=""refresh"" title=""刷新本页""><i class=""hr-icon"">&#xf021;</i></button>" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn hr-btn_fuch"" data-event=""alldata"" name=""alldata"" title=""全表数据"">全表数据</button>" & vbCrlf
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
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""eval"" title=""查看所有评价""><i class=""hr-icon"">&#xefe2;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element"", ""form""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table; element = layui.element, form=layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",height:'full-115',page:true,limit:30,skin:'',limits:[10,15,20,30,50,100,200,300],toolbar: '#toolBtn'" & vbCrlf
	tmpHtml = tmpHtml & "			,text:{none:'暂时没有数据'},cols: [[" & vbCrlf				'设置表头
	tmpHtml = tmpHtml & "				{type:'checkbox',unresize:true,align:'center',width:50}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Teacher',title:'测评教师',width:95,event:'details',style:'cursor: pointer;'}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherCode',title:'工号',width:80}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherDepart',title:'科室',minWidth:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherZC',title:'职称',width:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherZW',title:'职务',width:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'CountEvalu',title:'次数',sort:true,align:'right',width:95,templet:function(res){ return res.CountEvalu + '次' } }" & vbCrlf
	'tmpHtml = tmpHtml & "				,{field:'EduYear',title:'学年',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{title:'操作',align:'center',unresize:true,width:100, toolbar:'#barTable'}" & vbCrlf
	tmpHtml = tmpHtml & "			]],url:""" & layUrl & """,parseData: function(res){" & vbCrlf
	tmpHtml = tmpHtml & "				$("".Count"").text(res.count);" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf		'监听工具栏
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""eval""){" & vbCrlf
	tmpHtml = tmpHtml & "				location.href=""" & ParmPath & "EvaluateCEX/List.html?ygdm="" + data.TeacherCode;" & vbCrlf
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
	tmpHtml = tmpHtml & "					layer.open({type:2, content:'" & ParmPath & "EvaluateCEX/AddNew.html',title:[""添加mini-CEX <sup>plus</sup>记录"",""font-size:16""],area:[""850px"",""90%""],moveOut:true,maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""addnew"":" & vbCrlf
	tmpHtml = tmpHtml & "				case ""alldata"":location.href=""" & ParmPath & "EvaluateCEX/AllData.html""; break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""refresh"":location.reload(); break;" & vbCrlf
	tmpHtml = tmpHtml & "			};" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub GetJsonList()
	Dim tmpJson, rsGet, sqlGet
	Dim tLimit : tLimit = HR_Clng(Request("limit"))
	Dim tPage : tPage = HR_Clng(Request("page"))
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))

	sqlGet = "Select Count(0) As Count, a.Teacher, a.TeacherID"
	sqlGet = sqlGet & " From HR_EvaluateCEX a"
	sqlGet = sqlGet & " Group By a.Teacher, a.TeacherID"

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
			Dim tKSMC, tPRZC, tXZZW
			Do While Not rsGet.EOF
				tKSMC = strGetTypeName("HR_Teacher", "KSMC", "YGDM", HR_CLng(rsGet("TeacherID")))
				If i > 0 Then tmpJson = tmpJson & ","
				tmpJson = tmpJson & "{""Teacher"":""" & Trim(rsGet("Teacher")) & """,""TeacherCode"":""" & HR_CLng(rsGet("TeacherID")) & """,""TeacherDepart"":""" & Trim(tKSMC) & """,""TeacherZC"":""" & Trim(rsGet("TeacherID")) & """"
				tmpJson = tmpJson & ",""TeacherZW"":""" & Trim(rsGet("TeacherID")) & """,""EduYear"":""" & DefYear-1 & "-" & DefYear & """,""CountEvalu"":" & HR_CLng(rsGet("Count")) & "}"
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
	Dim layUrl : layUrl = ParmPath & "EvaluateCEX/ListData.html?ygdm=" & tYGDM

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

	tmpHtml = "<a href=""" & ParmPath & "EvaluateCEX/Index.html"">" & SiteTitle & "</a><a><cite>列表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title""><legend><a name="""">评价人信息</a></legend></fieldset>" & vbCrlf
	Response.Write "	<div class=""evalTable"">" & vbCrlf
	Response.Write "		<table class=""layui-table"">" & vbCrlf
	Response.Write "			<tbody><tr><td class=""t-title"">评价人：</td><td>" & tYGXM & "</td><td class=""t-title"">工号：</td><td colspan=""3"">" & tYGDM & "</td></tr>" & vbCrlf
	Response.Write "				<tr><td class=""t-title"">科室：</td><td>" & tKSMC & "</td><td class=""t-title"">职称：</td><td>" & tPRZC & "</td><td class=""t-title"">职务：</td><td>" & tXZZW & "</td></tr>" & vbCrlf
	Response.Write "			</tbody>" & vbCrlf
	Response.Write "		</table>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title""><legend><a name="""">所有评价</a></legend></fieldset>" & vbCrlf
	Response.Write "	<form name=""SearchForm"" id=""SearchForm"" class=""layui-form soBox"" action="""" method=""get""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input type=""text"" class=""layui-input"" name=""soWord"" value=""" & soWord & """ id=""soWord"" placeholder=""搜索学生"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-normal"" data-event=""reload""><i class=""hr-icon hr-icon-top"">&#xeba1;</i>搜索</button></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn"">" & vbCrlf
	'Response.Write "			<button type=""button"" class=""layui-btn hr-btn_peru"" data-event=""update"" name=""update"" title=""更新数据""><i class=""layui-icon layui-anim layui-anim-rotate layui-anim-loop"">&#xe63d;</i></button>" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn hr-btn_darkgreen"" data-event=""refresh"" name=""refresh"" title=""刷新本页""><i class=""hr-icon"">&#xf021;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""toolBtn"">" & vbCrlf	'表头模板
	Response.Write "		<div class=""hr-rows tpltools"">" & vbCrlf
	'Response.Write "			<div class=""layui-btn-group hr-tablebtn"">" & vbCrlf
	'Response.Write "				<button type=""button"" class=""layui-btn hr-btn_fuch"" lay-event=""batchDel"" title=""批量删除""><i class=""hr-icon"">&#xea64;</i></button>" & vbCrlf
	'Response.Write "				<button type=""button"" class=""layui-btn hr-btn_skyblue"" lay-event=""reload"" title=""重载数据""><i class=""hr-icon"">&#xf01e;</i></button>" & vbCrlf
	'Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""sumbar"">共<b class=""Count"">0</b>条记录</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf		'行工具
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""details"" title=""查看评价表""><i class=""hr-icon"">&#xea59;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element"", ""form""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table; element = layui.element, form=layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",height:'full-115',page:true,limit:30,skin:'',limits:[10,15,20,30,50,100,200,300],toolbar: '#toolBtn'" & vbCrlf
	tmpHtml = tmpHtml & "			,text:{none:'暂时没有数据'},cols: [[" & vbCrlf				'设置表头
	tmpHtml = tmpHtml & "				{field:'ID',title:'序号',align:'center', width:60}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Teacher',title:'测评教师',width:95,event:'details',style:'cursor: pointer;'}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherDepart',title:'科室',width:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Student',title:'学生姓名',align:'center',width:95}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Major',title:'学生专业',width:125}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'SutType',title:'类　别',width:135}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Major',title:'学生专业',width:125}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'PatientType',title:'病人类别',width:125}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'EvaluateTime',title:'测评时间',width:125}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'EvaluateAdd',title:'测评地点',minWidth:95}" & vbCrlf
	tmpHtml = tmpHtml & "				,{title:'操作',align:'center',unresize:true,width:100, toolbar:'#barTable'}" & vbCrlf
	tmpHtml = tmpHtml & "			]],url:""" & layUrl & """,parseData: function(res){" & vbCrlf
	tmpHtml = tmpHtml & "				$("".Count"").text(res.count);" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf		'监听工具栏
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""details""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""viewWin"", content:""" & ParmPath & "EvaluateCEX/Details.html?ID="" + data.ID,title:[""查看评价详情"",""font-size:16""],area:[""760px"", ""92%""],moveOut:true });" & vbCrlf
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
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub ListData()				'
	Dim tmpJson, rsGet, sqlGet, tIntro
	Dim tLimit : tLimit = HR_Clng(Request("limit"))
	Dim tPage : tPage = HR_Clng(Request("page"))
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tYGDM : tYGDM = HR_Clng(Trim(Request("ygdm")))

	sqlGet = "Select a.*,(Select KSMC From HR_Teacher Where YGDM=a.TeacherID) As KSMC"
	sqlGet = sqlGet & " From HR_EvaluateCEX a Where a.TeacherID=" & tYGDM
	If HR_IsNull(soWord) = False Then sqlGet = sqlGet & " And a.Student like '%" & soWord &"%'"
	sqlGet = sqlGet & " Order By a.CreateTime DESC"
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
			Do While Not rsGet.EOF
				If i > 0 Then tmpJson = tmpJson & ","
				tmpJson = tmpJson & "{""ID"":" & HR_CLng(rsGet("ID")) & ",""TeacherCode"":" & HR_CLng(rsGet("TeacherID")) & ",""Teacher"":""" & Trim(rsGet("Teacher")) & """,""TeacherDepart"":""" & Trim(rsGet("KSMC")) & """,""TeacherJob"":""" & Trim(rsGet("TeacherJob")) & """"
				tmpJson = tmpJson & ",""Student"":""" & Trim(rsGet("Student")) & """,""Major"":""" & Trim(rsGet("Major")) & """,""SutType"":""" & Trim(rsGet("SutType")) & """,""EvaluateTime"":""" & FormatDate(rsGet("EvaluateTime"), 2) & """,""EvaluateAdd"":""" & FilterHtmlToText(rsGet("EvaluateAdd")) & """"
				tmpJson = tmpJson & ",""PatientType"":""" & Trim(rsGet("PatientType")) & """,""CreateTime"":""" & FormatDate(rsGet("CreateTime"), 10) & """,""Passed"":" & LCase(HR_CBool(rsGet("Passed"))) & "}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""查询成功"",""count"":" & HR_Clng(TotalPut) & ",""data"":[" & tmpJson & "],""limit"":""" & MaxPerPage & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub

Sub Details()
	Dim tmpID : tmpID = HR_Clng(Trim(Request("ID")))
	Dim sqlGet, rsGet
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td.t-w1 {width:100px;text-align:right;background-color:#f3f3f3;padding:9px;} .t-title {width:140px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.w2 ul {margin-left:15px} .w2 li {list-style-type: disc;display:list-item;}" & vbCrlf
	tmpHtml = tmpHtml & "		.w3 {width:60px;text-align:center;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title""><legend><a name="""">mini-CEX <sup>plus</sup>评价详情</a></legend></fieldset>" & vbCrlf
	Response.Write "	<div class=""evalTable"">" & vbCrlf
	

	sqlGet = "Select a.*,b.KSMC"
	sqlGet = sqlGet & " From HR_EvaluateCEX a Left Join HR_Teacher b on a.TeacherID=b.YGDM Where a.ID=" & tmpID
	Set rsGet = Conn.Execute(sqlGet)
		Dim tFocus
		If Not(rsGet.BOF And rsGet.EOF) Then
			tFocus = Trim(rsGet("Focus"))
			If HR_IsNull(tFocus) = False Then tFocus = "<ul><li>" & Replace(tFocus,",","</li><li>") & "</li></ul>"
			Response.Write "		<table class=""layui-table"">" & vbCrlf
			Response.Write "			<tbody><tr><td class=""t-w1"">测评教师：</td><td>" & Trim(rsGet("Teacher")) & "</td><td class=""t-w1"">科室：</td><td>" & Trim(rsGet("KSMC")) & "</td><td class=""t-w1"">职务：</td><td>" & Trim(rsGet("TeacherJob")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">学生姓名：</td><td>" & Trim(rsGet("Student")) & "</td><td class=""t-w1"">学生专业：</td><td>" & Trim(rsGet("Major")) & "</td><td class=""t-w1"">类别：</td><td>" & Trim(rsGet("SutType")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">测评地点：</td><td colspan=""3"">" & Trim(rsGet("EvaluateAdd")) & "</td><td class=""t-w1"">测评时间：</td><td>" & FormatDate(rsGet("EvaluateTime"), 4) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">病人年龄：</td><td>" & Trim(rsGet("PatientAge")) & "岁</td><td class=""t-w1"">病人性别：</td><td>" & Trim(rsGet("PatientGender")) & "</td><td class=""t-w1"">病人类别：</td><td>" & Trim(rsGet("PatientType")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">科室：</td><td colspan=""5"">" & Trim(rsGet("PatientKSMC")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">初步诊断：</td><td colspan=""5"">" & Trim(rsGet("Impression")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">操作名称：</td><td colspan=""5"">" & Trim(rsGet("Treat")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">病情复杂程度：</td><td>" & Trim(rsGet("Complexity")) & "</td><td class=""t-w1"">操作难度：</td><td>" & Trim(rsGet("Difficulty")) & "</td><td class=""t-w1"">测评重点：</td><td class=""w2"">" & tFocus & "</td></tr>" & vbCrlf
			Response.Write "			</tbody>" & vbCrlf
			Response.Write "		</table>" & vbCrlf
			Response.Write "		<table class=""layui-table"">" & vbCrlf
			Response.Write "			<tbody><tr><td class=""t-title"">1.医疗问诊</td><td class=""w2""><ul><li>" & Replace(rsGet("Evaluate1"),",","</li><li>") & "</li></ul></td><td class=""w3"">" & Trim(rsGet("Score1")) & "分</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"">2.体格检查</td><td class=""w2""><ul><li>" & Replace(rsGet("Evaluate2"),",","</li><li>") & "</li></ul></td><td class=""w3"">" & Trim(rsGet("Score2")) & "分</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"">3.临床操作</td><td class=""w2""><ul><li>" & Replace(rsGet("Evaluate3"),",","</li><li>") & "</li></ul></td><td class=""w3"">" & Trim(rsGet("Score3")) & "分</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"">4.临床思维与治疗</td><td class=""w2""><ul><li>" & Replace(rsGet("Evaluate4"),",","</li><li>") & "</li></ul></td><td class=""w3"">" & Trim(rsGet("Score4")) & "分</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"">5.医疗咨询与宣教</td><td class=""w2""><ul><li>" & Replace(rsGet("Evaluate5"),",","</li><li>") & "</li></ul></td><td class=""w3"">" & Trim(rsGet("Score5")) & "分</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"">6.沟通技能与人文关怀</td><td class=""w2""><ul><li>" & Replace(rsGet("Evaluate6"),",","</li><li>") & "</li></ul></td><td class=""w3"">" & Trim(rsGet("Score6")) & "分</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-title"">7.整体表现</td><td class=""w2""><ul><li>" & Replace(rsGet("Evaluate7"),",","</li><li>") & "</li></ul></td><td class=""w3"">" & Trim(rsGet("Score7")) & "分</td></tr>" & vbCrlf
			Response.Write "			</tbody>" & vbCrlf
			Response.Write "		</table>" & vbCrlf
			Response.Write "		<table class=""layui-table"">" & vbCrlf
			Response.Write "			<tbody><tr><td colspan=""4"">本次测评时间：</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">直接观察：</td><td>" & Trim(rsGet("Duration")) & "分钟</td><td class=""t-w1"">反　馈：</td><td>" & Trim(rsGet("BackTime")) & "分钟</td></tr>" & vbCrlf
			Response.Write "				<tr><td colspan=""4"">教师评语：</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">值得肯定：</td><td colspan=""3"">" & Trim(rsGet("Rraise")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">需要改进：</td><td colspan=""3"">" & Trim(rsGet("Mend")) & "</td></tr>" & vbCrlf
			Response.Write "				<tr><td class=""t-w1"">下一步措施：</td><td colspan=""3"">" & Trim(rsGet("Means")) & "</td></tr>" & vbCrlf
			Response.Write "			</tbody>" & vbCrlf
			Response.Write "		</table>" & vbCrlf
		End If
	Set rsGet = Nothing
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
	Dim tTeacherID, tTeacher, tTeacherJob, tStudent, tMajor, tSutType, tEvaluateTime, tEvaluateAdd
	Set rs = Conn.Execute("Select * From HR_EvaluateCEX Where ID=" & tmpID)
		If Not(rs.BOF And rs.EOF) Then
			isModify = True
			tTeacher = Trim(rs("Teacher"))
			tTeacherID = HR_CLng(rs("TeacherID"))
			tTeacherJob = Trim(rs("TeacherJob"))
			tStudent = Trim(rs("Student"))
			tMajor = Trim(rs("Major"))
			tSutType = Trim(rs("SutType"))
			tEvaluateTime = FormatDate(rs("EvaluateTime"), 10)
			tEvaluateAdd = Trim(rs("EvaluateAdd"))
		Else
			tTeacherID = tYGDM
			tTeacher = tYGXM
		End If
	Set rs = Nothing

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .morebtn {padding:3px 0!important;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .tips {padding-left:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form class=""layui-form"" id=""EditForm"" name=""EditForm"" lay-filter=""EditForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layui-form-item""><label class=""layui-form-label"">选择评价人：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""text"" name=""Teacher"" id=""ygxm"" value=""" & tTeacher & """ lay-verify=""required"" autocomplete=""on"" title=""查找评价人"" class=""layui-input txt1"" readonly></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux morebtn""><span class=""layui-btn layui-btn-sm getBtn"" data-code=""ygdm"" data-name=""ygxm"">查找</span><span class=""tips"">请输入关键字搜索教师，必填项</span></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">工　　号：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""number"" name=""TeacherID"" id=""ygdm"" lay-verify=""number"" value=""" & tTeacherID & """ class=""layui-input txt1"" readonly></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">职务：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><select name=""TeacherJob"" id=""TeacherJob""><option value="""">选择职务</option>" & GetTeachJobOption(tTeacherJob, 0) & "</select></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">学生姓名：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""Student"" id=""Student"" lay-verify=""required"" value=""" & tStudent & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">学生专业：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><select name=""Major"" id=""Major""><option value="""">选择学生专业</option>" & GetStuMajorOption(tMajor, 0) & "</select></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">类别：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""SutType"" id=""SutType"" lay-verify=""required"" value=""" & tSutType & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">测评地点：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""EvaluateAdd"" id=""EvaluateAdd"" value=""" & tEvaluateAdd & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">测评时间：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""EvaluateTime"" id=""EvaluateTime"" lay-verify=""date"" autocomplete=""off"" value=""" & tEvaluateTime & """ class=""layui-input dataitem"" readonly></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">学生专业：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""Major"" id=""Major"" lay-verify=""required"" value=""" & tMajor & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
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
	tmpHtml = tmpHtml & "	layui.use([""form"", ""table"", ""element"", ""laydate""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table=layui.table, element=layui.element, form=layui.form, laydate=layui.laydate;" & vbCrlf
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

	tmpHtml = tmpHtml & "		lay("".dataitem"").each(function(){" & vbCrlf			'日期选择窗
	tmpHtml = tmpHtml & "			laydate.render({elem: this, format: 'yyyy-MM-dd'});" & vbCrlf
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

	Response.Write "{""err"":false,""errcode"":0,""icon"":1,""errmsg"":""评价保存成功！""}"
End Sub

Sub AllDataList()			'混合数据
	Dim tYGDM, tYGXM, tKSMC, tPRZC, tXZZW, k
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim layUrl : layUrl = ParmPath & "EvaluateCEX/jsonAllData.html"
	Dim arrTit : arrTit = Split("医疗问诊,体格检查,临床操作,临床思维与治疗,医疗咨询与宣教,沟通技能与人文关怀,整体表现",",")

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

	tmpHtml = "<a href=""" & ParmPath & "EvaluateCEX/Index.html"">" & SiteTitle & "</a><a><cite>列表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title""><legend><a name="""">所有评价</a></legend></fieldset>" & vbCrlf
	Response.Write "	<form name=""SearchForm"" id=""SearchForm"" class=""layui-form soBox"" action="""" method=""get""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input type=""text"" class=""layui-input"" name=""soWord"" value=""" & soWord & """ id=""soWord"" placeholder=""搜索学生"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input type=""text"" class=""layui-input"" name=""soTeacher"" value="""" id=""soTeacher"" placeholder=""搜索教师"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-normal"" data-event=""reload""><i class=""hr-icon hr-icon-top"">&#xeba1;</i>搜索</button></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn"">" & vbCrlf
	'Response.Write "			<button type=""button"" class=""layui-btn hr-btn_peru"" data-event=""update"" name=""update"" title=""更新数据""><i class=""layui-icon layui-anim layui-anim-rotate layui-anim-loop"">&#xe63d;</i></button>" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn hr-btn_darkgreen"" data-event=""refresh"" name=""refresh"" title=""刷新本页""><i class=""hr-icon"">&#xf021;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""toolBtn"">" & vbCrlf	'表头模板
	Response.Write "		<div class=""hr-rows tpltools"">" & vbCrlf
	'Response.Write "			<div class=""layui-btn-group hr-tablebtn"">" & vbCrlf
	'Response.Write "				<button type=""button"" class=""layui-btn hr-btn_fuch"" lay-event=""batchDel"" title=""批量删除""><i class=""hr-icon"">&#xea64;</i></button>" & vbCrlf
	'Response.Write "				<button type=""button"" class=""layui-btn hr-btn_skyblue"" lay-event=""reload"" title=""重载数据""><i class=""hr-icon"">&#xf01e;</i></button>" & vbCrlf
	'Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""sumbar"">共<b class=""Count"">0</b>条记录</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf		'行工具
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""details"" title=""查看评价表""><i class=""hr-icon"">&#xea59;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element"", ""form""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table; element = layui.element, form=layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",height:'full-115',page:true,limit:30,skin:'',limits:[10,20,50,100,200,300,500,800],toolbar: '#toolBtn'" & vbCrlf
	tmpHtml = tmpHtml & "			,text:{none:'暂时没有数据'},cols: [[" & vbCrlf				'设置表头
	tmpHtml = tmpHtml & "				{field:'ID',title:'序号',align:'center', width:60}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Teacher',title:'测评教师',width:95}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherCode',title:'工号',width:95}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'TeacherDepart',title:'科室',width:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'PRZC',title:'职称',width:110}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'XZZW',title:'职务',width:110}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'EvaluateTime',title:'测评时间',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Student',title:'学生姓名',align:'center',width:95}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Major',title:'学生专业',width:125}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'SutType',title:'类　别',width:135}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Major',title:'学生专业',width:125}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'PatientAge',title:'病人年龄',width:85}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'PatientGender',title:'病人性别',width:65}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'PatientKSMC',title:'所在科室',width:125}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'PatientType',title:'病人类别',width:125}" & vbCrlf
	For k=1 To 7
		tmpHtml = tmpHtml & "				,{field:'Evaluate" & k & "',title:'" & arrTit(k-1) & "',minWidth:125}" & vbCrlf
		tmpHtml = tmpHtml & "				,{field:'Score" & k & "',title:'评分',width:65}" & vbCrlf
	Next
	tmpHtml = tmpHtml & "				,{field:'Duration',title:'观察时间',minWidth:105}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'BackTime',title:'反馈时间',minWidth:105}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Rraise',title:'肯定内容',minWidth:105}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Mend',title:'改进内容',minWidth:105}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Means',title:'措施',minWidth:105}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'CreateTime',title:'提交时间',minWidth:125}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'EvaluateAdd',title:'测评地点',minWidth:95}" & vbCrlf
	tmpHtml = tmpHtml & "				,{title:'操作',align:'center',unresize:true,width:100, toolbar:'#barTable'}" & vbCrlf
	tmpHtml = tmpHtml & "			]],url:""" & layUrl & """,parseData: function(res){" & vbCrlf
	tmpHtml = tmpHtml & "				$("".Count"").text(res.count);" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf		'监听工具栏
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""details""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""viewWin"", content:""" & ParmPath & "EvaluateCEX/Details.html?ID="" + data.ID,title:[""查看评价详情"",""font-size:16""],area:[""760px"", ""92%""],moveOut:true });" & vbCrlf
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
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub jsonAllData()
	Dim tmpJson, rsGet, sqlGet, tIntro
	Dim tLimit : tLimit = HR_Clng(Request("limit"))
	Dim tPage : tPage = HR_Clng(Request("page"))
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim soTeacher : soTeacher = Trim(ReplaceBadChar(Request("soTeacher")))
	Dim tYGDM : tYGDM = HR_Clng(Trim(Request("ygdm")))

	sqlGet = "Select a.*,b.KSMC,b.PRZC,b.XZZW"
	sqlGet = sqlGet & " From HR_EvaluateCEX a Left Join (Select YGDM,KSMC,PRZC,XZZW From HR_Teacher) As b On b.YGDM=a.TeacherID Where a.TeacherID>0"
	If HR_IsNull(soWord) = False Then sqlGet = sqlGet & " And a.Student like '%" & soWord &"%'"
	If HR_IsNull(soTeacher) = False Then sqlGet = sqlGet & " And a.Teacher like '%" & soTeacher &"%'"
	sqlGet = sqlGet & " Order By a.CreateTime DESC"
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
			Do While Not rsGet.EOF
				If i > 0 Then tmpJson = tmpJson & ","
				tmpJson = tmpJson & "{""ID"":" & HR_CLng(rsGet("ID")) & ",""TeacherCode"":" & HR_CLng(rsGet("TeacherID")) & ",""Teacher"":""" & Trim(rsGet("Teacher")) & """,""TeacherDepart"":""" & Trim(rsGet("KSMC")) & """,""TeacherJob"":""" & Trim(rsGet("TeacherJob")) & """"
				tmpJson = tmpJson & ",""PRZC"":""" & Trim(rsGet("PRZC")) & """,""XZZW"":""" & Trim(rsGet("XZZW")) & """"
				tmpJson = tmpJson & ",""Student"":""" & Trim(rsGet("Student")) & """,""Major"":""" & Trim(rsGet("Major")) & """,""SutType"":""" & Trim(rsGet("SutType")) & """,""EvaluateTime"":""" & FormatDate(rsGet("EvaluateTime"), 2) & """,""EvaluateAdd"":""" & FilterHtmlToText(rsGet("EvaluateAdd")) & """"
				tmpJson = tmpJson & ",""PRZC"":""" & Trim(rsGet("PRZC")) & """,""XZZW"":""" & Trim(rsGet("XZZW")) & """"
				For k=1 To 7
					tmpJson = tmpJson & ",""Evaluate" & k & """:""" & Replace(rsGet("Evaluate" & k), ","," ") & """,""Score" & k & """:""" & rsGet("Score" & k) & """"
				Next
				tmpJson = tmpJson & ",""TotalScore"":""" & rsGet("TotalScore") & """,""Duration"":""" & Trim(rsGet("Duration")) & """,""BackTime"":""" & Trim(rsGet("BackTime")) & """,""Rraise"":""" & FilterHtmlToText(rsGet("Rraise")) & """,""Mend"":""" & FilterHtmlToText(rsGet("Mend")) & """,""Means"":""" & FilterHtmlToText(rsGet("Means")) & """"
				tmpJson = tmpJson & ",""PatientAge"":""" & Trim(rsGet("PatientAge")) & """,""PatientGender"":""" & Trim(rsGet("PatientGender")) & """,""PatientKSMC"":""" & Trim(rsGet("PatientKSMC")) & """"
				tmpJson = tmpJson & ",""PatientType"":""" & Trim(rsGet("PatientType")) & """,""Impression"":""" & FilterHtmlToText(rsGet("Impression")) & """,""Treat"":""" & FilterHtmlToText(rsGet("Treat")) & """,""Complexity"":""" & FilterHtmlToText(rsGet("Complexity")) & """,""Difficulty"":""" & FilterHtmlToText(rsGet("Difficulty")) & """"
				tmpJson = tmpJson & ",""Focus"":""" & Trim(rsGet("Focus")) & """,""CreateTime"":""" & FormatDate(rsGet("CreateTime"), 10) & """,""Passed"":" & LCase(HR_CBool(rsGet("Passed"))) & "}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""查询成功"",""count"":" & HR_Clng(TotalPut) & ",""data"":[" & tmpJson & "],""limit"":""" & MaxPerPage & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub


'=====================================================================
'函数名：GetTeachJobOption		【返回课程名称下拉】
'=====================================================================
Function GetTeachJobOption(fTeachJob, fType)
	Dim iFun, strFun
	For iFun = 0 To Ubound(arrTeachJob)
		strFun = strFun & "<option value=""" & arrTeachJob(iFun) & """"
		If Trim(fTeachJob) = Trim(arrTeachJob(iFun)) Then strFun = strFun & " selected"
		strFun = strFun & ">" & arrTeachJob(iFun) & "</option>"
	Next
	GetTeachJobOption = strFun
End Function

Function GetStuMajorOption(fMajor, fType)
	Dim iFun, strFun
	For iFun = 0 To Ubound(arrStuMajor)
		strFun = strFun & "<option value=""" & arrStuMajor(iFun) & """"
		If Trim(fMajor) = Trim(arrStuMajor(iFun)) Then strFun = strFun & " selected"
		strFun = strFun & ">" & arrStuMajor(iFun) & "</option>"
	Next
	GetStuMajorOption = strFun
End Function
%>