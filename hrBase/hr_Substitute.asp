<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
SiteTitle = "代课申请管理"
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "jsonList" Call GetJsonList()
	Case "Delete", "Empty" Call Delete()
	Case "Details" Call Details()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim passSwap : passSwap = False		'判断本人是否朋审核权限
	Dim tSwapAuth : tSwapAuth = GetTypeName("HR_User", "SwapPass", "YGDM", UserYGDM)

	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim layUrl : layUrl = ParmPath & "Substitute/jsonList.html"

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.sumbar b {color:#F60;padding:0 2px} .sumbar b.sumDebit{color:#080}" & vbCrlf		'表头汇总
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	tmpHtml = "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)

	tmpHtml = "<a href=""" & ParmPath & "Substitute/Index.html"">" & SiteTitle & "</a><a><cite>列表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form name=""SearchForm"" id=""SearchForm"" class=""layui-form soBox"" action="""" method=""get""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input type=""text"" class=""layui-input"" name=""soWord"" value="""" id=""soWord"" placeholder=""搜索申请人"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-normal"" data-event=""reload"" id=""SearchBtn""><i class=""layui-icon"">&#xe615;</i>搜索</button></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn"">" & vbCrlf
	'Response.Write "			<button type=""button"" class=""layui-btn hr-btn_deon"" data-event=""batchdel"" title=""批量删除""><i class=""layui-icon"">&#xe640;</i></button>" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn layui-bg-green"" data-event=""refresh"" title=""刷新""><i class=""hr-icon"">&#xebbb;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""toolBtn"">" & vbCrlf	'表头模板
	Response.Write "		<div class=""hr-rows tpltools"">" & vbCrlf
	Response.Write "			<div class=""sumbar"">共有<b class=""Count"">0</b>次申请</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""revokeTpl"">" & vbCrlf	'撤销状态
	Response.Write "		<input type=""checkbox"" name=""revoke"" value=""{{d.Revoke}}"" lay-skin=""primary"" disabled lay-filter=""revokeDemo"" {{ d.Revoke ? 'checked' : '' }}>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm"" lay-event=""details"" title=""查看详情""><i class=""hr-icon"">&#xefb9;</i></a>" & vbCrlf
	If tSwapAuth > 0 And UserRank > 1 Then
		Response.Write "			<a class=""layui-btn layui-btn-sm hr-btn_olive"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Else
		Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-disabled"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	End If
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",height:'full-125',page:true,limit:30,skin:'line',limits:[10,15,20,30,50,100,200],toolbar: '#toolBtn'" & vbCrlf
	tmpHtml = tmpHtml & "			,text:{none:'暂时没有调换课申请'},cols: [[" & vbCrlf				'设置表头
	tmpHtml = tmpHtml & "				{type:'checkbox',fixed:'left',unresize:true,align:'center',width:50}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Applyer',title:'申请人',sort:true,fixed:'left',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Department',title:'科室',sort:true,width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Duty',title:'职务',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'ApplyTime',title:'申请时间',align:'center',width:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Reason',title:'代课原因',minWidth:300,event:'viewIntro',style:'cursor: pointer;'}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Replacer',title:'替课教师',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'ItemName',title:'课程分类',width:100}" & vbCrlf	
	tmpHtml = tmpHtml & "				,{field:'Course',title:'课程名称',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Revoke',title:'撤销',align:'center',width:60,templet:'#revokeTpl'}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Passer',title:'教研主任审核',width:160}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Passer1',title:'教学处审核',width:160}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Passer2',title:'教辅审核',width:160}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'ID',title:'序号',sort:true,width:60,align:'center',unresize:true}" & vbCrlf
	tmpHtml = tmpHtml & "				,{title:'操作',fixed:'right',align:'center',unresize:true,width:110, toolbar:'#barTable'}" & vbCrlf
	tmpHtml = tmpHtml & "			]],url:""" & layUrl & """,parseData: function(res){" & vbCrlf
	tmpHtml = tmpHtml & "				$("".Count"").text(res.count);" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		$("".searchBtn button"").on(""click"", function(){" & vbCrlf	'搜索等按钮click事件
	tmpHtml = tmpHtml & "			var btnEvent = $(this).data(""event"");" & vbCrlf
	tmpHtml = tmpHtml & "			switch(btnEvent){" & vbCrlf
	tmpHtml = tmpHtml & "				case ""reload"":" & vbCrlf
	tmpHtml = tmpHtml & "					var sokey = $(""#soWord"").val();" & vbCrlf
	tmpHtml = tmpHtml & "					table.reload(""layList"", { url:""" & layUrl & """, where:{soWord:sokey} }); break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""batchdel"":" & vbCrlf			'批量删除
	tmpHtml = tmpHtml & "					var data = table.checkStatus(""layList"").data;" & vbCrlf
	tmpHtml = tmpHtml & "					if(data.length==0){layer.tips(""请选择您要删除的申请！"","".laytable-cell-checkbox"",{tips: [2, ""#F30""]});return false;}" & vbCrlf
	tmpHtml = tmpHtml & "					var arrID=[]; for(var i=0;i<data.length;i++){ arrID.push(data[i].ID); }" & vbCrlf
	tmpHtml = tmpHtml & "					layer.confirm(""确认要删除选中的 "" + data.length + "" 条申请？<br />删除后将无法恢复。"",{icon:3, title:[""删除警告"",""background-color:#f30""]},function(index){" & vbCrlf
	tmpHtml = tmpHtml & "						$.getJSON(""" & ParmPath & "Substitute/Delete.html"",{ID:arrID.join()}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "							layer.alert(reData.errmsg,{title:""删除结果"",icon:reData.icon},function(){ table.reload(""layList"");layer.close(layer.index); });" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "					}); break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""refresh"":location.reload(); break;" & vbCrlf
	tmpHtml = tmpHtml & "			};" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf		'监听工具栏
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""del""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.confirm(""您确定要删除该申请？"",{icon:3,title:[""删除警告"",""background-color:#f30""]},function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "Substitute/Delete.html"",{ID:data.ID}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.alert(reData.errmsg,{title:""删除结果"",icon:reData.icon, btn:""关闭""},function(){" & vbCrlf
	tmpHtml = tmpHtml & "							if(!reData.err){obj.del();table.reload(""layList"");} layer.close(layer.index);" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""details""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""EditWin"", content:'" & ParmPath & "Substitute/Details.html?ID='+data.ID,title:[""代课详情""],area:[""900px"", ""92%""]});" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = Replace(getPageFoot(1), "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub GetJsonList()
	Dim tmpJson, rsGet, sqlGet
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))

	sqlGet = "Select a.*,b.YGXM As Applyer,b.KSMC,b.XZZW From HR_Swap a Left Join HR_Teacher b On a.YGDM=b.YGDM Where a.newItemID=0 And a.newCourseID=0"
	If HR_CLng(soWord) > 0 Then
		sqlGet = sqlGet & " And a.YGDM=" & soWord
	ElseIf HR_IsNull(soWord) = False Then
		sqlGet = sqlGet & " And b.YGXM like '%" & soWord & "%'"
	End If
	sqlGet = sqlGet & " Order By a.ApplyTime DESC"
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
			Dim tReason, tApplyer, tApplyID, tDepartment, tDuty, tReplacer, tItemName, tCourse
			Dim tPasser, tPasser1, tPasser2, tTemplate, tSheetName, tRevoke
			Do While Not rsGet.EOF
				tSheetName = "HR_Sheet_" & rsGet("ItemID")
				tRevoke = "false" : If HR_Clng(rsGet("Process")) = 5 Then tRevoke = "true"
				If ChkTable(tSheetName) Then
					tReplacer = strGetTypeName("HR_Teacher", "YGXM", "YGDM", HR_Clng(rsGet("Replacer")))
					tPasser = strGetTypeName("HR_Teacher", "YGXM", "YGDM", HR_Clng(rsGet("Passer")))
					tPasser1 = strGetTypeName("HR_Teacher", "YGXM", "YGDM", HR_Clng(rsGet("Passer1")))
					tPasser2 = strGetTypeName("HR_Teacher", "YGXM", "YGDM", HR_Clng(rsGet("Passer2")))
					tItemName = GetTypeName("HR_Class", "ClassName", "ClassID", HR_Clng(rsGet("ItemID")))
					tTemplate = GetTypeName("HR_Class", "Template", "ClassID", HR_Clng(rsGet("ItemID")))
					If tTemplate = "TempTableA" Then
						tCourse = GetTypeName(tSheetName, "VA8", "ID", HR_Clng(rsGet("CourseID")))
					Else
					End If
				End If
				tApplyID = HR_Clng(rsGet("YGDM"))
				tApplyer = Trim(rsGet("Applyer"))
				tDepartment = Trim(rsGet("KSMC"))
				tDuty = Trim(rsGet("XZZW"))
				tReason = FilterHtmlToText(nohtml(rsGet("Reason")))
				If HR_IsNull(tReason) = False Then tReason = GetSubStr(tReason, 35, False)
				If i > 0 Then tmpJson = tmpJson & ","
				tmpJson = tmpJson & "{""ID"":" & rsGet("ID") & ",""Applyer"":""" & tApplyer & """,""ApplyID"":" & HR_CLng(tApplyID) & ",""ItemID"":" & HR_CLng(rsGet("ItemID")) & ",""CourseID"":" & HR_CLng(rsGet("CourseID")) & ""
				tmpJson = tmpJson & ",""Department"":""" & tDepartment & """,""Duty"":""" & tDuty & """,""Reason"":""" & tReason & """,""ReplacerID"":""" & HR_CLng(rsGet("Replacer")) & """"
				tmpJson = tmpJson & ",""ItemName"":""" & tItemName & """,""Course"":""" & tCourse & """,""Passer1"":""" & tPasser1 & """,""Passer2"":""" & tPasser2 & """"
				tmpJson = tmpJson & ",""Replacer"":""" & tReplacer & """,""Passer"":""" & tPasser & """,""Revoke"":" & tRevoke & ",""ApplyTime"":""" & FormatDate(rsGet("ApplyTime"), 10) & """}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""查询成功"",""count"":" & HR_Clng(TotalPut) & ",""data"":[" & tmpJson & "],""limit"":""" & MaxPerPage & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub

Sub Delete()
	If Action = "Empty" Then
		'Conn.Execute("Delete From HR_Swap")
		Response.Write "{""err"":true,""errcode"":0,""errmsg"":""The data emptied"",""icon"":1}" : Exit Sub
	End If
	Dim tmpJson, arrTmpID, iCountID, tmpID : tmpID = Trim(ReplaceBadChar(Request("ID")))
	tmpID = FilterArrNull(tmpID, ",")
	If HR_IsNull(tmpID) = False Then arrTmpID = Split(tmpID, ",") : iCountID = Ubound(arrTmpID) + 1
	If HR_IsNull(tmpID) Then
		tmpJson = "{""err"":true,""errcode"":500,""errmsg"":""未指定删除的数据！"",""icon"":2}"
	Else
		Conn.Execute("Delete From HR_Swap Where ID in(" & tmpID & ")")
		tmpJson = "{""err"":false,""errcode"":0,""errmsg"":""共有 " & HR_CLng(iCountID) & " 条记录删除成功！"",""icon"":1}"
	End If
	Response.Write tmpJson
End Sub

Sub Details()
	
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	SiteTitle = "代课详情"

	Dim tProposer, tReason, tReplacer, tReplacerTime, tPasser, tPassTime, tPasser1, tPassTime1, tPasser2, tPassTime2, tProcess, tExplain
	Dim tApplyTime, tCourse, tStuClass, tPlace, tCourseDate, tCourseTime, tPeriod
	Dim tSheetName, tTemplate, tItemName, tCourseID, tTeachDate
	Dim tProposerCode, tReplacerCode, tReplacerDept, tReplacerZC, tReplacerZW, tPasserCode, tPasser1Code, tPasser2Code
	Dim tProposerDept, tProposerZW, tProposerZC
	Dim tReplacePass, tPasserPass, tPasser1Pass, tPasser2Pass
	Dim tVA3, tVA4, tVA5, tVA6, tVA7, tVA8, tVA9, tVA10, tVA11, tVA12
	Dim newVA3, newVA4, newVA5, newVA6, newVA7, newVA8, newVA9, newVA10, newVA11, newVA12
	sqlTmp = "Select a.*,(Select YGXM From HR_Teacher Where YGDM=a.YGDM) As Proposer"
	sqlTmp = sqlTmp & ",(Select YGXM From HR_Teacher Where YGDM=a.Replacer) As ReplacerName"
	sqlTmp = sqlTmp & ",(Select YGXM From HR_Teacher Where YGDM=a.Passer) As PasserName"
	sqlTmp = sqlTmp & ",(Select YGXM From HR_Teacher Where YGDM=a.Passer1) As PasserName1"
	sqlTmp = sqlTmp & ",(Select YGXM From HR_Teacher Where YGDM=a.Passer2) As PasserName2"
	sqlTmp = sqlTmp & ",(Select Template From HR_Class Where ClassID=a.ItemID) As Template,(Select ClassName From HR_Class Where ClassID=a.ItemID) As ItemName From HR_Swap a Where a.ID=" & tmpID
	Set rs = Conn.Execute(sqlTmp)
		If Not(rs.BOF And rs.EOF) Then
			tSheetName = "HR_Sheet_" & rs("ItemID")		'数据表名
			tTemplate = rs("Template")
			tItemName = rs("ItemName")
			tProposer = Trim(rs("Proposer"))	'申请人
			tProposerCode = HR_Clng(rs("YGDM"))	'申请人工号
			tProposerDept = strGetTypeName("HR_Teacher", "KSMC", "YGDM", tProposerCode)
			tProposerZW = strGetTypeName("HR_Teacher", "XZZW", "YGDM", tProposerCode)
			tProposerZC = strGetTypeName("HR_Teacher", "PRZC", "YGDM", tProposerCode)

			tReason = Trim(rs("Reason"))		'申请理由
			tReplacer = Trim(rs("ReplacerName"))'替换教师
			tReplacerTime = FormatDate(rs("ReplacerTime"), 10)'替换教师
			tProcess = HR_Clng(rs("Process"))
			tReplacePass = HR_Clng(rs("ReplacePass"))

			tReplacerCode = Trim(rs("Replacer"))	'替换教师工号
			tReplacerDept = strGetTypeName("HR_Teacher", "KSMC", "YGDM", tReplacerCode)
			tReplacerZC = strGetTypeName("HR_Teacher", "PRZC", "YGDM", tReplacerCode)
			tReplacerZW = strGetTypeName("HR_Teacher", "XZZW", "YGDM", tReplacerCode)

			tPasser = Trim(rs("PasserName"))	'教研主任
			tPasserCode = HR_Clng(rs("Passer"))	'教研主任工号
			tPassTime = FormatDate(rs("PassTime"), 10)
			tPasserPass = HR_Clng(rs("PasserPass"))

			tPasser1 = Trim(rs("PasserName1"))	'教学处审核人
			tPassTime1 = FormatDate(rs("PassTime1"), 10)	'教学处审核时间
			tPasser1Pass = HR_Clng(rs("Passer1Pass"))

			tPasser2 = Trim(rs("PasserName2"))	'教辅审核人
			tPassTime2 = FormatDate(rs("PassTime2"), 10)	'教辅审核时间
			tPasser2Pass = HR_Clng(rs("Passer2Pass"))

			tExplain = Trim(rs("Explain"))	'审核说明
			tApplyTime = FormatDate(rs("ApplyTime"), 10)	'申请提交时间
			tCourseID = HR_Clng(rs("CourseID"))
			tVA3 = Trim(rs("VA3"))
			tVA4 = Trim(rs("VA4"))
			tVA5 = Trim(rs("VA5"))
			tVA6 = Trim(rs("VA6"))
			tVA7 = Trim(rs("VA7"))
			tVA8 = Trim(rs("VA8"))
			tVA9 = Trim(rs("VA9"))
			tVA10 = Trim(rs("VA10"))
			tVA11 = Trim(rs("VA11"))
			tVA12 = Trim(rs("VA12"))
			newVA3 = Trim(rs("newVA3"))
			newVA4 = Trim(rs("newVA4"))
			newVA5 = Trim(rs("newVA5"))
			newVA6 = Trim(rs("newVA6"))
			newVA7 = Trim(rs("newVA7"))
			newVA8 = Trim(rs("newVA8"))
			newVA9 = Trim(rs("newVA9"))
			newVA10 = Trim(rs("newVA10"))
			newVA11 = Trim(rs("newVA11"))
			newVA12 = Trim(rs("newVA12"))
		End If
	Set rs = Nothing
	If ChkTable(tSheetName) Then
		sql = "Select a.* From " & tSheetName & " a Where a.ID=" & tCourseID
		Set rs = Conn.Execute(sql)
			If Not(rs.BOF And rs.EOF) Then
				If tTemplate = "TempTableA" Then
					tCourse = rs("VA8")
					tStuClass = rs("VA10")
					tPlace = rs("VA11") & " " & rs("VA12")
					tTeachDate = FormatDate(ConvertNumDate(rs("VA4")), 4) & " 第" & Trim(rs("VA7")) & "节"
				Else
					tCourse = rs("VA6")
				End If
			End If
		Set rs = Nothing
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-step-box {padding-bottom:20px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-step-box .hr-step {width:25%; flex-shrink:0; box-sizing:border-box; padding:0 10px}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-step-box .hr-step h4 {font-size:1.1rem;padding-bottom:5px;}" & vbCrlf

	tmpHtml = tmpHtml & "		.hr-swap-items dl {padding:5px 0}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dt {width:100px;text-align:right;color:#999;flex-shrink:0}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dd {flex-grow:2;padding-left:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-rows-pan {display:flex;box-sizing:border-box;} .hr-rows-pan .row2 {width:50%;box-sizing:border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-rows-pan .row2 h3 {padding-left:15px;font-size:1.5rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.step0, .step1, .step2 {border:1px solid #ccc;padding:1px;font-size:14px;position:relative;left:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.step0 {color:#4FC0E8;border-color:#4FC0E8;}" & vbCrlf
	tmpHtml = tmpHtml & "		.step1 {color:#A0D468;border-color:#A0D468;}" & vbCrlf
	tmpHtml = tmpHtml & "		.step2 {color:#EC87BF;border-color:#EC87BF;}" & vbCrlf
	tmpHtml = tmpHtml & "		.passbtn {text-align:center;padding:20px 0 15px 0}" & vbCrlf
	
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Dim titField : titField = "<legend>审核进度</legend>"
	If tProcess >=5 Then titField = "<legend style=""color:#f30"">审核进度[已撤销]</legend>"
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title"">" & titField & "</fieldset>" & vbCrlf
	Response.Write "	<div class=""hr-rows hr-stretch hr-step-box"">" & vbCrlf
	Response.Write "		<div class=""hr-step""><h4>代课教师" & ShowAgreeProcess(tReplacePass) & "</h4><h5>" & tReplacer & "[" & tReplacerCode & "]</h5><h5>确认时间：" & tReplacerTime & "</h5></div>" & vbCrlf
	Response.Write "		<div class=""hr-step""><h4>教研主任" & ShowAgreeProcess(tPasserPass) & "</h4><h5>审核人：" & tPasser & "[" & tPasserCode & "]</h5><h5>审核时间：" & tPassTime & "</h5></div>" & vbCrlf
	Response.Write "		<div class=""hr-step""><h4>教学处" & ShowAgreeProcess(tPasser1Pass) & "</h4><h5>审核人：" & tPasser1 & "</h5><h5>审核时间：" & tPassTime1 & "</h5></div>" & vbCrlf
	Response.Write "		<div class=""hr-step""><h4>教辅" & ShowAgreeProcess(tPasser2Pass) & "</h4><h5>审核人：" & tPasser2 & "</h5><h5>审核时间：" & tPassTime2 & "</h5></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""layui-progress layui-progress-big"" lay-showPercent=""true"">" & vbCrlf
	If HR_CLng(tProcess) >= 5 Then
		Response.Write "		<div class=""layui-progress-bar layui-bg-red"" lay-percent=""100%""></div>" & vbCrlf
	Else
		Response.Write "		<div class=""layui-progress-bar layui-bg-red"" lay-percent=""" & HR_CLng(tProcess) & "/4""></div>" & vbCrlf
	End If
	Response.Write "	</div>" & vbCrlf

	'Response.Write "	<div class=""passbox"">" & vbCrlf		'审核
	'Response.Write "		<em class=""passbtn"">" & vbCrlf
	'Response.Write "			<button type=""button"" class=""layui-btn layui-bg-green"" data-event=""refresh"" title=""同意""><i class=""hr-icon"">&#xebbb;</i>同意</button>" & vbCrlf
	'Response.Write "			<button type=""button"" class=""layui-btn layui-bg-green"" data-event=""refresh"" title=""拒绝""><i class=""hr-icon"">&#xebbb;</i>拒绝</button>" & vbCrlf
	'Response.Write "		</em>" & vbCrlf
	'Response.Write "	</div>" & vbCrlf

	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title""><legend>申请人</legend></fieldset>" & vbCrlf
	Response.Write "	<div class=""hr-swap-items"">" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>申请姓名：</dt><dd>" & tProposer & " [" & tProposerCode & "]</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>科　室：</dt><dd>" & tProposerDept & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>职务/职称：</dt><dd>" & tProposerZC & " " & tProposerZW & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>申请时间：</dt><dd>" & tApplyTime & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>原　因：</dt><dd>" & tReason & "</dd></dl>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title""><legend>代课教师</legend></fieldset>" & vbCrlf
	Response.Write "	<div class=""hr-swap-items"">" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>代课教师：</dt><dd>" & tReplacer & "　" & tReplacerCode & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>科　室：</dt><dd>" & tReplacerDept & "</dd></dl>" & vbCrlf
	Response.Write "		<dl class=""hr-rows""><dt>职务/职称：</dt><dd>" & tReplacerZC & " " & tReplacerZW & "</dd></dl>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title site-title""><legend>课程信息</legend></fieldset>" & vbCrlf
	Response.Write "	<div class=""hr-rows-pan"">" & vbCrlf
	Response.Write "		<div class=""hr-swap-items row2"">" & vbCrlf
	Response.Write "			<h3 class=""hr-swap-tit"">原课程信息</h3>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>项目名称：</dt><dd>" & tItemName & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>授课时间：</dt><dd>" & tVA4 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>课时：</dt><dd>" & tVA3 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>周次：</dt><dd>" & tVA5 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>星期：</dt><dd>" & tVA6 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>节次：</dt><dd>" & tVA7 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>课程名称：</dt><dd>" & tVA8 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>课程内容：</dt><dd>" & tVA9 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>授课对象：</dt><dd>" & tVA10 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>校(院)区：</dt><dd>" & tVA11 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>授课地点：</dt><dd>" & tVA12 & "</dd></dl>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""hr-swap-items row2"">" & vbCrlf
	Response.Write "			<h3 class=""hr-swap-tit"">新课程信息</h3>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>项目名称：</dt><dd>" & tItemName & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>授课时间：</dt><dd>" & newVA4 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>课时：</dt><dd>" & newVA3 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>周次：</dt><dd>" & newVA5 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>星期：</dt><dd>" & newVA6 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>节次：</dt><dd>" & newVA7 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>课程名称：</dt><dd>" & newVA8 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>课程内容：</dt><dd>" & newVA9 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>授课对象：</dt><dd>" & newVA10 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>校(院)区：</dt><dd>" & newVA11 & "</dd></dl>" & vbCrlf
	Response.Write "			<dl class=""hr-rows""><dt>授课地点：</dt><dd>" & newVA12 & "</dd></dl>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element;" & vbCrlf

	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = Replace(getPageFoot(1), "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub EditBody()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tYGDM : tYGDM = Trim(Request("ygdm"))
	Dim tYGXM : tYGXM = strGetTypeName("HR_Teacher", "ygxm", "ygdm", tYGDM)
	Dim isModify : isModify = HR_CBool(Request("Modify"))
	Dim tTeacher, tTeacherID, tItemID, tItemName, tCourseID, tSheetName
	Dim tVA3, tVA4, tVA5, tVA6, tVA7, tVA8, tVA9, tVA10, tVA11, tVA12
	Set rs = Conn.Execute("Select * From HR_Evaluate Where ID=" & tmpID)
		If Not(rs.BOF And rs.EOF) Then
			tItemID = HR_CLng(rs("ItemID"))
			tVA3 = HR_CDbl(rs("newVA3"))
			tVA4 = FormatDate(rs("newVA4"), 2)
			tVA5 = Trim(rs("newVA5"))
			tVA6 = Trim(rs("newVA6"))
			tVA7 = Trim(rs("newVA7"))
			tVA8 = Trim(rs("newVA8"))
			tVA9 = Trim(rs("newVA9"))
			tVA10 = Trim(rs("newVA10"))
			tVA11 = Trim(rs("newVA11"))
			tVA12 = Trim(rs("newVA12"))
		End If
	Set rs = Nothing
	tSheetName = "HR_Sheet_" & tItemID
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .morebtn {padding:3px 0!important;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .tips {padding-left:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-slider {flex-grow:1;} .slider{box-sizing:border-box;padding:1px 5px 0 5px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form class=""layui-form"" id=""EditForm"" name=""EditForm"" lay-filter=""EditForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layui-form-item""><label class=""layui-form-label"">选择评价人：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""text"" name=""YGXM"" id=""ygxm"" value=""" & tYGXM & """ lay-verify=""required"" autocomplete=""on"" title=""查找评价人"" class=""layui-input txt1"" readonly></div>" & vbCrlf
	Response.Write "			<div class=""layui-form-mid layui-word-aux morebtn""><span class=""layui-btn layui-btn-sm getBtn"" data-code=""ygdm"" data-name=""ygxm"">查找</span><span class=""tips"">请输入关键字搜索教师，必填项</span></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">工　　号：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""text"" name=""YGDM"" id=""ygdm"" lay-verify=""required"" value=""" & tYGDM & """ class=""layui-input txt1"" readonly></div>" & vbCrlf
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
	Response.Write "				<div class=""layui-input-inline"" id=""CourseSelect""><select name=""CourseID"" id=""CourseID"" lay-filter=""CourseOption""><option value="""">请选择课程</option>" & GetItemCourseOption(tItemID, tCourseID, tYGDM, "") & "</select></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">授课时间：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""VA4"" id=""VA4"" lay-verify=""date"" value=""" & tVA4 & """ class=""layui-input dataitem"" readonly></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">星　期：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""text"" name=""VA6"" id=""VA6"" lay-verify=""required"" value=""" & tVA6 & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">周　次：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""number"" name=""VA5"" id=""VA5"" lay-verify=""required"" value=""" & tVA5 & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""layui-inline""><label class=""layui-form-label"">学　时：</label>" & vbCrlf
	Response.Write "				<div class=""layui-input-inline""><input type=""number"" name=""VA3"" id=""VA3"" lay-verify=""required"" value=""" & tVA3 & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item hr-rows"">" & vbCrlf		'节次
	Response.Write "			<label class=""layui-form-label"">节　次：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:80px""><input type=""text"" name=""VA7"" id=""VA7"" value=""" & tVA7 & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1"" readonly></div>"
	Response.Write "			<div class=""hr-slider""><div id=""slide7"" class=""slider""></div></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">课程名称：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:45%""><input type=""text"" name=""VA8"" id=""VA8"" lay-verify=""required"" value=""" & tVA8 & """ class=""layui-input txt1""></div>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><select class=""selectto"" name=""SelectVA8"" title=""VA8"" lay-search=""""><option value="""">选择/搜索</option>" & GetCourseOption(tVA8, 0) & "</select></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">授课内容：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:45%""><input type=""text"" name=""VA9"" id=""VA9"" lay-verify=""required"" value=""" & tVA9 & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">授课对象：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:45%""><input type=""text"" name=""VA10"" id=""VA10"" lay-verify=""required"" value=""" & tVA10 & """ class=""layui-input""></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">校(院)区：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:45%""><input type=""text"" name=""VA11"" id=""VA11"" lay-verify=""required"" value=""" & tVA11 & """ class=""layui-input txt1""></div>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><select class=""selectto"" name=""SelectVA11"" title=""VA11"" lay-search=""""><option value="""">选择/搜索</option>" & GetCampusOption(tVA11, 0) & "</select></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">授课地点：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline"" style=""width:45%""><input type=""text"" name=""VA12"" id=""VA12"" lay-verify=""required"" value=""" & tVA12 & """ class=""layui-input txt1""></div>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><select class=""selectto"" name=""SelectVA12"" title=""VA12"" lay-search=""""><option value="""">选择/搜索</option>" & GetClassroomOption(tVA12, 0) & "</select></div>" & vbCrlf
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

	tmpHtml = tmpHtml & "		lay("".dataitem"").each(function(){" & vbCrlf
	tmpHtml = tmpHtml & "			laydate.render({elem: this, format: 'yyyy-MM-dd',done:function(value, date, endDate){" & vbCrlf
	tmpHtml = tmpHtml & "				var today = new Array('日','一','二','三','四','五','六'), day = new Date(value);" & vbCrlf
	tmpHtml = tmpHtml & "				var week = today[day.getDay()];$(""#VA6"").val(week);" & vbCrlf
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
	tmpHtml = tmpHtml & "			$.post(""" & ParmPath & "Substitute/CourseOption.html"", {ItemID:data.value, ygdm:ygdm, CourseID:course}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#CourseSelect"").html(rsStr);form.render(""select"");" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		form.on(""select(CourseOption)"", function(data){" & vbCrlf		'监听课程下拉，更新课程数据
	tmpHtml = tmpHtml & "			var ygdm = $(""#ygdm"").val(), itemid = $(""#ItemID"").val();" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "Substitute/GetCourseData.html"", {ItemID:itemid, ygdm:ygdm, CourseID:data.value}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#VA4"").val(rsStr.VA4); $(""#VA3"").val(rsStr.VA3); $(""#VA5"").val(rsStr.VA5); $(""#VA6"").val(rsStr.VA6);" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#VA7"").val(rsStr.VA7); $(""#VA8"").val(rsStr.VA8); $(""#VA9"").val(rsStr.VA9); $(""#VA10"").val(rsStr.VA10);" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#VA11"").val(rsStr.VA11); $(""#VA12"").val(rsStr.VA12);" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		form.on(""select"", function(data){" & vbCrlf					'监听下拉，并赋值到指定表单
	tmpHtml = tmpHtml & "			var el = data.elem.title;" & vbCrlf
	tmpHtml = tmpHtml & "			$("".txt1"").each(function(){" & vbCrlf
	tmpHtml = tmpHtml & "				if($(this).attr(""name"")==el){$(this).val(data.value)};" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		var val7 = $(""#VA7"").val(), arrVal=[3,5];if(val7!=""""){arrVal=val7.split(""-"")}" & vbCrlf
	tmpHtml = tmpHtml & "		var slider7 = slider.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#slide7"",range: true,max: 20,theme:""#809"",value:[arrVal[0],arrVal[1]]," & vbCrlf
	tmpHtml = tmpHtml & "			change: function(value){$(""#VA7"").val(value[0] + ""-"" + value[1])}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = Replace(getPageFoot(1), "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
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

Function ShowAgreeProcess(fAgree)
	Dim fStr : fStr = "<span class=""step0"">待审</span>"
	If HR_CLng(fAgree) = 1 Then fStr = "<span class=""step1"">同意</span>"
	If HR_CLng(fAgree) = 2 Then fStr = "<span class=""step2"">拒绝</span>"
	ShowAgreeProcess = fStr
End Function
%>