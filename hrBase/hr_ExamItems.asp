<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./incKPI.asp"-->
<!--#include file="./hr_ExamItemsInc.asp"-->

<%
Server.ScriptTimeout=3600		'1小时
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim SubButTxt : SiteTitle = "考核项目管理"

Dim arrUnit : arrUnit = Split(XmlText("Common", "Unit", ""), "|")
Dim arrItemType : arrItemType = Split(XmlText("Common", "ItemType", ""), "|")
Dim arrStudentType : arrStudentType = Split(XmlText("Common", "StudentType", ""), "|")
Dim tItemType : tItemType = HR_Clng(Request("Type"))
Dim scriptCtrl

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveForm" Call SaveForm()
	Case "jsonList" Call GetJsonList()
	Case "Preview" Call Preview()
	Case "Delete" Call Delete()

	Case "DownExcel" Call DownExcel()
	Case "Template" Call TemplateForm()
	Case "DownExcel" Call DownExcel()
	Case "ShowTemp" Call ShowTemp()

	Case "Ratio" Call RatioForm()
	Case "SaveRatio" Call SaveRatio()

	Case "EditLevel" Call EditLevel()
	Case "ListLevel" Call ListLevel()
	Case "LevelForm" Call LevelForm()
	Case "SaveLevel" Call SaveLevel()
	Case "DelLevel" Call DelLevel()

	Case "SetFieldTitle" Call SetupField()
	Case "SaveField" Call SaveField()

	Case "EditGrade" Call EditGrade()
	Case "GradeData" Call GradeData()
	Case "SaveGrade" Call SaveGrade()
	Case "DelGrade" Call DelGrade()

	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	If tItemType = 0 Then tItemType = 1
	Dim layUrl : layUrl = ParmPath & "ExamItems/jsonList.html?Type=" & tItemType

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.iframe-nav .navBtn .navLayer {font-size: 16px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.iframe-nav .navBtn .navLayer i {font-size: 16px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding: 5px 8px;}" & vbCrlf

	tmpHtml = tmpHtml & "		.hr-pop-fix {position: absolute;}" & vbCrlf
	tmpHtml = tmpHtml & "		.formBtn .layui-btn-sm {line-height:27px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.formBtn .layui-btn-sm i {font-size: 16px!important;position: relative;top:3px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")

	tmpHtml = "<a href=""" & ParmPath & "ExamItems/Index.html?Type=" & tItemType & """>" & SiteTitle & "</a><a><cite>" & arrItemType(tItemType-1) & "项目列表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form soBox""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input class=""layui-input"" name=""soWord"" value=""" & soWord & """ id=""soWord"" placeholder=""项目名称"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn""><button class=""layui-btn layui-btn-normal"" data-type=""reload"" id=""SearchBtn""><i class=""hr-icon"">&#xeba1;</i>搜索</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-bg-cyan addNew"" data-type=""" & tItemType & """ title=""新增一级项目""><i class=""hr-icon"">&#xee41;</i>新增一级项目</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_fuch refresh"" data-type=""refresh"" id=""refresh"" title=""刷新本页""><i class=""hr-icon"">&#xebbb;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""tabhead"">" & vbCrlf		'表头模板
	Response.Write "		<div class=""totalbar"">数据汇总：<b class=""count"">0</b>个项目共<b class=""total"">0</b>条数据</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-bg-cyan"" lay-event=""detail"" title=""查看详情""><i class=""hr-icon"">&#xefb9;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	Response.Write "			{{#  if(d.Depth===""1""){ }}" & vbCrlf
	Response.Write "				<a class=""layui-btn layui-btn-sm layui-btn-disabled"" title=""不能添加下级项目""><i class=""hr-icon"">&#xe611;</i></a>" & vbCrlf
	Response.Write "			{{#  }else{ }}" & vbCrlf
	Response.Write "				<a class=""layui-btn layui-btn-sm"" lay-event=""addChild"" title=""添加二级项目""><i class=""hr-icon"">&#xe3ba;</i></a>" & vbCrlf
	Response.Write "			{{#  } }}" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm hr-btn_olive"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "			{{#  if(d.Child===0){ }}" & vbCrlf
	Response.Write "				<a class=""layui-btn layui-btn-sm hr-btn_peru"" lay-event=""setKey"" title=""字段标题""><i class=""hr-icon"">&#xf390;</i></a>" & vbCrlf
	Response.Write "			{{#  }else{ }}" & vbCrlf
	Response.Write "				<a class=""layui-btn layui-btn-sm layui-btn-disabled"" title=""项目分类不用设置""><i class=""hr-icon"">&#xf390;</i></a>" & vbCrlf
	Response.Write "			{{#  } }}" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""layer"", ""form"", ""upload"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, upload = layui.upload;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element, form = layui.form, layer = layui.layer;" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",limit:0, height:""full-115"", skin:""line"", toolbar:""#tabhead""" & vbCrlf
	tmpHtml = tmpHtml & "			,text:{none:'暂时没有考核项目'},cols: [[" & vbCrlf				'设置表头
	tmpHtml = tmpHtml & "				{field:'ItemID',title:'序号',sort:true,width:60,align:'center',unresize:true}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'ItemName',title:'项目名称',width:180}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Template',title:'模型',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Unit',title:'单位',width:70}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Ratio',title:'系数',width:130}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Intro',title:'备注说明',minWidth:300,event:'viewIntro',style:'cursor: pointer;'}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Count',title:'课程数',align:'right',width:60}" & vbCrlf
	tmpHtml = tmpHtml & "				,{fixed:'right',title:'操作',align:'center',unresize:true,width:240, toolbar:'#barTable'}" & vbCrlf
	tmpHtml = tmpHtml & "			]],url:""" & layUrl & """,parseData:function(res){$(""b.count"").html(res.count);$(""b.total"").html(res.sumtotal);}" & vbCrlf		'设置异步接口
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		$(""#refresh"").on(""click"", function(index){window.location.reload();});" & vbCrlf		'刷新当前页
	tmpHtml = tmpHtml & "		$("".addNew"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			layer.open({type:1,id:""AddWin"",content:"""", title:[""添加考核项目"",""font-size:16""], area:[""750px"", ""80%""], maxmin:true });" & vbCrlf
	tmpHtml = tmpHtml & "			var loadTips = layer.load(1), addType = $(this).data(""type"");" & vbCrlf
	tmpHtml = tmpHtml & "			$.get(""" & ParmPath & "ExamItems/AddNew.html"", {TypeID:addType, Modify:false}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#AddWin"").html(strForm);form.render();layer.close(loadTips);" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#DataModel"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "					var Temp1 = $(""#Template"").val(), itemName = $(""#ItemName"").val();" & vbCrlf
	tmpHtml = tmpHtml & "					layer.open({type:2,content:""" & ParmPath & "ExamItems/ShowTemp.html?Temp="" + Temp1 + ""&itemName="" + itemName ,title:[""查看模板"",""font-size:16""],area:[""90%"", ""70%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "				form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	tmpHtml = tmpHtml & "					$.post(""" & ParmPath & "ExamItems/SaveForm.html"", $(""#EditForm"").serialize(), function(Result){" & vbCrlf
	tmpHtml = tmpHtml & "						var reJson = eval(""("" + Result + "")""), icon=2; if(reJson.Return){icon=1;}" & vbCrlf
	tmpHtml = tmpHtml & "						layer.alert(reJson.reMessge, {icon:icon},function(layero, index){" & vbCrlf
	tmpHtml = tmpHtml & "							if(reJson.Return){layer.closeAll();window.location.reload();}" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					return false;" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""detail""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""viewWin"", content:""" & ParmPath & "ExamItems/Preview.html?ID="" + data.ItemID, title:[""查看考核项目信息"",""font-size:16""],area:[""700px"", ""80%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""addChild""||obj.event === ""edit""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.load(1);" & vbCrlf
	tmpHtml = tmpHtml & "				$.get(""" & ParmPath & "ExamItems/Edit.html"",{ItemID:data.ItemID,TypeID:data.ItemType,Eve:obj.event}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "					layer.open({type:1,content:strForm,title:[""编辑考核项目"",""font-size:16""],area:[""630px"", ""80%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					layer.closeAll(""loading"");form.render();" & vbCrlf
	tmpHtml = tmpHtml & "					$(""#DataModel"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "						var Temp1 = $(""#Template"").val(), itemName = $(""#ItemName"").val();" & vbCrlf
	tmpHtml = tmpHtml & "						layer.open({type:2,content:""" & ParmPath & "ExamItems/ShowTemp.html?Temp="" + Temp1 + ""&itemName="" + itemName ,title:[""查看模板"",""font-size:16""],area:[""85%"", ""70%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					$(""#btnExcel"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "						var tid = $(""#ClassID"").val(), itemName = $(""#ItemName"").val();" & vbCrlf
	tmpHtml = tmpHtml & "						layer.open({type:2,content:""" & ParmPath & "Course/ExcelTemp.html?ItemID="" + tid + ""&itemName="" + itemName ,title:[""Excel模板文件"",""font-size:16""],area:[""90%"", ""70%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf

	tmpHtml = tmpHtml & "					$(""#btnRatio"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "						var arrChkValue = [];" & vbCrlf
	tmpHtml = tmpHtml & "						$(""input[name='StudentType']:checked"").each(function(index, elem){arrChkValue.push($(elem).val());});console.log(arrChkValue.join());" & vbCrlf
	tmpHtml = tmpHtml & "						layer.open({type:2,content:""" & ParmPath & "ExamItems/Ratio.html?TypeID="" + data.ItemType + ""&ItemID="" + data.ItemID + ""&Eve="" + obj.event + ""&stuType="" + arrChkValue.join(),title:[""编辑考核系数"",""font-size:16""],area:[""700px"", ""450px""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf

	tmpHtml = tmpHtml & "					$(""#btnLevel"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "						var arrChkValue = [];" & vbCrlf
	tmpHtml = tmpHtml & "						$(""input[name='StudentType']:checked"").each(function(index, elem){arrChkValue.push($(elem).val());});" & vbCrlf
	tmpHtml = tmpHtml & "						layer.open({type:2,content:""" & ParmPath & "ExamItems/EditLevel.html?TypeID="" + data.ItemType + ""&ItemID="" + data.ItemID + ""&stuType="" + arrChkValue.join(),title:[""编辑级别"",""font-size:16""],area:[""750px"", ""75%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf

	tmpHtml = tmpHtml & "					form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.load(1,{shade:[0.1, ""#000""]});" & vbCrlf
	tmpHtml = tmpHtml & "						$.post(""" & ParmPath & "ExamItems/SaveForm.html"", $(""#EditForm"").serialize(), function(result){" & vbCrlf
	tmpHtml = tmpHtml & "							var reData = eval(""("" + result + "")"");layer.closeAll(""loading"");" & vbCrlf
	tmpHtml = tmpHtml & "							if(reData.Return){" & vbCrlf
	tmpHtml = tmpHtml & "								layer.alert(reData.reMessge, {icon:1},function(layero, index){layer.closeAll();table.reload(""layList"");});" & vbCrlf
	tmpHtml = tmpHtml & "							}else{" & vbCrlf
	tmpHtml = tmpHtml & "								layer.alert(reData.reMessge, {icon:2});" & vbCrlf
	tmpHtml = tmpHtml & "							}" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "						return false;" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf

	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""level""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2,content:'" & ParmPath & "ExamItems/EditLevel.html?ItemID='+ data.ItemID +'&TypeID='+data.ItemType,title:[""编辑考核等级"",""font-size:16""],area:[""700px"", ""90%""],maxmin:true});form.render();" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""setKey""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2,content:'" & ParmPath & "ExamItems/SetFieldTitle.html?ItemID='+ data.ItemID +'&TypeID='+data.ItemType,title:[""编辑字段标题"",""font-size:16""],area:[""500px"", ""70%""],offset:[""100px"", ""100px""],scrollbar:false});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""del""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.confirm(""警告：您确认删除选中的考核项目吗？<br />相关的课程进度、业绩报表等将同步删除而且无法恢复！"", {icon:3,title: ""删除警告""}, function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "ExamItems/Delete.html"",{ItemID:data.ItemID,Type:data.ItemType}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						var icon = 2; if(reData.Return){icon = 1;}" & vbCrlf
	tmpHtml = tmpHtml & "						layer.alert(reData.reMessge, {icon:icon,title: ""删除结果提示""},function(layero, index){" & vbCrlf
	tmpHtml = tmpHtml & "							layer.closeAll();table.reload(""TableList"");" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "					}); layer.close(index);" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""template""){" & vbCrlf
	tmpHtml = tmpHtml & "				$.get(""" & ParmPath & "ExamItems/Template.html"",{ItemID:data.ItemID,TypeID:data.ItemType}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "					layer.open({type:1,content:strForm,title:[""查看Excel数据模板"",""font-size:16""],area:[""95%"", ""90%""],maxmin:true});form.render();" & vbCrlf
	tmpHtml = tmpHtml & "					$(""#ExportPost"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "						var ItemID = $(""#ItemID"").val(), Template = $(""#Template"").val();" & vbCrlf
	tmpHtml = tmpHtml & "						$.getJSON(""" & ParmPath & "ExamItems/DownExcel.html"", {ItemID:ItemID, Template:Template}, function(result){" & vbCrlf
	tmpHtml = tmpHtml & "							if(result.Return){" & vbCrlf
	tmpHtml = tmpHtml & "								var downDoc = result.reMessge + ""<br><i class='hr-icon'>&#xf019;</i>：<a href="" + result.fileUrl + "">"" + result.fileUrl + ""</a>"";" & vbCrlf
	tmpHtml = tmpHtml & "								layer.alert(downDoc, {icon:1},function(layero, index){" & vbCrlf
	tmpHtml = tmpHtml & "									var downUrl = result.fileUrl;" & vbCrlf
	tmpHtml = tmpHtml & "								});" & vbCrlf
	tmpHtml = tmpHtml & "							}else{" & vbCrlf
	tmpHtml = tmpHtml & "								layer.alert(result.reMessge, {icon:2});" & vbCrlf
	tmpHtml = tmpHtml & "							}" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "						return false;" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf

	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub
Sub GetJsonList()
	Dim tmpJson, tmpData, rsGet, sqlGet, vCount, vMSG, tParentClass, rsCount, tCount, tSum
	Dim ItemType, ItemTypeName, tParentID, ParentClass, tItemName, tTemplate, tOrder, tSheetName
	Dim tType : tType = HR_Clng(Request("Type"))
	Dim tFieldLen : tFieldLen = 13
	sqlGet = "Select a.* From HR_Class a Where a.ModuleID=1001 And a.ClassType=" & tType
	sqlGet = sqlGet & " Order By a.ClassType ASC, a.RootID ASC, a.OrderID ASC"
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			tSum = 0
			Do While Not rsGet.EOF
				ItemType = HR_Clng(rsGet("ClassType")) : If ItemType > 0 Then ItemTypeName = arrItemType(ItemType - 1)
				tParentID = HR_Clng(rsGet("ParentID")) : ParentClass = ""
				tItemName = Trim(rsGet("ClassName"))
				tTemplate = Trim(rsGet("Template"))
				If tParentID > 0 Then
					ParentClass = GetTypeName("HR_Class", "ClassName", "ClassID", tParentID)
					Set rsTmp = Conn.Execute("Select Max(OrderID) From HR_Class Where ParentID=" & tParentID)
						tOrder = HR_Clng(rsTmp(0))
					Set rsTmp = Nothing
					If HR_Clng(rsGet("OrderID")) = tOrder Then tItemName = "└ " & tItemName Else tItemName = "├ " & tItemName
				End If

				'更新数据模板字段数
				If tTemplate = "TempTableA" Then tFieldLen = 13
				If tTemplate = "TempTableB" Then tFieldLen = 7
				If tTemplate = "TempTableC" Then tFieldLen = 8
				If tTemplate = "TempTableD" Then tFieldLen = 9
				If tTemplate = "TempTableE" Then tFieldLen = 10
				If tTemplate = "TempTableF" Then tFieldLen = 9
				If tTemplate = "TempTableG" Then tFieldLen = 8
				Conn.Execute("Update HR_Class Set FieldLen=" & HR_Clng(tFieldLen) & " Where ClassID=" & rsGet("ClassID"))
				'检查数据表是否存在并统计数据记录总数【无子栏目时】
				tCount = 0
				tSheetName = "HR_Sheet_" & rsGet("ClassID")		'数据表名
				If ChkTable(tSheetName) Then
					Set rsCount = Conn.Execute("Select Count(0) From " & tSheetName & " Where scYear=" & DefYear)
						tCount = HR_CLng(rsCount(0))
					Set rsCount = Nothing
				End If
				tSum = tSum + tCount
				If i > 0 Then tmpData = tmpData & ","
				tmpData = tmpData & "{""ItemID"":""" & rsGet("ClassID") & """,""ItemName"":""" & tItemName & """,""ParentID"":""" & tParentID & """,""ItemType"":""" & HR_Clng(rsGet("ClassType")) & """,""Depth"":""" & HR_Clng(rsGet("Depth")) & """,""GradeCount"":""" & HR_Clng(rsGet("GradeCount")) & """"
				tmpData = tmpData & ",""ItemTypeName"":""" & HR_HTMLEncode(ItemTypeName) & """,""Unit"":""" & Trim(rsGet("Unit")) & """,""StudentType"":""" & Trim(rsGet("StudentType")) & """,""Ratio"":""" & Trim(rsGet("Ratio")) & ""","
				tmpData = tmpData & """Count"":""" & tCount & """,""MaxScore"":""" & Trim(rsGet("MaxScore")) & """,""FieldLen"":""" & Trim(rsGet("FieldLen")) & """"
				tmpData = tmpData & ",""Child"":" & HR_Clng(rsGet("Child")) & ",""Template"":""" & Trim(rsGet("Template")) & """,""Sheet"":""" & Trim(rsGet("SheetName")) & """,""ParentClass"":""" & tParentClass & """,""Intro"":""" & HR_HTMLEncode(nohtml(rsGet("Tips"))) & """}"
				rsGet.MoveNext
				i = i + 1
			Loop
		End If
		vCount = rsGet.Recordcount
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""暂时没有添加项目数据"",""count"":" & vCount & ",""sumtotal"":" & tSum & ",""data"":[" & tmpData
	tmpJson = tmpJson & "],""limit"":""" & HR_Clng(MaxPerPage) & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub

Sub EditBody()
	Dim tmpID : tmpID = HR_Clng(Request("ItemID"))
	Dim tTypeID : tTypeID = HR_Clng(Request("TypeID"))
	Dim isModify : isModify = HR_CBool(Request("Modify"))
	Dim eveAction : eveAction = Trim(ReplaceBadChar(Request("Eve")))
	Dim tmpHtml, rsGet, sqlGet, arrTmp1, j, ItemLevels
	Dim tItemName, tItemType, tParentID, tChild, tTips, tUnit, tSheetName, tFieldID, tStudentType
	Dim tTemplate, tMaxScore, tFieldLen

	SubButTxt = "添加" : ItemLevels = "二级项目"
	sqlGet = "Select * From HR_Class Where ModuleID=1001"
	If eveAction = "addChild" Then
		tParentID = tmpID : tmpID = 0
		isModify = False
	ElseIf eveAction = "edit" Then
		 isModify = True
	End If
	If tmpID > 0 And isModify Then
		sqlGet = sqlGet & " And ClassID=" & tmpID : SubButTxt = "修改"
		Set rsGet = Conn.Execute(sqlGet)
			If rsGet.BOF And rsGet.EOF Then
				tmpHtml = "<div class=""layui-row"" style=""text-align:center;padding:20px 0"">"
				tmpHtml = tmpHtml & "<a class=""layui-btn layui-btn-lg layui-btn-danger""><i class=""layui-icon"">&#xe69c;</i> 您要修改的考核项目【ID：" & tmpID & "】不存在！</a></div>"
				Response.Write tmpHtml
				Exit Sub
			Else
				tItemName = Trim(rsGet("ClassName"))
				tTypeID = HR_Clng(rsGet("ClassType"))
				tParentID = HR_Clng(rsGet("ParentID"))
				tChild = HR_Clng(rsGet("Child"))
				tTips = Trim(rsGet("Tips"))
				tUnit = Trim(rsGet("Unit"))
				tSheetName = Trim(rsGet("SheetName"))
				tTemplate = Trim(rsGet("Template"))
				tFieldID = HR_Clng(rsGet("FieldID"))		'
				tFieldLen = HR_Clng(rsGet("FieldLen"))		'自定义字段长度【与数据模板里匹配】
				tStudentType = Trim(rsGet("StudentType"))
				tMaxScore = HR_CDbl(rsGet("MaxScore"))
			End If
		Set rsGet = Nothing
	End If
	If tParentID = 0 Then ItemLevels = "一级项目"
	If tTypeID > 0 Then tItemType = arrItemType(tTypeID - 1)


	'取级别及等级数
	Dim numRatio
	Set rsGet = Conn.Execute("Select Count(ID) From HR_ItemModel Where ClassID=" & tmpID)
		If rsGet(0) > 0 Then numRatio = "<span class=""layui-badge"">" & rsGet(0) & "</span>"
	Set rsGet = Nothing

	tmpHtml = "<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	tmpHtml = tmpHtml & "	<legend>" & SubButTxt & ItemLevels & "</legend>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layer-hr-box"">" & vbCrlf
	tmpHtml = tmpHtml & "		<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "				<label class=""layui-form-label"">考核类别：</label>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-input-block"">"
	For i = 0 To Ubound(arrItemType)
		If tTypeID = i + 1 Then
			tmpHtml = tmpHtml & "<input type=""radio"" name=""ItemType"" value=""" & i + 1 & """ title=""" & arrItemType(i) & """ checked>"
		Else
			tmpHtml = tmpHtml & "<input type=""radio"" name=""ItemType"" value=""" & i + 1 & """ title=""" & arrItemType(i) & """ disabled="""">"
		End If
	Next
	tmpHtml = tmpHtml & "				</div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf

	tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "				<label class=""layui-form-label"">项目名称：</label>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-input-block""><input type=""text"" name=""ItemName"" id=""ItemName"" value=""" & tItemName & """ placeholder=""项目名称必须填写"" lay-verify=""required"" autocomplete=""on"" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf

	tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-inline""><label class=""layui-form-label"">上级项目：</label>" & vbCrlf
	tmpHtml = tmpHtml & "					<div class=""layui-input-inline"">" & vbCrlf
	If tParentID > 0 Then
		tmpHtml = tmpHtml & "						<select name=""ParentID"" lay-verify=""required"" lay-search="""""
		If isModify Then tmpHtml = tmpHtml & " disabled"
		tmpHtml = tmpHtml & "><option value="""">直接选择或搜索选择</option>"
		tmpHtml = tmpHtml & GetClassOption(1001, tParentID, False)
		tmpHtml = tmpHtml & "</select>" & vbCrlf
	Else
		tmpHtml = tmpHtml & "						<select name=""ParentID"" disabled=""""><option value="""">一级科室不能选择上级</option></select>" & vbCrlf
	End If
	tmpHtml = tmpHtml & "					</div>" & vbCrlf
	tmpHtml = tmpHtml & "				</div>" & vbCrlf
	If tChild = 0 Then
		tmpHtml = tmpHtml & "				<div class=""layui-inline""><label class=""layui-form-label"">计量单位：</label>" & vbCrlf
		tmpHtml = tmpHtml & "					<div class=""layui-input-inline""><select name=""Unit"" lay-verify=""required"">"
		For i = 0 To Ubound(arrUnit)
			tmpHtml = tmpHtml & "<option value=""" & arrUnit(i) & """"
			If tUnit = arrUnit(i) Then tmpHtml = tmpHtml & " selected"
			tmpHtml = tmpHtml & ">" & arrUnit(i) & "</option>"
		Next
		tmpHtml = tmpHtml & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "				</div>" & vbCrlf
	End If
	tmpHtml = tmpHtml & "			</div>" & vbCrlf

	If tChild > 0 Then
		tmpHtml = tmpHtml & "			<div class=""layui-form-item""><div class=""layui-form-mid"" style=""color:#f30""><i class=""hr-icon"">&#xecfd;</i>项目分类不用设置类别、系数、级别等</div></div>" & vbCrlf
	Else
		tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "				<label class=""layui-form-label"">学生类别：</label>" & vbCrlf
		tmpHtml = tmpHtml & "				<div class=""layui-input-block"">" & vbCrlf
		For i = 0 To Ubound(arrStudentType)
			tmpHtml = tmpHtml & "<input type=""checkbox"" name=""StudentType"" value=""" & arrStudentType(i) & """ title=""" & arrStudentType(i) & """ lay-skin=""primary"""
			If HR_IsNull(tStudentType) = False Then
				arrTmp1 = Split(FilterArrNull(tStudentType, ","), ",")
				For j = 0 To Ubound(arrTmp1)
					If Trim(arrTmp1(j)) = arrStudentType(i) Then tmpHtml = tmpHtml & " checked"
				Next
			End if
			tmpHtml = tmpHtml & ">"
		Next
		tmpHtml = tmpHtml & "</div>" & vbCrlf
		tmpHtml = tmpHtml & "			</div>" & vbCrlf

		If isModify Then
			tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
			tmpHtml = tmpHtml & "				<div class=""layui-input-block formBtn""><button type=""button"" class=""layui-btn layui-btn-sm"" id=""btnRatio""><i class=""hr-icon"">&#xf03f;</i>系数</button>"
			tmpHtml = tmpHtml & "<button type=""button"" class=""layui-btn layui-btn-sm"" id=""btnLevel""><i class=""hr-icon"">&#xf34f;</i>级别" & numRatio & "</button></div>" & vbCrlf
			tmpHtml = tmpHtml & "			</div>" & vbCrlf
		Else
			tmpHtml = tmpHtml & "			<div class=""layui-form-item""><div class=""layui-form-mid"" style=""color:#f30""><i class=""hr-icon"">&#xecfd;</i>添加项目时不能设置系数和级别</div></div>" & vbCrlf
		End If

		tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "				<label class=""layui-form-label"">最高分值：</label>" & vbCrlf
		tmpHtml = tmpHtml & "				<div class=""layui-input-inline""><input type=""number"" name=""MaxScore"" value=""" & tMaxScore & """ placeholder=""不限定则为0"" class=""layui-input""></div>" & vbCrlf
		tmpHtml = tmpHtml & "			</div>" & vbCrlf
		tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "				<div class=""layui-inline""><label class=""layui-form-label"">数据模型：</label>" & vbCrlf
		tmpHtml = tmpHtml & "					<div class=""layui-input-inline""><select name=""Template"" id=""Template"" lay-verify=""required"""
		If isModify Then tmpHtml = tmpHtml & " disabled"
		tmpHtml = tmpHtml & ">" & GetTemplateOption(1, tTemplate) & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "				</div>" & vbCrlf
		tmpHtml = tmpHtml & "				<div class=""layui-inline formBtn""><button type=""button"" class=""layui-btn layui-btn-sm"" id=""DataModel"" data-name=""" & tTemplate & """ title=""预览数据模型式样""><i class=""hr-icon"">&#xecfd;</i>预览</button>"
		If isModify Then tmpHtml = tmpHtml & "<button type=""button"" class=""layui-btn layui-btn-sm layui-btn-normal"" id=""btnExcel"" title=""Excel模板文件""><i class=""hr-icon"">&#xedd3;</i></button>" & vbCrlf
		tmpHtml = tmpHtml & "				</div>" & vbCrlf
		tmpHtml = tmpHtml & "			</div>" & vbCrlf
	End If
	tmpHtml = tmpHtml & "			<div class=""layui-form-item"">"
	tmpHtml = tmpHtml & "				<label class=""layui-form-label"">项目说明：</label>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-input-block""><textarea name=""Tips"" id=""Tips"" placeholder=""注释"" lay-verify=""content"" class=""layui-textarea"">" & tTips & "</textarea></div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf

	If tmpID > 0 Then tmpHtml = tmpHtml & "			<input type=""hidden"" name=""ID"" id=""ClassID"" value=""" & tmpID & """><input type=""hidden"" name=""Modify"" value=""True"">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""hr-pop-fix"">" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-inline formBtn""><button class=""layui-btn layui-btn-sm"" lay-submit lay-filter=""SubPost""><i class=""hr-icon"">&#xf0c7;</i>" & SubButTxt & "</button><button type=""reset"" class=""layui-btn layui-btn-sm layui-btn-primary""><i class=""hr-icon"">&#xec75;</i>重置</button>"
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "		</form>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
	Response.Write tmpHtml
End Sub
Sub SaveForm()
	Dim rsSave, tmpJson, tItemName, tParentPath, tDepth, tRootID, tChild, arrChildID, tPrevID, tNextID, tOrderID, tStudentType
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	tItemName = Trim(ReplaceBadChar(Request("ItemName")))
	tStudentType = Trim(ReplaceBadChar(Request("StudentType")))
	tStudentType = FilterArrNull(tStudentType, ",")
	Dim tmpClassID : tmpClassID = GetNewID("HR_Class", "ClassID")
	Dim tParentID : tParentID = HR_Clng(Request("ParentID"))
	Dim tSheetName : tSheetName = "HR_Sheet_" & tmpClassID
	Dim tTempTable : tTempTable = Trim(Request("Template"))
	Dim tFieldLen : tFieldLen = 13
	Dim strField : strField = "序号,工号,教师,学时,日期,,,"
	If tTempTable = "TempTableB" Then tFieldLen = 8
	If tTempTable = "TempTableC" Then tFieldLen = 10

	If tParentID > 0 Then
		Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tParentID)
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				tRootID = rsTmp("RootID")
				tDepth = rsTmp("Depth") + 1
				tParentPath = rsTmp("ParentPath") & "," & tParentID
			End If
		Set rsTmp = Nothing
		Set rsTmp = Conn.Execute("Select Max(OrderID) From HR_Class Where ParentID=" & tParentID)
			tOrderID = HR_Clng(rsTmp(0))
		Set rsTmp = Nothing
		'------ 取同级上一个栏目ID
		Set rsTmp = Conn.Execute("Select top 1 ClassID From HR_Class Where ParentID=" & tParentID & " And OrderID=" & tOrderID)
			If Not(rsTmp.BOF And rsTmp.EOF) Then tPrevID = HR_Clng(rsTmp(0))
		Set rsTmp = Nothing
		tChild = 0 : tOrderID = tOrderID + 1
	Else
		Set rsTmp = Conn.Execute("Select Max(RootID) From HR_Class Where ParentID=0")
			tRootID = HR_Clng(rsTmp(0)) + 1
		Set rsTmp = Nothing
		tDepth = 0 : tParentPath = "0" : tChild = 0 : tPrevID = 0 : tNextID = 0
	End If
	arrChildID = tmpClassID
	SubButTxt = "添加" : If HR_CBool(Request("Modify")) Then SubButTxt = "修改"

	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open("Select * From HR_Class Where ClassID=" & tmpID), Conn, 1, 3
		If rsSave.BOF And rsSave.EOF Then
			rsSave.AddNew
			rsSave("ClassID") = tmpClassID
			rsSave("ClassType") = HR_Clng(Request("ItemType"))
			rsSave("ModuleID") = 1001
			rsSave("MaxPerPage") = 20
			rsSave("ParentID") = tParentID
			rsSave("ParentPath") = tParentPath
			rsSave("Depth") = tDepth
			rsSave("RootID") = tRootID
			rsSave("Child") = 0
			rsSave("arrChildID") = arrChildID
			rsSave("PrevID") = HR_Clng(tPrevID)
			rsSave("NextID") = HR_Clng(tNextID)
			rsSave("OrderID") = HR_Clng(tOrderID)
			rsSave("SheetName") = tSheetName
			rsSave("Template") = tTempTable
			rsSave("FieldLen") = tFieldLen
			rsSave("FieldHead") = strField
		Else
			tSheetName = "HR_Sheet_" & tmpID
			tTempTable = Trim(rsSave("Template"))
			tmpClassID = tmpID
		End If
		rsSave("ClassName") = tItemName
		rsSave("Tips") = Trim(Request("Tips"))
		rsSave("Unit") = Trim(ReplaceBadChar(Request("Unit")))
		rsSave("MaxScore") = HR_CDbl(Request("MaxScore"))
		rsSave("StudentType") = tStudentType
		rsSave.Update
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""项目 " & tItemName & " " & SubButTxt & "成功！"",""ReStr"":""操作成功！""}"
		rsSave.Close
	Set rsSave = Nothing
	
	'更新前一个项目及上级项目
	If tPrevID > 0 And tmpID = 0 Then
		Conn.Execute("Update HR_Class Set NextID=" & tmpClassID & " Where ClassID=" & tPrevID)
	End If
	If tParentID > 0 And tmpID = 0 Then
		Conn.Execute("Update HR_Class Set Child=Child+1,arrChildID=arrChildID+'," & tmpClassID & "' Where ClassID=" & tParentID)
	End If

	'新建表
	If Not(ChkTable(tSheetName)) Then
		sqlTmp = "Select * Into " & tSheetName & " From HR_" & tTempTable & " Where 1=0"
		Conn.Execute(sqlTmp)
		Conn.Execute("Alter Table " & tSheetName & " Add Primary Key (ID)")							'设置主键
		Conn.Execute("Alter Table " & tSheetName & " Add Default(" & tmpClassID & ") For ItemID")		'设置默认值
		Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For ID")		'设置默认值
		Conn.Execute("Alter Table " & tSheetName & " Add Default('') For StudentType")		'设置默认值
		Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For VA0")		'设置默认值
		Conn.Execute("Alter Table " & tSheetName & " Add Default('') For VA1")
		Conn.Execute("Alter Table " & tSheetName & " Add Default('') For VA2")
		Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For VA3")
		Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For Passed")
		Conn.Execute("Alter Table " & tSheetName & " Add Default('') For Explain")
		Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For UserID")
		Conn.Execute("Alter Table " & tSheetName & " Add Default(getdate()) For AppendTime")
		Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For State")
		Conn.Execute("Alter Table " & tSheetName & " Add Default('') For KSMC")
		Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For KSDM")
		Conn.Execute("Alter Table " & tSheetName & " Add Default('') For YGXB")
		Conn.Execute("Alter Table " & tSheetName & " Add Default('') For PRZC")
		Conn.Execute("Delete From " & tSheetName)
	End If
	Call RecordFrontLog(tmpID, SubButTxt & tSheetName & "表记录", SubButTxt & "项目[ID：" & tmpClassID & "]，ItemName：" & Trim(Request("tItemName")), True, "Save")
	Call UpdateKPIField()	'更新KPI字段
	'Call UpdateItemKPI(tmpClassID)	'更新项目KPI

	Response.Write tmpJson
End Sub
Sub Delete()
	Dim tmpJson, rsDel, sqlDel
	Dim tTypeID : tTypeID = HR_Clng(Request("Type"))
	Dim tClassID : tClassID = HR_Clng(Request("ItemID"))
	Dim tParentID, tClassName, tTableName, tTemplate, backup
	Dim tChild, tPrevID, tNextID, tOrderID

	tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""项目 " & tClassName & " 删除失败！"",""ReStr"":""删除失败！""}"
	backup = BackSQLData(0, "")		'备份数据库

	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tClassID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tClassName = rsTmp("ClassName")
			tTableName = Trim(rsTmp("SheetName"))
			tParentID = HR_Clng(rsTmp("ParentID"))
			tChild = HR_Clng(rsTmp("Child"))
			tPrevID = HR_Clng(rsTmp("PrevID"))
			tNextID = HR_Clng(rsTmp("NextID"))
			tOrderID = HR_Clng(rsTmp("OrderID"))
		End If
	Set rsTmp = Nothing
	'含有子栏目不能删除
	If tChild > 0 Then
		tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""项目 " & tClassName & " 含有二级项目，不能删除！"",""ReStr"":""删除失败！""}"
		Response.Write tmpJson : Exit Sub
	End If

	'删除等级，表名：HR_ItemGrade
	Conn.Execute("Delete From HR_ItemGrade Where ClassID=" & tClassID)
	'删除级别，表名：HR_ItemModel
	Conn.Execute("Delete From HR_ItemModel Where ClassID=" & tClassID)

	'删除附件，表名：HR_Attach（同时物理文件）
	Dim delNum
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_Attach Where ClassID=" & tClassID), Conn, 1, 3
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			delNum = 0
			Do While Not rsTmp.EOF
				delNum = delNum + DelUploadFiles(rsTmp("FilePath"))	'删除指定文件（物理删除）
				rsTmp.MoveNext
			Loop
		End If
	Set rsTmp = Nothing
	Conn.Execute("Delete From HR_Attach Where ClassID=" & tClassID)		'删除记录

	'更新上一个目录(OrderID等)
	If tPrevID > 0 Then
		Conn.Execute("Update HR_Class Set NextID=" & tNextID & " Where ClassID=" & tPrevID)
	End If
	'更新下一个目录
	If tNextID > 0 Then
		Conn.Execute("Update HR_Class Set PrevID=" & tPrevID & " Where ClassID=" & tNextID)
	End If
	'更新父目录
	Dim tmpChild, tArrChild, newChild, iArr
	If tParentID > 0 Then
		tmpChild = GetTypeName("HR_Class", "arrChildID", "ClassID", tParentID)
		If tmpChild <> "" Then				'将本栏目从中过滤
			tmpChild = FilterArrNull(tmpChild, ",")
			tArrChild = Split(tmpChild, ",")
			newChild = ""
			For iArr = 0 To Ubound(tArrChild)
				If tClassID <> HR_Clng(tArrChild(iArr)) Then newChild = newChild & tArrChild(iArr) & ","
			Next
			newChild = FilterArrNull(newChild, ",")
		End If
		Conn.Execute("Update HR_Class Set Child=Child-1,arrChildID='" & newChild & "' Where ClassID=" & tParentID)
	End If

	'删除项目
	If ChkTable(tTableName) Then	'删除课程表
		Conn.Execute("DROP table " & tTableName)
	End If
	Conn.Execute("Delete From HR_Class Where ClassID=" & tClassID)		'删除考核项目记录

	Call RecordFrontLog(tClassID, "删除考核项目" & tClassName & "", "删除考核项目[ID：" & tClassID & "]，ItemName：" & Trim(tClassName), True, "Delete")
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""考核项目" & tClassName & "及相关的课程记录、附件、数据库等成功删除！"",""ReStr"":""删除完成！""}"
	Response.Write tmpJson
End Sub

Sub Preview()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim rsShow, strChk, tParent, tItemName, tTypeID, tType, tStudentType, tRatio

	tmpHtml = "<div class=""hr-workZones hr-shrink-x10"">"
	Set rsShow = Conn.Execute("Select * From HR_Class Where ClassID=" & tmpID )
		If rsShow.BOF And rsShow.EOF Then
			ErrMsg = "您要查看的考核项目信息【ID：" & tmpID & "】不存在！"
			Response.Write GetErrBody(2) : Exit Sub
		Else
			tItemName = Trim(rsShow("ClassName"))
			tTypeID = HR_Clng(rsShow("ClassType"))
			If tTypeID > 0 Then tType = arrItemType(tTypeID - 1)
			tStudentType = Trim(rsShow("StudentType"))
			tStudentType = Replace(tStudentType, ",", "　")
			tRatio = Trim(rsShow("Ratio"))
			tRatio = Replace(tRatio, ",", "　")

			tmpHtml = tmpHtml & "<fieldset class=""layui-elem-field layui-field-title""><legend>考核项目 " & tItemName & " 预览</legend>"
			tmpHtml = tmpHtml & "<div class=""layui-form layer-hr-box""><table class=""layui-table"">"
			tmpHtml = tmpHtml & "<colgroup><col width=""120""><col><col width=""120""><col></colgroup>"
			tmpHtml = tmpHtml & "<tbody>"

			tmpHtml = tmpHtml & "<tr><td style=""text-align:right;"">考核项目：</td><td colspan=""3"">" & tItemName & "</td></tr>" & vbCrlf
			tmpHtml = tmpHtml & "<tr><td style=""text-align:right;"">序　　号：</td><td>" & HR_Clng(rsShow("ClassID")) & "</td>"
			tmpHtml = tmpHtml & "<td style=""text-align:right;"">类　　别：</td><td>" & Trim(tType) & "</td></tr>" & vbCrlf
			tmpHtml = tmpHtml & "<tr><td style=""text-align:right;"">上级项目：</td><td>" & tParent & "</td>"
			tmpHtml = tmpHtml & "<td style=""text-align:right;"">计量单位：</td><td>" & Trim(rsShow("Unit")) & "</td></tr>"
			tmpHtml = tmpHtml & "<tr><td style=""text-align:right;"">数据模板：</td><td>" & Trim(rsShow("Template")) & "</td>"
			tmpHtml = tmpHtml & "<td style=""text-align:right;"">数据表名：</td><td>" & Trim(rsShow("SheetName")) & "</td></tr>"
			If tStudentType <> "" Then tmpHtml = tmpHtml & "<tr><td style=""text-align:right;"">学生类别：</td><td colspan=""3"">" & tStudentType & "</td></tr>"
			tmpHtml = tmpHtml & "<tr><td style=""text-align:right;"">考核系数：</td><td colspan=""3"">" & tRatio & "</td></tr>"
			tmpHtml = tmpHtml & "<tr><td style=""text-align:right;"">项目说明：</td><td colspan=""3"">" & Trim(rsShow("Tips")) & "</td></tr>"
			tmpHtml = tmpHtml & "</tbody>"
			tmpHtml = tmpHtml & "</table></div>" & vbCrlf
			tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
		End If
	Set rsShow = Nothing
	Set rsShow = Conn.Execute("Select * From HR_ItemModel Where ClassID=" & tmpID )
		If Not(rsShow.BOF And rsShow.EOF) Then
			tmpHtml = tmpHtml & "<fieldset class=""layui-elem-field layui-field-title""><legend> 级别 </legend>"
			tmpHtml = tmpHtml & "<div class=""layui-form layer-hr-box""><table class=""layui-table"">"
			tmpHtml = tmpHtml & "<thead><tr><th width=""120"">名　称</th><th width=""70"">系数值</th><th width=""70"">单位</th><th>备注</th></tr></thead>"
			tmpHtml = tmpHtml & "<tbody>"
			Do While Not rsShow.EOF
				tmpHtml = tmpHtml & "<tr><td>" & Trim(rsShow("FieldName")) & "</td><td>" & HR_CDbl(rsShow("Formula")) & "</td><td>" & Trim(rsShow("Unit")) & "</td><td>" & Trim(rsShow("Intro")) & "</td></tr>"
				rsShow.MoveNext
			Loop
			tmpHtml = tmpHtml & "</tbody>"
			tmpHtml = tmpHtml & "</table></div>" & vbCrlf
			tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
		End If
	Set rsShow = Nothing
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	strHtml = strHtml & tmpHtml
	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub

Sub EditLevel()		'编辑级别
	Dim rsList, strData, tItemName, tTemplate, FieldNum
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tTypeID : tTypeID = HR_Clng(Request("TypeID"))
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = "HR_" & Trim(rsTmp("Template"))
		Else
			Response.Write "业绩考核项目不存在！"
			Exit Sub
		End If
	Set rsTmp = Nothing

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding: 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-pop-fix {position: absolute;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")

	strHtml = strHtml & "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	strHtml = strHtml & "	<fieldset class=""layui-elem-field layui-field-title""><legend>" & tItemName & " 级别管理</legend></fieldset>" & vbCrlf
	strHtml = strHtml & "	<div class=""layui-form"">" & vbCrlf
	strHtml = strHtml & "	<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "ExamItems/ListLevel.html?ItemID=" & tItemID & "',id:'TableLevel'}"" lay-filter=""TableLevel"">" & vbCrlf
	strHtml = strHtml & "		<thead><tr>" & vbCrlf
	strHtml = strHtml & "			<th lay-data=""{field:'ID',align:'center',unresize:true,width:60}"">序号</th>" & vbCrlf
	strHtml = strHtml & "			<th lay-data=""{field:'FieldName',width:200}"">考核项</th>" & vbCrlf
	strHtml = strHtml & "			<th lay-data=""{field:'Formula',align:'center',width:60}"">系数</th>" & vbCrlf
	strHtml = strHtml & "			<th lay-data=""{field:'Unit',align:'center',unresize:true, width:60}"">单位</th>" & vbCrlf
	strHtml = strHtml & "			<th lay-data=""{field:'Intro'}"">备　注</th>" & vbCrlf
	strHtml = strHtml & "			<th lay-data=""{field:'GradeNum',align:'center',unresize:true, width:70}"">等级数</th>" & vbCrlf
	strHtml = strHtml & "			<th lay-data=""{fixed:'right',align:'center',unresize:true,width:160, toolbar: '#barLevel'}"">操作</th>" & vbCrlf
	strHtml = strHtml & "		</tr></thead>" & vbCrlf
	strHtml = strHtml & "	</table>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/html"" id=""barLevel"">" & vbCrlf
	strHtml = strHtml & "		<div class=""layui-btn-group"">" & vbCrlf
	strHtml = strHtml & "		<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	strHtml = strHtml & "		<a class=""layui-btn layui-btn-sm"" lay-event=""grade"" title=""等级""><i class=""hr-icon"">&#xe992;</i></a>" & vbCrlf
	strHtml = strHtml & "		<a class=""layui-btn layui-btn-sm layui-btn-danger"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	strHtml = strHtml & "		</div>" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHtml = strHtml & "	</div>" & vbCrlf
	strHtml = strHtml & "	<div class=""hr-pop-fix"" id=""ExportBox"">" & vbCrlf
	strHtml = strHtml & "		<input type=""hidden"" name=""Template"" id=""Template"" value=""" & tTemplate & """><input type=""hidden"" name=""ItemID"" id=""ItemID"" value=""" & tItemID & """>" & vbCrlf
	strHtml = strHtml & "		<div class=""Export""><button class=""layui-btn layui-btn-sm"" name=""AddPost"" id=""AddPost""><i class=""layui-icon"">&#xe654;</i>添加级别</button><button class=""layui-btn layui-btn-sm"" id=""refresh"" title=""刷新""><i class=""hr-icon"">&#xf343;</i>刷新</button></div>" & vbCrlf
	strHtml = strHtml & "	</div>" & vbCrlf
	strHtml = strHtml & "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableLevel)"", function(obj){" & vbCrlf
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""edit""){" & vbCrlf
	tmpHtml = tmpHtml & "				$.get(""" & ParmPath & "ExamItems/LevelForm.html"",{ItemID:data.ItemID, ID:data.ID}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "					layer.open({type:1,content:strForm,title:[""编辑级别"",""font-size:16""],area:[""450px"", ""72%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					form.render();" & vbCrlf
	tmpHtml = tmpHtml & "					form.on(""submit(LevelPost)"", function(PostData){" & vbCrlf
	tmpHtml = tmpHtml & "						var loadMsg = layer.load(1,{shade:[0.1, ""#000""]});" & vbCrlf
	tmpHtml = tmpHtml & "						$.post(""" & ParmPath & "ExamItems/SaveLevel.html"", PostData.field, function(result){" & vbCrlf
	tmpHtml = tmpHtml & "							var reData = eval(""("" + result + "")""); layer.close(loadMsg);" & vbCrlf
	tmpHtml = tmpHtml & "							if(reData.Return){" & vbCrlf
	tmpHtml = tmpHtml & "								layer.alert(reData.reMessge, {icon:1},function(layero, index){layer.closeAll();window.location.reload();});" & vbCrlf
	tmpHtml = tmpHtml & "							}else{" & vbCrlf
	tmpHtml = tmpHtml & "								layer.alert(reData.reMessge, {icon:2});" & vbCrlf
	tmpHtml = tmpHtml & "							}" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "						return false;" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""del""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.confirm('真的删除选中的级别项吗？<br />相关的数据将同步删除而且无法恢复！', {icon:3,title: ""删除提醒""}, function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					var loadMsg = layer.load(1,{shade:[0.1, ""#000""]});" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "ExamItems/DelLevel.html?ID="" + data.ID, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.close(loadMsg);" & vbCrlf
	tmpHtml = tmpHtml & "						if(reData.Return){;" & vbCrlf
	tmpHtml = tmpHtml & "							layer.alert(reData.reMessge, {icon:1,title: ""系统提示""},function(layero, index){layer.close(layer.index);table.reload(""TableLevel"");});" & vbCrlf
	tmpHtml = tmpHtml & "						}else{" & vbCrlf
	tmpHtml = tmpHtml & "							layer.alert(reData.reMessge, {icon:2,title: ""系统提示""});" & vbCrlf
	tmpHtml = tmpHtml & "						}" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					layer.close(index);" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""grade""){" & vbCrlf
	tmpHtml = tmpHtml & "				parent.layer.open({type:2,content:""" & ParmPath & "ExamItems/EditGrade.html?ItemID="" + data.ItemID + ""&ID="" + data.ID,title:[""编辑等级"",""font-size:16""],area:[""600px"", ""400px""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#AddPost"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var ItemID = $(""#ItemID"").val();" & vbCrlf
	tmpHtml = tmpHtml & "			$.get(""" & ParmPath & "ExamItems/LevelForm.html"",{ItemID:ItemID, ID:0}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:1,content:strForm,title:[""添加级别"",""font-size:16""],area:[""600px"", ""72%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "				form.render();" & vbCrlf
	tmpHtml = tmpHtml & "				form.on(""submit(LevelPost)"", function(PostData){" & vbCrlf
	tmpHtml = tmpHtml & "					var loadMsg = layer.load(1,{shade:[0.1, ""#000""]});" & vbCrlf
	tmpHtml = tmpHtml & "					$.post(""" & ParmPath & "ExamItems/SaveLevel.html"", PostData.field, function(result){" & vbCrlf
	tmpHtml = tmpHtml & "						var reData = eval(""("" + result + "")"");layer.close(loadMsg);" & vbCrlf
	tmpHtml = tmpHtml & "						if(reData.Return){" & vbCrlf
	tmpHtml = tmpHtml & "							layer.alert(reData.reMessge, {icon:1},function(layero, index){window.location.reload();});" & vbCrlf
	tmpHtml = tmpHtml & "						}else{" & vbCrlf
	tmpHtml = tmpHtml & "							layer.alert(reData.reMessge, {icon:2});" & vbCrlf
	tmpHtml = tmpHtml & "						}" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					return false;" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf

	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#refresh"").on(""click"", function(index){window.location.reload();});" & vbCrlf		'刷新当前页
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = strHtml & getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
End Sub
Sub ListLevel()
	Dim tmpJson, tmpData, rsGet, sqlGet, vCount, vMSG, tIntro
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tGradeNum : tGradeNum = 0

	sqlGet = "Select * From HR_ItemModel Where ClassID=" & tItemID & " Order By ID ASC"
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			Do While Not rsGet.EOF
				Set rsTmp = Conn.Execute("Select Count(ID) From HR_ItemGrade Where LevelID=" & rsGet("ID"))
					tGradeNum = rsTmp(0)
				Set rsTmp = Nothing
				tIntro = nohtml(rsGet("Intro")) : tIntro = Replace(tIntro, Chr(10), "")
				If i > 0 Then tmpData = tmpData & ","
				tmpData = tmpData & "{""ID"":""" & rsGet("ID") & """,""FieldName"":""" & Trim(rsGet("FieldName")) & """,""Formula"":""" & FormatNumber(HR_CDbl(rsGet("Formula")), 1, -1) & """"
				tmpData = tmpData & ",""GradeNum"":" & tGradeNum & ",""Unit"":""" & Trim(rsGet("Unit")) & """,""Intro"":""" & tIntro & """,""ItemID"":" & HR_Clng(rsGet("ClassID")) & "}"
				rsGet.MoveNext
				i = i + 1
			Loop
		End If
		vCount = rsGet.Recordcount
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""暂时没有添加等级"",""count"":" & vCount & ",""data"":[" & tmpData
	tmpJson = tmpJson & "],""limit"":""" & HR_Clng(MaxPerPage) & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub
Sub LevelForm()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpHtml : SubButTxt = "添加"
	Dim tItemName, tFieldName, tNumValue, tUnit, tFormula, tIntro

	tItemName = GetTypeName("HR_Class", "ClassName", "ClassID", tItemID)
	sqlTmp = "Select * From HR_ItemModel Where ID=" & tmpID
	If tmpID > 0 Then
		SubButTxt = "修改"
		Set rsTmp = Conn.Execute(sqlTmp)
			If rsTmp.BOF And rsTmp.EOF Then
				tmpHtml = "<div class=""layui-row"" style=""text-align:center;padding:20px 0""><a class=""layui-btn layui-btn-lg layui-btn-danger""><i class=""layui-icon"">&#xe69c;</i> 您要修改的级别【ID：" & tmpID & "】不存在！</a></div>"
				Response.Write tmpHtml : Exit Sub
			Else
				tItemID = HR_Clng(rsTmp("ClassID"))
				tFieldName = Trim(rsTmp("FieldName"))
				tNumValue = HR_Clng(rsTmp("NumValue"))
				tUnit = Trim(rsTmp("Unit"))
				tFormula = FormatNumber(rsTmp("Formula"), 2, -1)
				tIntro = Trim(rsTmp("Intro"))
			End If
		Set rsTmp = Nothing
	End If

	tmpHtml = "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:10px;"">" & vbCrlf
	tmpHtml = tmpHtml & "	<legend>" & SubButTxt & " " & tItemName & " 级别</legend>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layer-hr-box"">" & vbCrlf
	tmpHtml = tmpHtml & "		<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "				<label class=""layui-form-label"">级　　别：</label>"
	tmpHtml = tmpHtml & "				<div class=""layui-input-block""><input type=""text"" name=""FieldName"" value=""" & tFieldName & """ placeholder=""级别不能为空"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "				<label class=""layui-form-label"">系　　数：</label>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-input-inline""><input type=""text"" name=""Formula"" id=""Formula"" value=""" & tFormula & """ placeholder=""请输入考核系数"" lay-verify=""number"" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-form-item""><label class=""layui-form-label"">计量单位：</label>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-input-inline""><select name=""Unit"" lay-verify=""required"">"
	For i = 0 To Ubound(arrUnit)
		tmpHtml = tmpHtml & "<option value=""" & arrUnit(i) & """"
		If tUnit = arrUnit(i) Then tmpHtml = tmpHtml & " selected"
		tmpHtml = tmpHtml & ">" & arrUnit(i) & "</option>"
	Next
	tmpHtml = tmpHtml & "</select></div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "				<label class=""layui-form-label"">说　　明：</label>" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-input-block""><textarea name=""Intro"" placeholder=""请输入内容"" class=""layui-textarea"">" & tIntro & "</textarea></div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "			<input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""ItemID"" value=""" & tItemID & """><input type=""hidden"" name=""Modify"" value=""True"">" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""hr-pop-fix"">" & vbCrlf
	tmpHtml = tmpHtml & "				<div class=""layui-inline""><button class=""layui-btn layui-btn-sm"" lay-submit lay-filter=""LevelPost"">" & SubButTxt & "</button><button type=""reset"" class=""layui-btn layui-btn-sm layui-btn-primary"">重置</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "		</form>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
	Response.Write tmpHtml
End Sub
Sub SaveLevel()
	Dim tmpJson, tType
	Dim newID : newID = GetNewID("HR_ItemModel", "ID")
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	tType = GetTypeName("HR_Class", "ClassType", "ClassID", tItemID)
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_ItemModel Where ID=" & tmpID), Conn, 1, 3
		If rsTmp.BOF And rsTmp.EOF Then
			rsTmp.AddNew
			rsTmp("ID") = newID
			rsTmp("TypeID") = HR_Clng(tType)
			rsTmp("ModelID") = 0
			rsTmp("ClassID") = tItemID
		Else
			newID = tmpID
		End If
		rsTmp("FieldName") = Trim(Request("FieldName"))
		rsTmp("NumValue") = HR_Clng(Request("NumValue"))
		rsTmp("Unit") = Trim(Request("Unit"))
		rsTmp("Formula") = HR_CDbl(Request("Formula"))
		rsTmp("Intro") = Trim(Request("Intro"))
		rsTmp.Update
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""级别 " & Trim(Request("FieldName")) & " 更新成功！"",""ReStr"":""操作成功！""}"
		rsTmp.Close
	Set rsTmp = Nothing
	Call RecordFrontLog(tItemID, "保存至HR_ItemModel表", "修改级别[ID：" & newID & "]，FieldName：" & Trim(Request("FieldName")), True, "Save")
	Call UpdateItemKPI(tItemID)		'更新项目KPI
	Response.Write tmpJson
End Sub

Sub DelLevel()
	Dim tmpJson, rsDel, sqlDel, strDel, arrDel, iDel, tmpErr, tItemID
	strDel = Trim(ReplaceBadChar(Request("ID")))
	strDel = DelRightComma(strDel)
	arrDel = Split(strDel, ",")
	iDel = 0
	For i = 0 To Ubound(arrDel)
		Set rsDel = Server.CreateObject("ADODB.RecordSet")
			rsDel.Open("Select * From HR_ItemModel Where ID=" & HR_Clng(arrDel(i))), Conn, 1, 3
			If Not(rsDel.BOF And rsDel.EOF) Then
				tItemID = rsDel("ClassID")
				rsDel.Delete
				iDel = iDel + 1
				rsDel.Close
			End If
		Set rsDel = Nothing
	Next
	Call UpdateItemKPI(tItemID)		'更新项目KPI
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & iDel & "/" & Ubound(arrDel) + 1 & " 条记录删除成功！" & tmpErr & """,""ReStr"":""操作成功！""}"
	Call RecordFrontLog(tItemID, "删除HR_ItemModel表记录", "删除级别[ID：" & strDel & "]，FieldName：" & Trim(Request("FieldName")), True, "Save")
	Response.Write tmpJson
End Sub
%>