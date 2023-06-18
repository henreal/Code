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
Dim SubButTxt : SiteTitle = "通知"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveForm" Call SaveForm()
	Case "jsonList" Call GetJsonList()
	Case "Delete" Call Delete()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim layUrl : layUrl = ParmPath & "Notice/jsonList.html"

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	'tmpHtml = tmpHtml & "		.mediaPdf iframe {border:0;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	'tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	'tmpHtml = tmpHtml & "		layui.use([""layer""], function(){ layer.load(1); });" & vbCrlf
	'tmpHtml = tmpHtml & "	</script>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	tmpHtml = "<a href=""" & ParmPath & "Notice/Index.html"">" & SiteTitle & "</a><a><cite>列表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form soBox""><div class=""layui-inline"">搜索：</div><div class=""layui-inline""><input class=""layui-input"" name=""SearchWord"" id=""SearchWord"" placeholder=""搜索标题"" autocomplete=""off"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn""><button class=""layui-btn layui-btn-normal"" data-type=""reload"" id=""SearchBtn""><i class=""layui-icon"">&#xe615;</i>搜索</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_deon"" data-type=""delete"" id=""BatchDel"" title=""批量删除""><i class=""layui-icon"">&#xe640;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-bg-green"" data-type=""refresh"" id=""refresh"" title=""刷新""><i class=""hr-icon"">&#xebbb;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-normal"" data-type=""addNew"" id=""addNew"" title=""新增""><i class=""layui-icon"">&#xe654;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_peru"" data-type=""sendmsg"" id=""sendmsg"" title=""发送消息""><i class=""hr-icon"">&#xe9b4;</i>发送消息</button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm hr-btn_peru"" lay-event=""details"" title=""查看详情""><i class=""hr-icon"">&#xefb9;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm hr-btn_olive"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table, element = layui.element;" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "		var tableIns = table.render({" & vbCrlf
	strHtml = strHtml & "			elem:""#TableList"",id:""layList"",height:'full-125',page:true,limit:30,skin:'line',limits:[10,15,20,30,50,100,200]" & vbCrlf
	strHtml = strHtml & "			,text:{none:'暂时没有通知'},cols: [[" & vbCrlf				'设置表头
	strHtml = strHtml & "				{type:'checkbox',unresize:true,align:'center',width:50}" & vbCrlf
	strHtml = strHtml & "				,{field:'ID',title:'序号',sort:true,width:60,align:'center',unresize:true}" & vbCrlf
	strHtml = strHtml & "				,{field:'Title',title:'标题',width:250}" & vbCrlf
	strHtml = strHtml & "				,{field:'Intro',title:'内容',minWidth:300,event:'viewIntro',style:'cursor: pointer;'}" & vbCrlf
	strHtml = strHtml & "				,{field:'PublishesTime',title:'发布时间',align:'center',width:150}" & vbCrlf
	strHtml = strHtml & "				,{title:'操作',align:'center',unresize:true,width:150, toolbar:'#barTable'}" & vbCrlf
	strHtml = strHtml & "			]],url:""" & layUrl & """" & vbCrlf		'设置异步接口
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#addNew"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:2, id:""popBody"",content:""" & ParmPath & "Notice/AddNew.html"",title:[""添加通知"",""font-size:16""],area:[""860px"", ""560px""],maxmin:true });" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#sendmsg"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			location.href = """ & ParmPath & "Message/AddNew.html"";" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	strHtml = strHtml & "			var data = obj.data;" & vbCrlf
	strHtml = strHtml & "			if(obj.event === ""edit""){" & vbCrlf
	strHtml = strHtml & "				layer.open({type:2, id:""popBody"",content:""" & ParmPath & "Notice/Edit.html?ID="" + data.ID,title:[""修改通知"",""font-size:16""],area:[""860px"", ""560px""],maxmin:true });" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""details""){" & vbCrlf
	strHtml = strHtml & "				layer.open({type:2, id:""viewWin"", content:""" & InstallDir & "Desktop/Notice/Details.html?ID="" + data.ID, title:false, area:[""640px"", ""82%""] });" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""del""){" & vbCrlf
	strHtml = strHtml & "				layer.confirm(""您确定要删除该通知？删除后无法恢复"",{icon: 3},function(index){" & vbCrlf
	strHtml = strHtml & "					$.getJSON(""" & ParmPath & "Notice/Delete.html"",{ID:data.ID}, function(reJson){" & vbCrlf
	strHtml = strHtml & "						layer.msg(reJson.reMessge,{title:""删除结果"",btn:""关闭"",time:0},function(){ obj.del(); });" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#BatchDel"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var chkObj = table.checkStatus(""layList"").data, arrID=[];" & vbCrlf
	strHtml = strHtml & "			if(chkObj.length==0){layer.tips(""请选择您要删除的通知！"","".laytable-cell-checkbox"",{tips: [1, ""#F60""]});return false;}" & vbCrlf
	strHtml = strHtml & "			for(var i=0;i<chkObj.length;i++){ arrID.push(chkObj[i].ID); }" & vbCrlf
	strHtml = strHtml & "			layer.confirm(""确认要删除选中的“"" + chkObj.length + ""”条数据？<br />删除后将无法恢复。"",{icon:3, title:""删除警告""},function(index){" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "Notice/Delete.html"",{ID:arrID.join()}, function(reJson){" & vbCrlf
	strHtml = strHtml & "					layer.msg(reJson.reMessge,{title:""删除结果"",btn:""关闭"",time:0},function(){ table.reload(""layList""); });" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub

Sub GetJsonList()
	Dim tmpJson, rsGet, sqlGet
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))

	sqlGet = "Select * From HR_Notice Where ID>0"
	sqlGet = sqlGet & " Order By PublishesTime DESC"
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
			Dim tIntro
			Do While Not rsGet.EOF
				tIntro = nohtml(rsGet("Content")) : tIntro = Replace(tIntro, chr(10), "") : tIntro = Replace(nohtml(tIntro), chr(13), "") : tIntro = GetSubStr(tIntro, 110, True)
				If i > 0 Then tmpJson = tmpJson & ","
				tmpJson = tmpJson & "{""ID"":" & rsGet("ID") & ",""Title"":""" & rsGet("Title") & """,""Intro"":""" & tIntro & """"
				tmpJson = tmpJson & ",""KeyWord"":""" & Trim(rsGet("KeyWord")) & """,""Hits"":" & HR_Clng(rsGet("Hits")) & ",""PublishesTime"":""" & FormatDate(rsGet("PublishesTime"), 1) & """}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""查询成功"",""count"":" & HR_Clng(TotalPut) & ",""data"":[" & tmpJson & "],""limit"":""" & MaxPerPage & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub

Sub EditBody()
	SubButTxt = "添加"
	Dim isModify : isModify = HR_CBool(Request("Modify"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	If Action = "Edit" And tmpID > 0 Then isModify = True

	Dim tTitle, tContent
	If isModify Then
		Set rsTmp = Conn.Execute("Select * From HR_Notice Where ID=" & tmpID )
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				tTitle = rsTmp("Title")
				tContent = Trim(rsTmp("Content"))
			End If
		Set rsTmp = Nothing
		SubButTxt = "修改"
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-pop-fix {z-index:1001;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	Response.Write "<legend>" & SubButTxt & "通知公告</legend>"
	Response.Write "<div class=""hr-shrink-x10""></div>"
	Response.Write "<form class=""layui-form"" id=""FloatForm"" name=""FloatForm"" lay-filter=""FloatForm"" action="""">" & vbCrlf
	Response.Write "	<div class=""layui-form-item"">" & vbCrlf
	Response.Write "		<label class=""layui-form-label"">标　题：</label>" & vbCrlf
	Response.Write "		<div class=""layui-input-block""><input type=""text"" name=""title"" id=""title"" lay-verify=""required"" value=""" & tTitle & """ class=""layui-input txt1""></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""layui-form-item"">" & vbCrlf
	Response.Write "		<label class=""layui-form-label"">内　容：</label>" & vbCrlf
	Response.Write "		<div class=""layui-input-block""><script type=""text/plain"" name=""Content"" id=""content"" style=""width:100%;height:220px;"">" & tContent & "</script></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-pop-fix"">" & vbCrlf
	Response.Write "		<div class=""formBtn""><button class=""layui-btn layui-btn-sm layui-bg-cyan"" lay-filter=""FloatPost"" lay-submit id=""FloatPost"">" & SubButTxt & "</button><button type=""reset"" class=""layui-btn layui-btn-sm hr-btn_peru"">重置</button></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	If tmpID > 0 Then Response.Write "<input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""Modify"" value=""True"">"
	Response.Write "</form>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	
	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "UEditor/ueditor.config.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"" src=""" & InstallDir & "UEditor/ueditor.all.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	var editor1 = UE.getEditor(""content"",{" & vbCrlf
	strHtml = strHtml & "		toolbars: [['FullScreen', 'Source', '|', 'fontfamily', 'fontsize', 'bold', 'italic', 'forecolor', 'backcolor', 'underline', 'strikethrough', '|', 'justifyleft', 'justifyright', 'justifycenter', 'indent', 'horizontal', 'insertorderedlist', 'insertunorderedlist', '|', 'emotion', 'spechars', 'simpleupload', 'link', 'test']]" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table, form = layui.form, element = layui.element;" & vbCrlf
	strHtml = strHtml & "		form.on(""submit(FloatPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "			editor1.sync();" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Notice/SaveForm.html"", PostData.field, function(result){" & vbCrlf
	strHtml = strHtml & "				var reData = eval(""("" + result + "")""), icon=0;if(reData.Return){icon=1};" & vbCrlf
	strHtml = strHtml & "				layer.alert(reData.reMessge,{icon:icon,time:0,btn:""关闭""},function(){" & vbCrlf
	strHtml = strHtml & "					if(reData.Err==0){" & vbCrlf
	strHtml = strHtml & "						var index1 = parent.layer.getFrameIndex(window.name);" & vbCrlf
	strHtml = strHtml & "						parent.layui.table.reload(""layList"");" & vbCrlf
	strHtml = strHtml & "						parent.layer.close(index1);" & vbCrlf
	strHtml = strHtml & "					}else{" & vbCrlf
	strHtml = strHtml & "						layer.close(layer.index);" & vbCrlf
	strHtml = strHtml & "					}" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub
Sub SaveForm()
	Dim tmpJson, rsSave
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tTitle : tTitle = Trim(Request("title"))
	Dim tContent : tContent = Trim(Request("Content"))
	If HR_IsNull(tTitle) Then Response.Write "{""Return"":false,""Err"":500,""reMessge"":""通知标题不能为空！""}" : Exit Sub
	If HR_IsNull(tContent) Then Response.Write "{""Return"":false,""Err"":500,""reMessge"":""通知内容没有填写！""}" : Exit Sub

	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open("Select * From HR_Notice Where ID=" & tmpID), Conn, 1, 3
		If rsSave.BOF And rsSave.EOF Then
			rsSave.AddNew
			rsSave("ID") = GetNewID("HR_Notice", "ID")
			rsSave("PublishesTime") = Now()
			rsSave("Hits") = 0
		End If
		rsSave("Title") = tTitle
		rsSave("Content") = Trim(tContent)
		rsSave("KeyWord") = Trim(Request("KeyWord"))
		rsSave.Update
	Set rsSave = Nothing
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""通知修改成功！""}"
	Response.Write tmpJson
End Sub

Sub Delete()
	Dim tmpJson, arrTmpID, tmpID : tmpID = Trim(ReplaceBadChar(Request("ID")))
	tmpID = FilterArrNull(tmpID, ",")
	If HR_IsNull(tmpID) Then
		tmpJson = "{""Return"":false,""Err"":400,""reMessge"":""未选择删除数据！""}"
	Else
		arrTmpID = Split(tmpID, ",")
		Conn.Execute("Delete From HR_Notice Where ID in(" & tmpID & ")")
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & HR_Clng(Ubound(arrTmpID) + 1) & " 条数据删除成功！""}"
	End If
	Response.Write tmpJson
End Sub
%>