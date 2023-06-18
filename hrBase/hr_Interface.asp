<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="./incCommon.asp"-->
<%
SiteTitle = "接口管理"
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveEdit" Call SaveEdit()
	Case "jsonList" Call GetJsonList()
	Case "Delete", "Empty" Call Delete()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim layUrl : layUrl = ParmPath & "Interface/jsonList.html"

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.sumbar b {color:#F60;padding:0 2px} .sumbar b.sumDebit{color:#080}" & vbCrlf		'表头汇总
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	tmpHtml = "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)

	tmpHtml = "<a href=""" & ParmPath & "Interface/Index.html"">" & SiteTitle & "</a><a><cite>列表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form name=""SearchForm"" id=""SearchForm"" class=""layui-form soBox"" action="""" method=""get""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input type=""text"" class=""layui-input"" name=""soWord"" value="""" id=""soWord"" placeholder=""搜索申请人"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-btn-normal"" data-event=""reload"" id=""SearchBtn""><i class=""layui-icon"">&#xe615;</i>搜索</button></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn"">" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn hr-btn_deon"" data-event=""batchdel"" title=""批量删除""><i class=""layui-icon"">&#xe640;</i></button>" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn layui-btn-normal"" data-event=""addnew"" title=""新增""><i class=""layui-icon"">&#xe654;</i></button>" & vbCrlf
	Response.Write "			<button type=""button"" class=""layui-btn layui-bg-green"" data-event=""refresh"" title=""刷新""><i class=""hr-icon"">&#xebbb;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf	'行工具
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-xs layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-xs hr-btn_olive"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",limit:300,skin:'line'" & vbCrlf
	tmpHtml = tmpHtml & "			,text:{none:'暂时没有接口'},cols: [[" & vbCrlf				'设置表头
	tmpHtml = tmpHtml & "				{type:'checkbox',fixed:'left',unresize:true,align:'center',width:50}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'ID',title:'序号',sort:true,width:60,align:'center',unresize:true}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'ApiName',title:'接口名称',width:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Key',title:'key',width:220}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'Domain',title:'请求域名',align:'left',minWidth:150}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'ApplyTime',title:'申请时间',width:120}" & vbCrlf
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
	tmpHtml = tmpHtml & "						$.getJSON(""" & ParmPath & "Interface/Delete.html"",{ID:arrID.join()}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "							layer.alert(reData.errmsg,{title:""删除结果"",icon:reData.icon},function(){ table.reload(""layList"");layer.close(layer.index); });" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "					}); break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""addnew"":	" & vbCrlf			'添加
	tmpHtml = tmpHtml & "					layer.open({type:2, content:'" & ParmPath & "Interface/AddNew.html',title:[""添加接口"",""font-size:16""],area:[""650px"",""360px""]});" & vbCrlf
	tmpHtml = tmpHtml & "					break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""refresh"":location.reload(); break;" & vbCrlf
	tmpHtml = tmpHtml & "			};" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf		'监听工具栏
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""del""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.confirm(""您确定要删除该条记录？删除后无法恢复"",{icon:3,title:[""删除警告"",""background-color:#f30""]},function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "Interface/Delete.html"",{ID:data.ID}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.msg(reData.errmsg,{title:""删除结果"",btn:""关闭"",time:0},function(){ obj.del(); });" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""edit""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""EditWin"", content:'" & ParmPath & "Interface/Edit.html?ID='+data.ID,title:[""修改接口"",""font-size:16""],area:[""650px"",""360px""]});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""details""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""viewWin"", content:""" & ParmPath & "Interface/Details.html?ID="" + data.ID,title:[""查看详情"",""font-size:16""],area:[""760px"", ""92%""]});" & vbCrlf
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

	sqlGet = "Select a.* From HR_Interface a Where a.ID>0"
	If HR_IsNull(soWord) = False Then sqlGet = sqlGet & " And a.ApiName like '%" & soWord & "%'"
	sqlGet = sqlGet & " Order By a.ApplyTime DESC"

	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			i = 0 : TotalPut = rsGet.Recordcount
			Do While Not rsGet.EOF
				If i > 0 Then tmpJson = tmpJson & ","
				tmpJson = tmpJson & "{""ID"":" & rsGet("ID") & ",""ApiName"":""" & Trim(rsGet("ApiName")) & """,""Key"":""" & Trim(rsGet("ApiKey")) & """,""Domain"":""" & Trim(rsGet("Domain")) & """,""ApplyTime"":""" & FormatDate(rsGet("ApplyTime"), 10) & """"
				tmpJson = tmpJson & "}"
				rsGet.MoveNext
				i = i + 1
			Loop
		End If
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""查询成功"",""count"":" & HR_Clng(TotalPut) & ",""data"":[" & tmpJson & "],""limit"":""" & MaxPerPage & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub

Sub EditBody()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim apiKey : apiKey = GetRndPassword(12)		'取随机数12位
	Dim isModify : isModify = HR_CBool(Request("Modify"))
	Dim tApiName, tApiKey, tDomain

	sqlTmp = "Select * From HR_Interface Where ID=" & tmpID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			isModify = True
			tApiName = Trim(rsTmp("ApiName"))
			tApiKey = Trim(rsTmp("ApiKey"))
			tDomain = Trim(rsTmp("Domain"))
		End If
	Set rsTmp = Nothing

	If HR_IsNull(tApiKey) Then tApiKey = apiKey

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .morebtn {padding:3px 0!important;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .tips {padding-left:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<form class=""layui-form"" id=""EditForm"" name=""EditForm"" lay-filter=""EditForm"" action="""">" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">接口名称：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><input type=""text"" name=""ApiName"" id=""ApiName"" lay-verify=""required"" value=""" & tApiName & """ class=""layui-input txt1""></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">密钥：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-inline""><input type=""text"" name=""ApiKey"" id=""ApiKey"" lay-verify=""required"" placeholder=""请输入12位字符"" value=""" & tApiKey & """ class=""layui-input txt1""></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""layui-form-item"">" & vbCrlf
	Response.Write "			<label class=""layui-form-label"">请求域名：</label>" & vbCrlf
	Response.Write "			<div class=""layui-input-block""><input type=""text"" name=""Domain"" id=""Domain"" lay-verify=""required"" placeholder=""如：http://****"" value=""" & tDomain & """ class=""layui-input txt1""></div>" & vbCrlf
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
	tmpHtml = tmpHtml & "	layui.use([""form"", ""table"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table=layui.table, element=layui.element, form=layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#EditPost"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			$.post(""" & ParmPath & "Interface/SaveEdit.html"",$(""#EditForm"").serialize(), function(formResult){" & vbCrlf
	tmpHtml = tmpHtml & "				var reData = eval(""("" + formResult + "")"");" & vbCrlf
	tmpHtml = tmpHtml & "				layer.alert(reData.errmsg,{icon:reData.icon},function(){" & vbCrlf
	tmpHtml = tmpHtml & "					if(reData.err){ layer.close(layer.index); }else{" & vbCrlf
	tmpHtml = tmpHtml & "						var index1 = parent.layer.getFrameIndex(window.name);" & vbCrlf
	tmpHtml = tmpHtml & "						parent.layui.table.reload(""layList"");" & vbCrlf		'重构列表
	tmpHtml = tmpHtml & "						parent.layer.close(index1);" & vbCrlf					'关闭[在iframe页面]
	tmpHtml = tmpHtml & "						return false;" & vbCrlf
	tmpHtml = tmpHtml & "					}" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = Replace(getPageFoot(1), "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub SaveEdit()
	Dim tmpJson, rsSave, sqlSave : ErrMsg = ""
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tApiName : tApiName = Trim(Request("ApiName"))
	Dim tApiKey : tApiKey = Trim(Request("ApiKey"))
	Dim tDomain : tDomain = Trim(ReplaceBadUrl(Request("Domain")))
	Dim isModify : isModify = HR_CBool(Request("Modify"))

	sqlSave = "Select * From HR_Interface Where ID=" & tmpID
	'------ 判断标题等是否为空 ------
	If Len(tApiName) < 3 Then ErrMsg = "名称太短，至少要输入3个字符！"
	If HR_IsNull(tApiKey) Then ErrMsg = "密钥不能为空"
	If HR_IsNull(tDomain) Then ErrMsg = "请求域名不能为空"
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""err"":true,""icon"":2,""errcode"":500,""errmsg"":""" & ErrMsg & """}" : Exit Sub

	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open sqlSave, Conn, 1, 3
		If rsSave.BOF And rsSave.EOF Then
			rsSave.AddNew
			tmpID = GetNewID("HR_Interface", "ID")
			rsSave("ID") = tmpID
			rsSave("ApplyTime") = Now()
		End If
		rsSave("ApiName") = tApiName
		rsSave("ApiKey") = tApiKey
		rsSave("Domain") = tDomain
		rsSave.Update
		tmpJson = "{""err"":false,""icon"":1,""errcode"":0,""errmsg"":""接口保存成功！""}"
	Set rsSave = Nothing
	Response.Write tmpJson
End Sub

Sub Delete()
	If Action = "Empty" Then
		Conn.Execute("Delete From HR_Interface")
		Response.Write "{""err"":true,""errcode"":0,""errmsg"":""The data emptied"",""icon"":1}" : Exit Sub
	End If
	Dim tmpJson, arrTmpID, iCountID, tmpID : tmpID = Trim(ReplaceBadChar(Request("ID")))
	tmpID = FilterArrNull(tmpID, ",")
	If HR_IsNull(tmpID) = False Then arrTmpID = Split(tmpID, ",") : iCountID = Ubound(arrTmpID) + 1
	If HR_IsNull(tmpID) Then
		tmpJson = "{""err"":false,""errcode"":500,""errmsg"":""未指定删除的数据！"",""icon"":2}"
	Else
		Conn.Execute("Delete From HR_Interface Where ID in(" & tmpID & ")")
		tmpJson = "{""err"":true,""errcode"":0,""errmsg"":""共有 " & HR_CLng(iCountID) & " 条记录删除成功！"",""icon"":1}"
	End If
	Response.Write tmpJson
End Sub
%>