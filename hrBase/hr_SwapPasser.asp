<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
Dim SubButTxt : SiteTitle = "调换课审核员管理"
Dim SwapPasser : SwapPasser = Split(XmlText("Common", "SwapPass", "0"), "|")

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "jsonList" Call GetJsonList()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveForm" Call SaveForm()
	Case "Delete", "Empty" Call Delete()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim layUrl : layUrl = ParmPath & "SwapPasser/jsonList.html"

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)

	tmpHtml = "<a href=""" & ParmPath & "SwapPasser/Index.html"">" & SiteTitle & "</a><a><cite>列表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form soBox"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn"">" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-bg-cyan"" data-event=""addnew"" title=""新增调换课审核员""><i class=""hr-icon"">&#xe7fe;</i>新增调换课审核员</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_fuch"" data-event=""refresh"" id=""refresh"" title=""刷新本页""><i class=""hr-icon"">&#xebbb;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<table class=""layui-table"" id=""TableList"" lay-filter=""TableList""></table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm hr-btn_olive"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""layer"", ""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element, form = layui.form, layer = layui.layer;" & vbCrlf
	tmpHtml = tmpHtml & "		var tableIns = table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem:""#TableList"",id:""layList"",limit:0, skin:""line"", text:{none:'暂时没有审核员'}, cols:[[" & vbCrlf
	tmpHtml = tmpHtml & "				{field:'UserID',title:'序号',sort:true,width:60,align:'center',unresize:true}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'YGXM',title:'姓名',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'YGDM',title:'工号',width:100}" & vbCrlf
	tmpHtml = tmpHtml & "				,{field:'PassAuth',title:'审核权限',minWidth:170}" & vbCrlf
	tmpHtml = tmpHtml & "				,{fixed:'right',title:'操作',align:'center',unresize:true,width:150, toolbar:'#barTable'}" & vbCrlf
	tmpHtml = tmpHtml & "			]],url:""" & layUrl & """" & vbCrlf		'设置异步接口
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		$("".searchBtn button"").on(""click"", function(){" & vbCrlf	'搜索等按钮click事件
	tmpHtml = tmpHtml & "			var btnEvent = $(this).data(""event"");" & vbCrlf
	tmpHtml = tmpHtml & "			switch(btnEvent){" & vbCrlf
	tmpHtml = tmpHtml & "				case ""addnew"":" & vbCrlf
	tmpHtml = tmpHtml & "					layer.open({type:2, content:'" & ParmPath & "SwapPasser/AddNew.html',title:""添加审核员"",area:[""650px"",""400px""]}); break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""refresh"":location.reload(); break;" & vbCrlf
	tmpHtml = tmpHtml & "			};" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf		'监听工具栏
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""del""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.confirm(""您确定要删除该审核员？"",{icon:3,title:[""删除警告"",""background-color:#f30""]},function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "SwapPasser/Delete.html"",{ID:data.UserID}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.alert(reData.errmsg,{title:""删除结果"",icon:reData.icon, btn:""关闭"", time:0},function(){" & vbCrlf
	tmpHtml = tmpHtml & "							if(!reData.err){obj.del();table.reload(""layList"");}layer.close(layer.index);" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""edit""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""EditWin"", content:'" & ParmPath & "SwapPasser/Edit.html?ID='+data.UserID,title:[""编辑调换课审核员""],area:[""650px"", ""380px""]});" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub GetJsonList()
	Dim tmpJson, tmpData, rsGet, sqlGet, tCount
	Dim tPasser, tDepart, tPassAuth
	sqlGet = "Select a.* From HR_User a Where a.SwapPass>0"
	sqlGet = sqlGet & " Order By a.YGDM ASC"
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			i = 0
			Do While Not rsGet.EOF
				tPassAuth = ""
				If rsGet("SwapPass") > 0 Then tPassAuth = SwapPasser(rsGet("SwapPass") - 1)
				If i > 0 Then tmpData = tmpData & ","
				tmpData = tmpData & "{""UserID"":""" & rsGet("UserID") & """,""YGXM"":""" & Trim(rsGet("YGXM")) & """,""YGDM"":""" & HR_Clng(rsGet("YGDM")) & """,""PassAuth"":""" & tPassAuth & """"
				tmpData = tmpData & ",""ManageRank"":" & HR_Clng(rsGet("ManageRank")) & "}"
				rsGet.MoveNext
				i = i + 1
			Loop
		End If
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""暂时没有添加审核员"",""count"":" & HR_Clng(TotalPut) & ",""data"":[" & tmpData & "],""limit"":""" & HR_Clng(MaxPerPage) & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub

Sub EditBody()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim isModify : isModify = False
	Dim tYGXM, tYGDM, tSwapPass
	sqlTmp = "Select a.* From HR_User a Where a.SwapPass>0 And a.UserID=" & tmpID
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tYGXM = rsTmp("YGXM")
			tYGDM = rsTmp("YGDM")
			tSwapPass = HR_CLng(rsTmp("SwapPass"))
		End If
	Set rsTmp = Nothing

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.tips {padding-left:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", vbCrlf)
	Response.Write ReplaceCommonLabel(strHtml)

	tmpHtml = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	tmpHtml = tmpHtml & "<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">选择教师：</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><input type=""text"" name=""ygxm"" id=""ygxm"" value=""" & tYGXM & """ lay-verify=""required"" class=""layui-input"" readonly></div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><span class=""layui-btn layui-btn-sm getBtn"" data-code=""ygdm"" data-name=""ygxm"">查找</span><span class=""tips"">请输入关键字搜索教师，必填项</span></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">工　号：</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><input type=""text"" name=""ygdm"" id=""ygdm"" lay-verify=""number"" value=""" & tYGDM & """ class=""layui-input"" readonly></div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<label class=""layui-form-label"">审核权限：</label>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-input-block"">" & vbCrlf
	Dim tChecked
	For i = 0 To Ubound(SwapPasser)
		tChecked = ""
		If tSwapPass = i + 1 Then tChecked = " checked"
		tmpHtml = tmpHtml & "		<input type=""radio"" name=""StudentType"" value=""" & i + 1 & """ title=""" & SwapPasser(i) & """ lay-skin=""primary""" & tChecked & ">" & vbCrlf
	Next
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf

	If Action = "Edit" And tmpID > 0 Then tmpHtml = tmpHtml & "	<input type=""hidden"" name=""UserID"" id=""UserID"" value=""" & tmpID & """><input type=""hidden"" name=""Modify"" value=""True"">"
	tmpHtml = tmpHtml & "	<div class=""hr-pop-fix"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""hr-grids hr-btn-group"">" & vbCrlf
	tmpHtml = tmpHtml & "			<em><button class=""layui-btn hr-btn_fuch"" type=""button"" lay-submit lay-filter=""EditPost"" id=""EditPost"" title=""保存""><i class=""hr-icon"">&#xf0c7;</i>保存</button></em>" & vbCrlf
	tmpHtml = tmpHtml & "			<em><button class=""layui-btn layui-btn-primary"" type=""reset"" name=""reset"" id=""refresh"" title=""重置""><i class=""hr-icon"">&#xf343;</i>重置</button></em>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf

	tmpHtml = tmpHtml & "</form>"
	tmpHtml = tmpHtml & "</div>"
	Response.Write tmpHtml
	Response.Write "<div class=""hr-place-h50""></div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""layer"", ""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, element = layui.element, form = layui.form, layer = layui.layer;" & vbCrlf

	tmpHtml = tmpHtml & "		$("".getBtn"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var elcode=$(this).data(""code""), elname=$(this).data(""name"");" & vbCrlf		'返回员工代码及名称时的对象
	tmpHtml = tmpHtml & "			var openurl=""" & InstallDir & "Desktop/Contacts/Float.html?Type=0"";" & vbCrlf
	tmpHtml = tmpHtml & "			parent.layer.open({type:2, id:""SelectWin"",content:openurl, title:[""查找教师"",""font-size:16""],area:[""500px"", ""80%""],scrollbar:false,success:function(layero, index){" & vbCrlf
	tmpHtml = tmpHtml & "					var objIframe = $(layero).find('iframe')[0].contentWindow.document.body;" & vbCrlf
	tmpHtml = tmpHtml & "					var obj1 = $(objIframe).contents().find(""#listGroup"");" & vbCrlf
	tmpHtml = tmpHtml & "					obj1.attr(""value"",window.name);obj1.attr(""code"", elcode); obj1.attr(""name"", elname);" & vbCrlf
	tmpHtml = tmpHtml & "				}" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		form.on(""submit(EditPost)"", function(data){" & vbCrlf
	tmpHtml = tmpHtml & "			$.post(""" & ParmPath & "SwapPasser/SaveForm.html"", $(""#EditForm"").serialize(), function(reform){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.alert(reform.errmsg,{icon:reform.icon,time:0,btn:""关闭""},function(){" & vbCrlf
	tmpHtml = tmpHtml & "					if(!reform.err){var index1=parent.layer.getFrameIndex(window.name); parent.layui.table.reload(""layList""); parent.layer.close(index1); }else{ layer.close(layer.index) }" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			},""json"");" & vbCrlf
	tmpHtml = tmpHtml & "			return false;" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub SaveForm()
	Dim tYGXM : tYGXM = Trim(Request("ygxm"))
	Dim tYGDM : tYGDM = HR_CLng(Request("ygdm"))
	Dim tTeachAuth : tTeachAuth = HR_CLng(Request("StudentType"))
	ErrMsg = ""
	If tTeachAuth = 0 Then ErrMsg = "请选择审核权限"
	If tYGDM = 0 Then ErrMsg = "请选择教师"
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""err"":true,""errcode"":500,""errmsg"":""" & ErrMsg & """,""icon"":2}" : Exit Sub

	Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open("Select * From HR_User Where YGDM=" & tYGDM), Conn, 1, 3
		If rs.BOF And rs.EOF Then
			rs.AddNew
			rs("UserID") = GetNewID("HR_User", "UserID")
			rs("ManageRank") = 0
			rs("RegTime") = Now()
			rs("Passed") = True
			rs("Locked") = False
			rs("YGXM") = tYGXM
			rs("YGDM") = tYGDM
		End If
		rs("SwapPass") = tTeachAuth
		rs.Update
	Set rs = Nothing
	ErrMsg = "{""err"":false,""errcode"":0,""errmsg"":""审核员保存成功"",""icon"":1}"
	Response.Write ErrMsg
End Sub

Sub Delete()
	If Action = "Empty" Then
		Conn.Execute("Update HR_User Set SwapPass=0")
		Response.Write "{""err"":true,""errcode"":0,""errmsg"":""The data emptied"",""icon"":1}" : Exit Sub
	End If
	Dim tmpJson, arrTmpID, iCountID, tmpID : tmpID = Trim(ReplaceBadChar(Request("ID")))
	tmpID = FilterArrNull(tmpID, ",")
	If HR_IsNull(tmpID) = False Then arrTmpID = Split(tmpID, ",") : iCountID = Ubound(arrTmpID) + 1
	If HR_IsNull(tmpID) Then
		tmpJson = "{""err"":false,""errcode"":500,""errmsg"":""未指定删除的数据！"",""icon"":2}"
	Else
		Conn.Execute("Update HR_User Set SwapPass=0 Where UserID in(" & tmpID & ")")
		tmpJson = "{""err"":true,""errcode"":0,""errmsg"":""共有 " & HR_CLng(iCountID) & " 条记录删除成功！"",""icon"":1}"
	End If
	Response.Write tmpJson
End Sub
%>