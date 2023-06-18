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
Dim arrItemType : arrItemType = Split(XmlText("Common", "ItemType", ""), "|")
Dim SubButTxt : SiteTitle = "数据模型"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

If UserRank < 0 Then
	ErrMsg = "您没有 " & SiteTitle & " 的管理权限！"
	Response.Write GetErrBody(0) : Response.End
End If

Select Case Action
	Case "Index" Call MainBody()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveForm" Call SaveForm()
	Case "jsonList" Call GetJsonList()

	Case "Config" Call Config()
	Case "SaveConfig" Call SaveConfig()
	Case "ChkEduYear" Call ChkEduYear()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	Dim layUrl : layUrl = ParmPath & "DataModel/jsonList.html"

	tmpHtml = "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.setup {background:#754;}" & vbCrlf
	tmpHtml = tmpHtml & "		.soBox .searchBtn {vertical-align:top} .soBox .layui-inline {margin-bottom:8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.soBox .layui-form-select dl {top: 31px;} .soBox .layui-input {height: 30px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.soBox .layui-btn {height: 30px;line-height: 30px;padding: 0 12px;} .soBox .layui-form-select dl dd {padding: 0 5px;line-height: 30px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")

	tmpHtml = "<a href=""" & ParmPath & "DataModel/Index.html"">" & SiteTitle & "</a><a><cite>模型列表</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form soBox""><div class=""layui-inline"">筛选：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input class=""layui-input"" name=""soWord"" value=""" & soWord & """ id=""soWord"" placeholder=""模型名称"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn""><button class=""layui-btn layui-btn-normal"" data-type=""reload"" id=""SearchBtn""><i class=""hr-icon"">&#xeba1;</i>搜索</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-bg-cyan addNew"" data-type=""2"" title=""新增模型""><i class=""hr-icon"">&#xee41;</i>新增模型</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn setup"" data-type=""2"" title=""基本参数""><i class=""hr-icon"">&#xee39;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn dict"" data-type=""dictdata"" title=""数据字典""><i class=""hr-icon"">&#xec10;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<table class=""layui-table"" lay-data=""{url:'" & layUrl & "',text:{none:'没有数据模型'},id:'TableList'}"" lay-filter=""TableList"" lay-skin=""nob_1"">"
	Response.Write "		<thead><tr>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'ModelID',unresize:true, align:'center',width:60}"">序号</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'ModelName',width:200}"">模型名称</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'Type',width:130}"">类　　别</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'FieldsLen', width:60}"">字段数</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'CountField',unresize:true,align:'center',width:70}"">统计字段</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'DateField',unresize:true,align:'center',width:70}"">日期字段</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'IncItem',minWidth:150}"">项目ID</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'ModelDescr',minWidth:250}"">说　　明</th>" & vbCrlf
	Response.Write "			<th lay-data=""{fixed:'right',align:'center',unresize:true,width:150, toolbar: '#barTable'}"">操作</th>" & vbCrlf
	Response.Write "		</tr></thead>" & vbCrlf
	Response.Write "	</table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-danger layui-btn-disabled"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm hr-btn_fuch"" lay-event=""chkyear"" title=""检查学年/学期""><i class=""hr-icon"">&#xf274;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#SearchBtn"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var soWord = $(""#soWord"").val();" & vbCrlf
	tmpHtml = tmpHtml & "			table.reload(""TableList"", {" & vbCrlf
	tmpHtml = tmpHtml & "				url:""" & layUrl & """,where:{soWord:soWord}" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".addNew"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var loadTips = layer.load(1);" & vbCrlf
	tmpHtml = tmpHtml & "			$.get(""" & ParmPath & "DataModel/AddNew.html?Eve=addNew"", function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:1,content:strForm,title:[""添加数据模型"",""font-size:16""],area:[""700px"", ""80%""],maxmin:true });" & vbCrlf
	tmpHtml = tmpHtml & "				form.render();layer.close(loadTips);" & vbCrlf
	tmpHtml = tmpHtml & "				form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	tmpHtml = tmpHtml & "					$.post(""" & ParmPath & "DataModel/SaveForm.html?ID="",PostData.field, function(result){" & vbCrlf
	tmpHtml = tmpHtml & "						var reData = eval(""("" + result + "")"");" & vbCrlf
	tmpHtml = tmpHtml & "						if(reData.Return){" & vbCrlf
	tmpHtml = tmpHtml & "							layer.alert(reData.reMessge, {icon:1},function(layero, index){layer.closeAll();window.location.reload();});" & vbCrlf
	tmpHtml = tmpHtml & "						}else{" & vbCrlf
	tmpHtml = tmpHtml & "							layer.alert(reData.reMessge, {icon:2});" & vbCrlf
	tmpHtml = tmpHtml & "						}" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					return false;" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		$("".setup"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			layer.open({type:2,content:""" & ParmPath & "DataModel/Config.html"",title:[""系统配置"",""font-size:16""],area:[""700px"", ""80%""],maxmin:true });" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$("".dict"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			location.href=""" & ParmPath & "DataDict/Index.html?ID="";" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""edit""){" & vbCrlf
	tmpHtml = tmpHtml & "				$.get(""" & ParmPath & "DataModel/Edit.html"",{ModelID:data.ModelID, Type:data.TypeID, Eve:obj.event}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "					layer.open({type:1,content:strForm,title:[""编辑数据模型"",""font-size:16""],area:[""700px"", ""80%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					form.render();" & vbCrlf
	tmpHtml = tmpHtml & "					form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	tmpHtml = tmpHtml & "						$.post(""" & ParmPath & "DataModel/SaveForm.html"",PostData.field, function(result){" & vbCrlf
	tmpHtml = tmpHtml & "							var reData = eval(""("" + result + "")"");" & vbCrlf
	tmpHtml = tmpHtml & "							if(reData.Return){" & vbCrlf
	tmpHtml = tmpHtml & "								layer.alert(reData.reMessge, {icon:1},function(layero, index){layer.closeAll();window.location.reload();});" & vbCrlf
	tmpHtml = tmpHtml & "							}else{" & vbCrlf
	tmpHtml = tmpHtml & "								layer.alert(reData.reMessge, {icon:2});" & vbCrlf
	tmpHtml = tmpHtml & "							}" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "						return false;" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					$("".layui-layer-content"").niceScroll();" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""chkyear""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2,content:""" & ParmPath & "DataModel/ChkEduYear.html?ModelID=""+data.ModelID,title:[""检查学年/学期"",""font-size:16""],area:[""540px"", ""360px""],anim:1,moveOut:true});" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		layer.closeAll(""loading"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub GetJsonList()
	Dim vCount, vMSG, tmpJson, tmpData, rsGet, sqlGet, tType
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))

	sqlGet = "Select * From HR_DataModel Where ModelID>0"
	If HR_IsNull(soWord) = False Then sqlGet = sqlGet & " And ModelName like '%" & soWord & "%'"
	sqlGet = sqlGet & " Order By ModelID ASC"
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
			Dim tIncItem
			Do While Not rsGet.EOF
				tIncItem = GetModelIncItem(rsGet("ModelName"))
				If HR_Clng(rsGet("TypeID")) > 0 Then tType = arrItemType(HR_Clng(rsGet("TypeID")) - 1)

				If i > 0 Then tmpData = tmpData & ","
				tmpData = tmpData & "{""ModelID"":" & rsGet("ModelID") & ",""Type"":""" & tType & """,""IncItem"":""" & tIncItem & """"
				For m = 1 To rsGet.Fields.Count - 1
					tmpData = tmpData & ",""" & rsGet.Fields(m).Name & """:""" & rsGet.Fields(m).Value & """"
				Next
				tmpData = tmpData & "}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
		vCount = rsGet.Recordcount
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""未添加数据模型"",""count"":" & vCount & ",""data"":[" & tmpData & "],""limit"":""" & HR_Clng(MaxPerPage) & """,""page"":""" & CurrentPage & """}"
	Response.Write tmpJson
End Sub

Sub EditBody()
	Dim tModuleID : tModuleID = HR_Clng(Request("ModelID"))
	Dim tmpHtml, rsEdit, sqlEdit, arrKey, arrValue
	SubButTxt = "添加"
	sqlEdit = "Select * From HR_DataModel Where ModelID=" & tModuleID
	Set rsEdit = Server.CreateObject("ADODB.RecordSet")
		rsEdit.Open(sqlEdit), Conn, 1, 1
		Redim arrKey(rsEdit.Fields.count - 1)
		Redim arrValue(rsEdit.Fields.count - 1)
		For i = 0 To rsEdit.Fields.count-1
			arrKey(i) = rsEdit.Fields(i).Name
		Next

		If Not(rsEdit.BOF And rsEdit.EOF) Then
			SubButTxt = "修改"
			For i = 0 To rsEdit.Fields.count-1
				arrValue(i) = rsEdit.Fields(i).Value
			Next
		Else
			For i = 0 To rsEdit.Fields.count-1
				arrValue(i) = ""
			Next
		End If
	Set rsEdit = Nothing

	tmpHtml = "<div class=""layer-hr-box"">"
	tmpHtml = tmpHtml & "<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">"

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<label class=""layui-form-label"">类　型:</label>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-block"">"
	For i = 0 To Ubound(arrItemType)
		tmpHtml = tmpHtml & "<input type=""radio"" name=""" & arrKey(1) & """ value=""" & i + 1 & """ title=""" & arrItemType(i) & """ lay-verify=""required"""
		If HR_Clng(arrValue(1)) = i + 1 Then tmpHtml = tmpHtml & " checked"
		tmpHtml = tmpHtml & ">"
	Next
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<label class=""layui-form-label"">模型名称:</label>"
	tmpHtml = tmpHtml & "	<div class=""layui-input-inline""><input type=""text"" name=""" & arrKey(2) & """ value=""" & arrValue(2) & """ placeholder=""模型名称不能为空"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "	<div class=""layui-form-mid""><em class=""hr-help"" data-name=""" & arrKey(2) & """><i class=""hr-icon"">&#xecfd;</i>模型名称不能为空</em></div>"
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item""><label class=""layui-form-label"">自定字段数:</label>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-inline""><input type=""number"" name=""" & arrKey(3) & """ value=""" & arrValue(3) & """ lay-verify=""number"" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-mid layui-word-aux"">字段数必须与数据模板中一致</div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">计分字段:</label>"
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""number"" name=""" & arrKey(4) & """ value=""" & arrValue(4) & """ lay-verify=""number"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">日期字段:</label>"
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""number"" name=""" & arrKey(5) & """ value=""" & arrValue(5) & """ lay-verify=""number"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "	</div>"
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item""><label class=""layui-form-label"">等级字段:</label>"
	tmpHtml = tmpHtml & "<div class=""layui-input-inline""><input type=""number"" name=""" & arrKey(6) & """ value=""" & arrValue(6) & """ lay-verify=""number"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item""><label class=""layui-form-label"">Excel模板:</label>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-block""><input type=""text"" name=""" & arrKey(9) & """ value=""" & arrValue(9) & """ lay-verify=""required"" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item""><label class=""layui-form-label"">字段标题:</label>"
	tmpHtml = tmpHtml & "<div class=""layui-input-block""><textarea name=""" & arrKey(7) & """ id=""" & arrKey(8) & """ placeholder=""标题之间用,分隔"" lay-verify=""content"" class=""layui-textarea"">" & arrValue(7) & "</textarea></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item""><label class=""layui-form-label"">说　明:</label>"
	tmpHtml = tmpHtml & "<div class=""layui-input-block""><textarea name=""" & arrKey(8) & """ id=""" & arrKey(8) & """ placeholder=""数据模型说明"" class=""layui-textarea"">" & arrValue(8) & "</textarea></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<input type=""hidden"" name=""" & arrKey(0) & """ id=""" & arrKey(0) & """ value=""" & arrValue(0) & """><input type=""hidden"" name=""Modify"" value=""True"">"
	tmpHtml = tmpHtml & "<div class=""layui-form-item"">"
	tmpHtml = tmpHtml & "<div class=""layui-input-block""><button class=""layui-btn"" lay-submit lay-filter=""SubPost"">" & SubButTxt & "</button><button type=""reset"" class=""layui-btn layui-btn-primary"">重置</button>"
	tmpHtml = tmpHtml & "<button type=""button"" id=""Preview"" class=""layui-btn layui-btn-normal"">预览</button></div>"
	tmpHtml = tmpHtml & "</div>"
	tmpHtml = tmpHtml & "</form>"
	tmpHtml = tmpHtml & "</div>"

	Response.Write tmpHtml
End Sub
Sub SaveForm()
	Dim tmpJson
	Dim tmpID : tmpID = HR_Clng(Request("ModelID"))
	SubButTxt = "修改"

	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_DataModel Where ModelID=" & tmpID), Conn, 1, 3
		If rsTmp.BOF And rsTmp.EOF Then
			rsTmp.AddNew
			rsTmp("ModelID") = GetNewID("HR_DataModel", "ModelID")
			SubButTxt = "添加"
		End If
		For i = 1 To rsTmp.Fields.count-1
			rsTmp(rsTmp.Fields(i).Name) = Request(rsTmp.Fields(i).Name)
		Next
		rsTmp.Update
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""数据模型 " & Trim(Request("ModelName")) & " " & SubButTxt & "成功！"",""ReStr"":""操作成功！""}"
		rsTmp.Close
	Set rsTmp = Nothing
	Response.Write tmpJson
End Sub

Sub Config()
	Dim tArrField, tStrData
	Dim tArrTitle : tArrTitle = Split("序号,系统名称,标题,网址,根路径,管理目录,上传目录,版权信息,SEO关键,SEO描述,FSO脚本,Mail组件,Mail服务器,Mail帐号,Mail密码,Mail域名,URL类型,验证码文件", ",")
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_Config Where ID=" & ConfigID), Conn, 1, 1
		Redim tArrField(1, rsTmp.Fields.count-1)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			For i = 1 To rsTmp.Fields.count-2
				If i > 1 Then tStrData = tStrData & ","
				tStrData = tStrData & "{""fid"":" & i+1 & ",""name"":""" & rsTmp.Fields(i).Name & """,""value"":""" & rsTmp.Fields(i).Value & """,""title"":""" & tArrTitle(i) & "：""}"
				tArrField(0, i) = rsTmp.Fields(i).Name
				tArrField(1, i) = rsTmp.Fields(i).Value
			Next
		End If
	Set rsTmp = Nothing

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table-header {display: none;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	Response.Write "		<legend>基本参数</legend>" & vbCrlf
	'Response.Write "		<div class=""hr-shrink-x10"">" & vbCrlf
	Response.Write "			<table class=""layui-table"" id=""EditTable"" lay-filter=""EditTable""></table>" & vbCrlf
	'Response.Write "		</div>" & vbCrlf
	Response.Write "	</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "		table.render({" & vbCrlf
	tmpHtml = tmpHtml & "			elem: ""#EditTable"", limit:" & Ubound(tArrTitle) + 1 & ", text:{none:""没有检索到数据""}" & vbCrlf
	tmpHtml = tmpHtml & "			,cols: [[{field:'title',align:'right',style:'color:#39c',width:120,title:'名称'},{field:'value',edit:'text',title:'参数值'}]]" & vbCrlf
	tmpHtml = tmpHtml & "			,data:[" & tStrData & "]" & vbCrlf
	tmpHtml = tmpHtml & "			,skin: 'nob'" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""edit(EditTable)"", function(obj){" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "DataModel/SaveConfig.html"",{fid:obj.data.fid, name:obj.data.name, value:obj.value}, function(reData){ layer.msg(reData.reMessge,{icon:0}); location.reload(); });" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		layer.closeAll(""loading"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub SaveConfig()
	Dim tmpJson, tArrTitle : tArrTitle = Split("序号,系统名称,标题,网址,根路径,管理目录,上传目录,版权信息,SEO关键,SEO描述,FSO脚本,Mail组件,Mail服务器,Mail帐号,Mail密码,Mail域名,URL类型,验证码文件", ",")
	Dim fID : fID = HR_Clng(Request("fid"))
	Dim fName : fName = Trim(ReplaceBadChar(Request("name")))
	Dim fValue : fValue = Trim(Request("value"))
	If fID > 0 And Not(HR_IsNull(fValue)) Then
		Conn.Execute("Update HR_Config Set " & fName & "='" & fValue & "' Where ID=" & ConfigID)
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""" & tArrTitle(fID-1) & " 修改成功！"",""ReStr"":""操作成功！"",""name"":""" & fName & """,""id"":""" & fID & """}"
	End If
	Response.Write tmpJson
End Sub

Sub ChkEduYear()
	Dim tModelID : tModelID = HR_CLng(Trim(Request("ModelID")))
	Dim tModelName, strArrItem, tArrItem, tItemName, tItemID, tSheetName, tVA4
	Set rs = Conn.Execute("Select * From HR_DataModel Where ModelID=" & tModelID)
		If rs.BOF And rs.EOF Then
			ErrMsg = "数据模型不存在或已经删除！"
			Response.Write GetErrBody(1) : Exit Sub
		Else
			tModelName = Trim(rs("ModelName"))
		End If
	Set rs = Nothing
	If FoundInArr("1,3,4,5", tModelID, ",") = False Then
		ErrMsg = tModelName & " 中未包含日期项！"
	End If
	strArrItem = GetModelIncItem(tModelName)
	If HR_IsNull(strArrItem) Then ErrMsg = "没有考核项目使用“" & tModelName & "”模型！"
	If HR_IsNull(ErrMsg) = False Then Response.Write GetErrBody(1) : Exit Sub

	tArrItem = Split(strArrItem, ",")
	ErrMsg = "<h3>更新结果：</h3>"
	ErrMsg = ErrMsg & "<ul class=""result"">" & vbCrlf
	For i = 0 To Ubound(tArrItem)
		tItemID = HR_CLng(tArrItem(i))
		tItemName = GetTypeName("HR_Class", "ClassName", "ClassID", tItemID)
		tSheetName = "HR_Sheet_" & tItemID		'数据表名
		If ChkTable(tSheetName) Then
			sqlTmp = "Select a.*,b.ClassName From " & tSheetName & " a Left Join HR_Class b on a.ItemID=b.ClassID"
			Set rsTmp = Server.CreateObject("ADODB.RecordSet")
				rsTmp.Open sqlTmp, Conn, 1, 3
				If rsTmp.BOF And rsTmp.EOF Then
					ErrMsg = ErrMsg & "<li>项目 " & tItemName & " 中没有数据！</li>" & vbCrlf
				Else
					m = 0
					Do While Not rsTmp.EOF
						tVA4 = GetSchoolYear(ConvertNumDate(HR_CLng(rsTmp("VA4"))), 2)
						rsTmp("scYear") = HR_CLng(tVA4)
						rsTmp("scTerm") = HR_CLng(GetSchoolYear(ConvertNumDate(HR_CLng(rsTmp("VA4"))), 3))
						rsTmp.Update
						m = m + 1
						rsTmp.MoveNext
					Loop
					ErrMsg = ErrMsg & "<li>" & tItemName & "[" & tItemID & "] 共检查 " & m & "/" & HR_CLng(rsTmp.Recordcount) & " 条</li>" & vbCrlf
				End If
			Set rsTmp = Nothing
		Else
			ErrMsg = ErrMsg & "<li>项目 " & tItemName & " 对应的数据表未建立！</li>" & vbCrlf
		End If
	Next
	ErrMsg = ErrMsg & "</ul>" & vbCrlf

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		body {background-color:#f1f1f1;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-err-box {position:absolute;left:10px;top:10px;right:10px;} .error {width:100%;}" & vbCrlf


	tmpHtml = tmpHtml & "	</style>"
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	strHtml = strHtml & "<div class=""hr-err-box"">" & vbCrlf
	strHtml = strHtml & "	<div class=""error"">" & vbCrlf
	strHtml = strHtml & "		<div class=""errorInfo"">" & vbCrlf
	strHtml = strHtml & "			" & ErrMsg & "" & vbCrlf
	strHtml = strHtml & "		</div>" & vbCrlf
	strHtml = strHtml & "	</div>" & vbCrlf
	strHtml = strHtml & "</div>" & vbCrlf
	strHtml = strHtml & getPageFoot(0)
	strHtml = Replace(strHtml, "[@FootScript]", "")

	Response.Write ReplaceCommonLabel(strHtml)
End Sub
%>