<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
Dim isModify : isModify = False
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim SubButTxt : SiteTitle = "管理员管理"

Dim arrSex : arrSex = Split(XmlText("Config", "Sex", ""), "|")
Dim arrManageRank : arrManageRank = Split(XmlText("Config", "ManageRank", ""), "|")
Dim arrStudentType : arrStudentType = Split(XmlText("Common", "StudentType", ""), "|")

If UserRank < 0 Then
	ErrMsg = "对不起，您没有此管理权限！"
	Response.Write UserRank : Response.End
End If

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index", "List" Call MainBody()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveEdit" Call SaveEdit()
	Case "AllData" Call getUserList()
	Case "Preview" Call Preview()
	Case "Delete" Call Delete()

	Case "Authority" Call Authority()
	Case "EditAuth" Call EditAuth()
	Case "SaveAuth" Call SaveAuth()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.soBox .searchBtn {vertical-align:top} .soBox .layui-inline {margin-bottom:8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.soBox .layui-form-select dl {top: 31px;} .soBox .layui-input {height: 30px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.soBox .layui-btn {height: 30px;line-height: 30px;padding: 0 12px;} .soBox .layui-form-select dl dd {padding: 0 5px;line-height: 30px;}" & vbCrlf

	tmpHtml = tmpHtml & "		.hr-false {color:#777}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-label {width:110px}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-input-block {margin-left:140px}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")

	tmpHtml = "<a href=""" & ParmPath & "User/Index.html"">" & SiteTitle & "</a><a><cite>" & SubButTxt & "</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form soBox""><div class=""layui-inline"">搜索管理员：</div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><input class=""layui-input"" name=""soWord"" value=""" & soWord & """ id=""soWord"" placeholder=""会员帐号"" autocomplete=""off"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn""><button class=""layui-btn layui-btn-normal"" data-type=""reload"" id=""SearchBtn""><i class=""layui-icon"">&#xe615;</i>搜索</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_deon"" data-type=""addNew"" id=""addNew"" title=""新增会员""><i class=""hr-icon"">&#xf234;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_fuch"" data-type=""delete"" id=""BatchDel"" title=""批量删除""><i class=""hr-icon"">&#xea64;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-bg-cyan"" data-type=""refresh"" id=""Refresh"" title=""刷新""><i class=""hr-icon"">&#xebbb;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "User/AllData.html',height:'full-150',text:{none:'没有检索到管理员！'},page:true,limit:20,limits:[10,15,20,30,50,100],id:'TableList'}"" lay-filter=""TableList"">"
	Response.Write "		<thead><tr>" & vbCrlf
	Response.Write "			<th lay-data=""{type:'checkbox'}""></th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'UserName',align:'center',width:90,sort: true}"">工　号</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'TrueName',align:'center',width:110}"">姓名</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'ManageRank',unresize:true,align:'center', width:130}"">管理级别</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'StuType',unresize:true}"">管理权限</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'UserID',align:'center',width:70,sort: true}"">UserID</th>" & vbCrlf
	Response.Write "			<th lay-data=""{align:'center',unresize:true,width:170, toolbar: '#barTable'}"">操作</th>" & vbCrlf
	Response.Write "		</tr></thead>" & vbCrlf
	Response.Write "	</table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm"" lay-event=""detail"" title=""查看""><i class=""hr-icon"">&#xf308;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-danger"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"",""form"",""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element, form = layui.form;;" & vbCrlf
	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""detail""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:1,id:""detailWin"",content:"" "",title:[""查看管理员信息"",""font-size:16""],area:[""660px"", ""450px""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "				$.get(""" & ParmPath & "User/Preview.html"",{ID:data.UserID}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "					$(""#detailWin"").html(strForm);" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			} else if(obj.event === 'del'){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.confirm('真的删除选中的管理员吗？<br />删除后无法恢复', {icon:3,title: ""删除提醒""}, function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "User/Delete.html"",{ID:data.UserID}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						if(reData.Return){;" & vbCrlf
	tmpHtml = tmpHtml & "							layer.alert(reData.reMessge, {icon:1,title: ""修改结果""},function(layero, index){layer.closeAll();table.reload(""TableList"");});" & vbCrlf
	tmpHtml = tmpHtml & "						}else{" & vbCrlf
	tmpHtml = tmpHtml & "							layer.alert(reData.reMessge, {icon:2,title: ""删除结果""});" & vbCrlf
	tmpHtml = tmpHtml & "						}" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					layer.close(index);" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			} else if(obj.event === ""edit""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2, id:""EditWin"", content:""" & ParmPath & "User/Edit.html?ID="" + data.UserID,title:[""修改管理员"",""font-size:16""],area:[""760px"", ""450px""], offset:[""100px"",""100px""], maxmin:true });" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		$(""#addNew"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			layer.open({type:2, id:""AddWin"", content:""" & ParmPath & "User/Edit.html?ID="", title:[""添加新管理员"",""font-size:16""], area:[""760px"", ""450px""], offset:[""120px"",""120px""],maxmin:true });" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	'------ 批量删除
	tmpHtml = tmpHtml & "		$(""#BatchDel"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var chkObj = table.checkStatus(""TableList"").data, arrID=[];" & vbCrlf
	tmpHtml = tmpHtml & "			if(chkObj.length==0){layer.tips(""请选择您要删除的会员！"","".laytable-cell-checkbox"",{tips: [1, ""#F60""]});return false;}" & vbCrlf
	tmpHtml = tmpHtml & "			for(var i=0;i<chkObj.length;i++){ arrID.push(chkObj[i].UserID); }" & vbCrlf
	tmpHtml = tmpHtml & "			layer.confirm(""确认要删除选中的“"" + chkObj.length + ""”位会员？<br />删除后将无法恢复。"",{icon:3, title:""删除警告""},function(index){" & vbCrlf
	tmpHtml = tmpHtml & "				$.getJSON(""" & ParmPath & "User/Delete.html"",{ID:arrID.join()}, function(reJson){" & vbCrlf
	tmpHtml = tmpHtml & "					layer.msg(reJson.reMessge,{title:""删除结果"",btn:""关闭"",time:0},function(){ table.reload(""TableList""); });" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	'------ 搜索、刷新
	tmpHtml = tmpHtml & "		$(""#SearchBtn"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var soWord = $(""#soWord"").val();" & vbCrlf
	tmpHtml = tmpHtml & "			table.reload(""TableList"", {" & vbCrlf
	tmpHtml = tmpHtml & "				url:""" & ParmPath & "User/AllData.html"",where:{soWord:soWord}" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#Refresh"").on(""click"", function(){ location.reload();});" & vbCrlf		' 刷新本页
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub getUserList()
	Dim tmpJson, tmpData, rsGet, sqlGet
	Dim vCount, vMSG, tRank, tStuType
	Dim tWord : tWord = Trim(ReplaceBadChar(Request("soWord")))

	sqlGet = "Select a.* From HR_User a Where UserID>0"
	If tWord <> "" Then sqlGet = sqlGet & " And YGXM like '%" & tWord & "%'"
	sqlGet = sqlGet & " Order By UserID ASC"
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
				tRank = arrManageRank(HR_Clng(rsGet("ManageRank")))
				tStuType = FilterArrNull(Trim(rsGet("StuType")), ",")
				If tStuType <> "" Then tStuType = Replace(tStuType, ",", "，")
				If i > 0 Then tmpData = tmpData & ","
				tmpData = tmpData & "{""UserID"":""" & HR_Clng(rsGet("UserID")) & """,""UserName"":""" & rsGet("YGDM") & """,""TrueName"":""" & rsGet("YGXM") & """,""StuType"":""" & tStuType & """,""ManageRank"":""" & tRank & """}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
		vCount = rsGet.Recordcount
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""暂无数据"",""count"":" & vCount & ",""data"":[" & tmpData
	tmpJson = tmpJson & "]}"
	Response.Write tmpJson
End Sub

Sub Preview()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim rsShow, strChk, tManageRank, tSign, tMobile, tTenQQ, tEmail, tSex
	Set rsShow = Conn.Execute("Select * From HR_User Where UserID=" & tmpID )
		If rsShow.BOF And rsShow.EOF Then
			strHtml = "<div class=""layui-row"" style=""text-align:center;padding:20px 0"">"
			strHtml = strHtml & "<a class=""layui-btn layui-btn-lg layui-btn-danger""><i class=""layui-icon"">&#xe69c;</i> 您要查看的管理员信息【ID：" & tmpID & "】不存在！</a></div>"
			Response.Write strHtml
			Exit Sub
		Else
			tManageRank = arrManageRank(HR_Clng(rsShow("ManageRank")))
			strHtml = "<div class=""layui-form layer-hr-box""><table class=""layui-table"" lay-skin=""nob"">"
			strHtml = strHtml & "<colgroup><col width=""120""><col></colgroup>"
			strHtml = strHtml & "<tbody>"

			strHtml = strHtml & "<tr><td style=""text-align:right;"">序　号：</td><td>" & Trim(rsShow("UserID")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">工　号：</td><td>" & HR_Clng(rsShow("YGDM")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">姓　名：</td><td>" & Trim(rsShow("YGXM")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">管理级别：</td><td>" & tManageRank & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">权　限：</td><td>" & Trim(rsShow("StuType")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">备　注：</td><td>" & Trim(rsShow("Describe")) & "</td></tr>"
			strHtml = strHtml & "</tbody>"
			strHtml = strHtml & "</table></div>"
		End If
	Set rsShow = Nothing
	Response.Write strHtml
End Sub

Sub EditBody()
	Dim tUserID : tUserID = HR_Clng(Request("ID"))
	Dim tManageRank, tYGDM, tYGXM, tSign, tTrueName
	Dim tMobile, tTenQQ, tEmail, tAuthA, tAuthB, tStuType, arrAuth
	SubButTxt = "添加"
	sqlTmp = "Select a.*,b.YGXM From HR_User a Left Join HR_Teacher b on a.YGDM=b.YGDM"
	If Action = "Edit" And tUserID > 0 Then
		sqlTmp = sqlTmp & " Where a.UserID=" & tUserID : SubButTxt = "修改"
		Set rsTmp = Conn.Execute(sqlTmp)
			If rsTmp.BOF And rsTmp.EOF Then
				tmpHtml = "<div class=""layui-row"" style=""text-align:center;padding:20px 0"">"
				tmpHtml = tmpHtml & "<a class=""layui-btn layui-btn-lg layui-btn-danger""><i class=""layui-icon"">&#xe69c;</i> 您要修改的会员【ID：" & tUserID & "】不存在！</a></div>"
				Response.Write tmpHtml
				Exit Sub
			Else
				tYGXM = Trim(rsTmp("YGXM"))
				tManageRank = HR_Clng(rsTmp("ManageRank"))
				tYGDM = Trim(rsTmp("YGDM"))
				tAuthA = FilterArrNull(rsTmp("AuthA"), ",")
				tAuthB = FilterArrNull(rsTmp("AuthB"), ",")
				tStuType = FilterArrNull(rsTmp("StuType"), ",")
				tSign = Trim(rsTmp("Describe"))
			End If
		Set rsTmp = Nothing
	End If

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	tmpHtml = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	tmpHtml = tmpHtml & "<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">选择教师：</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><input type=""text"" name=""ygxm"" id=""ygxm"" value=""" & tYGXM & """ lay-verify=""required"" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><span class=""layui-btn layui-btn-sm getBtn"" data-code=""ygdm"" data-name=""ygxm"">查找</span><span class=""tips"">请输入关键字搜索教师，必填项</span></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">工　号：</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><input type=""text"" name=""ygdm"" id=""ygdm"" lay-verify=""number"" value=""" & tYGDM & """ class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">管理级别：</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><select name=""ManageRank"" id=""ManageRank"" lay-verify=""required"">" & vbCrlf
	For i = 0 To Ubound(arrManageRank)
		tmpHtml = tmpHtml & "<option value=""" & i & """"
		If tManageRank = i Then tmpHtml = tmpHtml & " selected"
		tmpHtml = tmpHtml & ">" & arrManageRank(i) & "</option>"
	Next
	tmpHtml = tmpHtml & "</select></div></div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">"  & vbCrlf
	tmpHtml = tmpHtml & "	<label class=""layui-form-label"">学生类别权限：</label>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-block"">"
	For i = 0 To Ubound(arrStudentType)
		tmpHtml = tmpHtml & "<input type=""checkbox"" name=""StuType"" class=""StuType"" lay-skin=""primary"" value=""" & arrStudentType(i) & """ title=""" & arrStudentType(i) & """"
		If FoundInArr(tStuType , arrStudentType(i), ",") Then tmpHtml = tmpHtml & " checked"
		tmpHtml = tmpHtml & ">"
	Next
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""layui-form-item"">"
	tmpHtml = tmpHtml & "<label class=""layui-form-label"">备　　注：</label>"
	tmpHtml = tmpHtml & "<div class=""layui-input-block""><textarea name=""Describe"" id=""Describe"" placeholder=""备注"" class=""layui-textarea"">" & tSign & "</textarea></div>"
	tmpHtml = tmpHtml & "</div>"
	If Action = "Edit" And tUserID > 0 Then tmpHtml = tmpHtml & "<input type=""hidden"" name=""UserID"" id=""UserID"" value=""" & tUserID & """><input type=""hidden"" name=""Modify"" value=""True"">"
	tmpHtml = tmpHtml & "<div class=""hr-pop-fix"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""hr-grids hr-btn-group"">" & vbCrlf
	tmpHtml = tmpHtml & "		<em><button class=""layui-btn hr-btn_fuch"" lay-submit lay-filter=""EditPost"" id=""EditPost"" title=""保存""><i class=""hr-icon"">&#xf0c7;</i>保存</button></em>" & vbCrlf
	tmpHtml = tmpHtml & "		<em><button class=""layui-btn layui-btn-primary"" type=""reset"" name=""reset"" id=""refresh"" title=""重置""><i class=""hr-icon"">&#xf343;</i>重置</button></em>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "</form>"
	tmpHtml = tmpHtml & "</div>"
	Response.Write tmpHtml
	Response.Write "<div class=""hr-place-h50""></div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, form = layui.form, laydate = layui.laydate;" & vbCrlf
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
	tmpHtml = tmpHtml & "			$.post(""" & ParmPath & "User/SaveEdit.html"", $(""#EditForm"").serialize(), function(formResult){" & vbCrlf
	tmpHtml = tmpHtml & "				var icon=0; if(formResult.Return){icon=1};" & vbCrlf
	tmpHtml = tmpHtml & "				layer.alert(formResult.reMessge,{icon:icon,time:0,btn:""关闭""},function(){" & vbCrlf
	tmpHtml = tmpHtml & "					if(formResult.Return){var index1=parent.layer.getFrameIndex(window.name); parent.layui.table.reload(""TableList""); parent.layer.close(index1); }else{ layer.close(layer.index) }" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			},""json"");" & vbCrlf
	tmpHtml = tmpHtml & "			return false;" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub
Sub SaveEdit()
	ErrMsg = ""
	Dim tmpJson, rsGet, sqlGet
	Dim uID : uID = HR_Clng(Request("UserID"))
	Dim tYGXM : tYGXM = Trim(ReplaceBadChar(Request("ygxm")))
	Dim tYGDM : tYGDM = HR_Clng(Request("ygdm"))
	Dim tManageRank : tManageRank = HR_Clng(Request("ManageRank"))
	Dim tmpStuType : tmpStuType = Trim(Request("StuType"))
	Dim tDescribe : tDescribe = Trim(Request("Describe"))

	tmpStuType = FilterArrNull(tmpStuType, ",")

	tmpJson = "{""Return"":false,""Err"":500,""reMessge"":""[@ErrMsg]"",""ReStr"":""操作失败！""}"
	sqlGet = "Select * From HR_User Where YGDM=" & tYGDM						'判断工号是否已经存在
	If uID>0 Then sqlGet = sqlGet & " And UserID<>" & uID	'当修改时
	Set rsGet = Conn.Execute(sqlGet)
		If Not(rsGet.BOF And rsGet.EOF) Then
			ErrMsg = "该工号已存在！"
			Response.Write Replace(tmpJson, "[@ErrMsg]", ErrMsg) : Exit Sub
		End If
	Set rsGet = Nothing

	sqlGet = "Select * From HR_User Where UserID=" & uID		'判断帐号是否已经存在
	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 3
		If rsGet.BOF And rsGet.EOF Then
			rsGet.AddNew
			uID = GetNewID("HR_User", "UserID")
			rsGet("UserID") = uID
			rsGet("RegTime") = Now()
			rsGet("Passed") = 1
			rsGet("Locked") = 0
			SubButTxt = "添加"
		End If
		rsGet("YGXM") = tYGXM
		rsGet("YGDM") = tYGDM
		rsGet("ManageRank") = tManageRank
		rsGet("StuType") = tmpStuType
		rsGet("Describe") = tDescribe
		rsGet.Update
		rsGet.Close
		ErrMsg = "管理员" & tYGXM & " 的资料保存成功！"
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作成功！""}"
		Call RecordFrontLog(1, SubButTxt & "管理员", SubButTxt & "管理员“" & tYGXM & "”成功！操作者：" & UserYGXM & "[" & UserID & "]", 1, "")
	Set rsGet = Nothing
	Response.Write tmpJson
End Sub

Sub Delete()
	Dim tmpJson, rsDel, sqlDel, strDel, arrDel, iDel
	strDel = Trim(ReplaceBadChar(Request("ID")))
	strDel = DelRightComma(strDel)
	arrDel = Split(strDel, ",")
	iDel = 0
	For i = 0 To Ubound(arrDel)
		Set rsTmp = Server.CreateObject("ADODB.RecordSet")
			rsTmp.Open("Select * From HR_User Where UserID=" & HR_Clng(arrDel(i))), Conn, 1, 3
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				If rsTmp("UserID") = UserID Then
					tmpJson = "管理员“" & rsTmp("UserName") & "”不能删除自己！<br>"
				Else
					rsTmp.Delete
					rsTmp.Close
					iDel = iDel + 1
				End If
			End If
		Set rsTmp = Nothing
	Next
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""提示：" & tmpJson & "" & iDel & "/" & Ubound(arrDel) + 1 & "删除成功！"",""ReStr"":""操作成功！""}"
	Call RecordFrontLog(1, "删除管理员", "管理员ID：" & UserID & "，操作：删除管理员“" & iDel & "”名", 1, Request.QueryString())
	Response.Write tmpJson
End Sub

Sub Authority()
	SubButTxt = "管理员权限设置"
	strHtml = "<link rel=""stylesheet"" type=""text/css"" href=""" & InstallDir & "Static/Admin/css/hr.lay.css?v=1.0.1"" />"
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "	</style>"
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", Page_Title)
	strHeadHtml = Replace(strHeadHtml, "[@Head_style]", strHtml)
	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.min.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/layui/layui.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "		layui.use([""layer"", ""form"",""element""], function(){" & vbCrlf
	strHtml = strHtml & "			var layer = layui.layer, form = layui.form,element = layui.element;" & vbCrlf
	strHtml = strHtml & "			layer.config({skin:""layer-hr""});var loadInit = layer.load();" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Head_script]", strHtml)
	Response.Write strHeadHtml
	strHtml = "<a href=""" & ParmPath & "User/Index.html"">" & Page_Title & "</a><a href=""" & ParmPath & "User/Authority.html"">管理权限</a><cite>" & SubButTxt & "</cite></a>"
	strNavPath = Replace(strNavPath, "[@Module_Path]", strHtml)
	Response.Write strNavPath

	Response.Write "<div class=""hr-padd"">" & vbCrlf
	Response.Write "	<div class=""layer-hr_searchBox"">搜索会员：<div class=""layui-inline""><input class=""layui-input"" name=""SearchWord"" id=""SearchWord"" autocomplete=""off"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group""><button class=""layui-btn layui-btn-normal"" data-type=""reload"" id=""SearchBtn""><i class=""layui-icon"">&#xe615;</i>搜索</button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "User/AllData.html',page:true,limit:20,limits:[10,15,20,30,50],id:'TableList'}"" lay-filter=""TableList"">"
	Response.Write "		<thead><tr>" & vbCrlf
	Response.Write "			<th lay-data=""{type:'checkbox'}""></th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'UserName',unresize:true,width:120,sort: true}"">登陆帐号</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'NickName',unresize:true,width:120}"">会员昵称</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'ManageRank',unresize:true,align:'center', width:130}"">管理级别</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'Area',unresize:true}"">管理区域</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'Passed',align:'center',width:60}"">权限</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'UserID',align:'center',width:60,sort: true}"">ID</th>" & vbCrlf
	Response.Write "			<th lay-data=""{align:'center',unresize:true,width:120, toolbar: '#barTable'}"">操作</th>" & vbCrlf
	Response.Write "		</tr></thead>" & vbCrlf
	Response.Write "	</table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm"" lay-event=""detail"" title=""查看""><i class=""hr-icon"">&#xf308;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""editAuth"" title=""编辑权限""><i class=""hr-icon"">&#xea68;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"",""form"",""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form;;" & vbCrlf
	strHtml = strHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	strHtml = strHtml & "			var data = obj.data;" & vbCrlf
	strHtml = strHtml & "			if(obj.event === ""detail""){" & vbCrlf
	strHtml = strHtml & "				$.get(""" & ParmPath & "User/Preview.html"",{ID:data.UserID}, function(strForm){" & vbCrlf
	strHtml = strHtml & "					layer.open({type:1,content:strForm,title:[""查看管理员信息"",""font-size:16""],area:[""600px"", ""80%""],maxmin:true});" & vbCrlf
	strHtml = strHtml & "					form.render();$("".layui-layer-content"").niceScroll();" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			} else if(obj.event === ""editAuth""){" & vbCrlf
	strHtml = strHtml & "				$.get(""" & ParmPath & "User/EditAuth.html?ID="" + data.UserID, function(strForm){" & vbCrlf
	strHtml = strHtml & "					layer.open({type:1,content:strForm,title:[""修改管理权限"",""font-size:16""],area:[""500px"", ""400px""],maxmin:true });" & vbCrlf
	strHtml = strHtml & "					form.render();" & vbCrlf
	strHtml = strHtml & "					form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "						$.post(""" & ParmPath & "User/SaveAuth.html"", PostData.field, function(result){" & vbCrlf
	strHtml = strHtml & "							var reData = eval(""("" + result + "")"");" & vbCrlf
	strHtml = strHtml & "							if(reData.Return){" & vbCrlf
	strHtml = strHtml & "								layer.alert(reData.reMessge, {icon:1,title: ""修改结果""},function(layero, index){layer.closeAll();table.reload(""TableList"");});" & vbCrlf
	strHtml = strHtml & "							}else{" & vbCrlf
	strHtml = strHtml & "								layer.alert(reData.reMessge, {icon:2,title: ""修改结果""});" & vbCrlf
	strHtml = strHtml & "							}" & vbCrlf
	strHtml = strHtml & "						});" & vbCrlf
	strHtml = strHtml & "						return false;" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#SearchBtn"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var key1 = $(""#SearchWord"").val();" & vbCrlf
	strHtml = strHtml & "			table.reload(""TableList"", {" & vbCrlf
	strHtml = strHtml & "				url:""" & ParmPath & "User/AllData.html"",where: {SearchWord:key1}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@Foot_script]", strHtml)
	Response.Write strFootHtml
End Sub
Sub EditAuth()
	Dim tmpID, tUserID : tUserID = HR_Clng(Request("ID"))
	Dim tmpHtml, rsGet, tAreaID, tAreaAuth, tArrtID, tAttrAuth, tRouteID, tRouteAuth, tChecked
	Set rsGet = Conn.Execute("Select Top 1 * From HR_UserAuth Where UserID=" & tUserID)
		If Not(rsGet.BOF And rsGet.EOF) Then
			tmpID = HR_Clng(rsGet("ID"))
			tAreaID = rsGet("AreaID")
			tAreaAuth = HR_Clng(rsGet("AreaAuth"))
		End If
	Set rsGet = Nothing
	SubButTxt = "更新" : If tAreaAuth = 1 Then tChecked = " checked"

	tmpHtml = "<div class=""layer-hr-box"">" & vbCrlf
	tmpHtml = tmpHtml & "<form class=""layui-form layui-form-pane"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<label class=""layui-form-label"">选择区域：</label>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-input-block""><select name=""AreaID"" lay-verify=""required"" lay-search>"
	Set rsTmp = Conn.Execute("Select * From HR_Area Order By AreaID ASC")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				tmpHtml = tmpHtml & "<option value=""" & rsTmp("AreaID") & """"
				If tAreaID = rsTmp("AreaID") Then tmpHtml = tmpHtml & " selected"
				tmpHtml = tmpHtml & ">" & Trim(rsTmp("AreaName")) & "</option>"
				rsTmp.MoveNext
			Loop
		End If
	Set rsTmp = Nothing
	tmpHtml = tmpHtml & "</select></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"" pane=""""><label class=""layui-form-label"">管理权限：</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-block""><input type=""checkbox"" name=""AreaAuth"" value=""1"" lay-skin=""switch"" lay-filter=""switchDoc"" lay-text=""ON|OFF""" & tChecked & "></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""UserID"" value=""" & tUserID & """><input type=""hidden"" name=""Modify"" value=""True"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-input-block""><button class=""layui-btn layui-btn-sm"" lay-submit lay-filter=""SubPost"">" & SubButTxt & "</button><button type=""reset"" class=""layui-btn layui-btn-primary layui-btn-sm"">重置</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "</form>"
	tmpHtml = tmpHtml & "</div>"

	Response.Write tmpHtml
End sub
Sub SaveAuth()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim rsSave, tmpJson, tUserName, tUserID
	tUserID = HR_Clng(Request("UserID"))
	tUserID = HR_Clng(GetTypeName("HR_User", "UserID", "UserID", tUserID))
	tUserName = Trim(GetTypeName("HR_User", "UserName", "UserID", tUserID))
	If UserRank = 2 And tUserID>0 Then
		Set rsSave = Server.CreateObject("ADODB.RecordSet")
			rsSave.Open("Select Top 1 * From HR_UserAuth Where UserID=" & tUserID & " Order By ID ASC" ), Conn, 1, 3
			If rsSave.BOF And rsSave.EOF Then
				rsSave.AddNew
				rsSave("ID") = GetNewID("HR_UserAuth", "ID")
				rsSave("UserID") = tUserID
			End If
			rsSave("AreaID") = HR_Clng(Request("AreaID"))
			rsSave("AreaAuth") = HR_Clng(Request("AreaAuth"))
			rsSave("ArrtID") = HR_Clng(Request("ArrtID"))
			rsSave("AttrAuth") = HR_Clng(Request("AttrAuth"))
			rsSave("RouteID") = HR_Clng(Request("RouteID"))
			rsSave("RouteAuth") = HR_Clng(Request("RouteAuth"))
			rsSave.Update
			rsSave.Close
		Set rsTmp = Nothing
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""管理员“" & tUserName & "”权限更新成功！"",""ReStr"":""操作成功！""}"
		Call RecordFrontLog(1, "设置管理权限", "管理员ID：" & UserID & "，操作：更新管理员“" & tUserName & "”权限", 1, Request.QueryString())
	Else
		tmpJson = "{""Return"":false,""Err"":400,""reMessge"":""对不起，您不是超级管理员或该管理员不存在！" & tmpID & """,""ReStr"":""操作失败！""}"
		Call RecordFrontLog(1, "设置管理权限", "管理员ID：" & UserID & "，操作：更新管理员权限失败", 0, Request.QueryString())
	End If
	Response.Write tmpJson
End Sub

%>