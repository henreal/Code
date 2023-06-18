<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="../hrBase/incKPI.asp"-->
<!--#include file="./pc_CourseInc.asp"-->

<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "课程管理"

Dim arrStudentType : arrStudentType = Split(XmlText("Common", "StudentType", ""), "|")
Dim arrSex : arrSex = Split(XmlText("Config", "Sex", ""), "|")
Dim IsAdd : IsAdd = HR_CBool(XmlText("Common", "AddSwitch", "0"))
Dim IsImport : IsImport = HR_CBool(XmlText("Common", "ImportSwitch", "0"))
Dim strExtname : strExtname = "jpg,jpeg,png,bmp,gif,xls,xlsx,pdf,doc,docx,txt,rar,zip"
Dim scriptCtrl, SubButTxt

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveEdit" Call SaveEdit()
	Case "winSelectTeacher" Call winSelectTeacher()
	Case "winTeacherData" Call winTeacherData()
	
	Case "jsonList" Call GetJsonList()
	Case "Preview" Call Preview()
	Case "viewAttach" Call viewAttach()

	Case "Delete" Call Delete()

	Case "levelData" Call levelData()
	Case "CampusData" Call CampusData()

	Case "applyModify" Call applyModify()
	Case "SaveApply" Call SaveApply()

	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tItemName, rsList, tChild, tTemplate, tSheetName, tStuType, arrStuType, tFieldHead, arrHead, tFieldLen
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tSearchWord : tSearchWord = Trim(ReplaceBadChar(Request("SearchWord")))
	Dim soYear : soYear = HR_Clng(Request("soYear"))
	If soYear < 2000 Then soYear = DefYear	'如果学年不正确，取系统默认学年

	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
		If rsTmp.BOF And rsTmp.EOF Then
			ErrMsg = "您要查询的业绩项目不存在！"
			Response.Write GetErrBody(0) : Exit Sub
		Else
			tItemName = rsTmp("ClassName")
			tChild = HR_Clng(rsTmp("Child"))
			tTemplate = Trim(rsTmp("Template"))
			tStuType = Trim(rsTmp("StudentType"))
			tFieldLen = HR_Clng(rsTmp("FieldLen"))
			tFieldHead = Trim(rsTmp("FieldHead"))
			tSheetName = "HR_Sheet_" & tItemID
		End If
	Set rsTmp = Nothing

	'取标题
	Redim arrHead(tFieldLen)
	If tFieldHead <> "" Then
		arrHead = Split(tFieldHead, ",")
		If Ubound(arrHead) <> tFieldLen-1 Then Redim Preserve arrHead(tFieldLen-1)
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.iframe-nav .navBtn .navLayer {font-size: 16px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.iframe-nav .navBtn .navLayer i {font-size: 16px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-field-title {margin-bottom:6px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.xlsData {box-sizing:border-box;padding:0 15px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.xlsData ul{padding-left:15px;line-height:180%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding: 5px 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.panBox {display:flex;align-items:stretch;flex-wrap:wrap;}" & vbCrlf
	tmpHtml = tmpHtml & "		.panItem {width:370px;box-sizing:border-box;padding:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.panItem .panBorder {margin:0px;display:block;width:100%;border:1px solid #ccc;box-sizing:border-box;padding:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.panItem .panBorder em {font-size:16px;} .panItem .panBorder em i {font-size:18px;color:#1E9FFF}" & vbCrlf
	tmpHtml = tmpHtml & "		.panItem .panBorder span {display:block;color:#999;height:90px;overflow:auto}" & vbCrlf
	tmpHtml = tmpHtml & "		.tipsWarn {color:#f30}" & vbCrlf
	tmpHtml = tmpHtml & "		.pass-true {background-color:#740} .pass-false {background-color:#060}" & vbCrlf
	tmpHtml = tmpHtml & "		.btnAffirm {background:#158;} .reSum {color:#f30;}" & vbCrlf

	tmpHtml = tmpHtml & "		.hr-tablebtn .layui-btn {height:28px;line-height:28px;padding:0 10px;font-size:1.1rem;}" & vbCrlf		'表头工具集
	tmpHtml = tmpHtml & "		.sumbar b {color:#F60;padding:0 2px}" & vbCrlf		'表头汇总

	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead("Desktop", 1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	tmpHtml = "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		layui.use([""layer"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var layer=layui.layer, element=layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	
	tmpHtml = "<a href=""" & ParmPath & "Course.html?ItemID=" & tItemID & """>" & SiteTitle & "</a><a><cite>" & tItemName & "</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)

	Dim layUrl : layUrl = ParmPath & "Course/jsonList.html"

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<fieldset class=""layui-elem-field layui-field-title"" style=""margin-top:1px;""><legend>" & tItemName & "</legend></fieldset>" & vbCrlf


		Response.Write "	<form class=""layui-form soBox"" id=""SearchForm""><div class=""layui-inline"">筛选：</div>"
		Response.Write "<div class=""layui-inline"" style=""width:150px""><select name=""SchoolYear"" id=""SchoolYear""><option value="""">选择学年</option>" & GetYearOption(0, soYear) & "</select></div>" & vbCrlf
		If tTemplate = "TempTableA" Then Response.Write "<div class=""layui-inline""><select name=""Campus"" id=""Campus"" lay-search=""""><option value="""">选择/搜索校区</option>" & GetCampusOption("", 0) & "</select></div>" & vbCrlf
		Response.Write "<div class=""layui-inline""><select name=""soSort"" id=""soSort""><option value="""">选择排序方式</option><option value=""importTimeUP"">上传时间正序↑</option><option value=""importTimeDown"">上传时间倒序↓</option>"
		Response.Write "<option value=""xhUP"">序号正序↑</option><option value=""xhDown"">序号倒序↓</option></select></div>" & vbCrlf
		
		If tTemplate = "TempTableA" Then
			Response.Write "		<div class=""layui-inline""><select name=""VA9"" id=""VA9"" lay-verify=""required"" lay-search=""""><option value="""">选择/搜索课程名称</option>" & GetCourseOption("", 0) & "</select></div>" & vbCrlf
			Response.Write "		<div class=""layui-inline""><input class=""layui-input"" name=""soWord"" id=""soWord"" placeholder=""课程内容"" autocomplete=""off"" /></div>" & vbCrlf
		Else
			Response.Write "		<div class=""layui-inline""><input class=""layui-input"" name=""soWord"" id=""soWord"" placeholder=""项目名称"" autocomplete=""off"" /></div>" & vbCrlf
		End If
		If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Then
			Response.Write "		<div class=""layui-inline""><input class=""layui-input"" name=""Period"" id=""Period"" placeholder=""时间段"" autocomplete=""off"" /></div>" & vbCrlf
		End If

		If tStuType <> "" Then
			Response.Write "		<div class=""layui-inline""><select name=""StudentType"" id=""StudentType"" lay-verify=""required""><option value="""">选择学生类别</option>" & vbCrlf
			arrStuType = Split(tStuType, ",")
			For i = 0 To Ubound(arrStuType)
				Response.Write "<option value=""" & arrStuType(i) & """>" & arrStuType(i) & "</option>" & vbCrlf
			Next
			Response.Write "</select></div>" & vbCrlf
		End If
		Response.Write "		<div class=""layui-inline""><input type=""checkbox"" name=""IsPass"" id=""IsPass"" value=""1"" title=""未审核""></div>" & vbCrlf
		Response.Write "		<div class=""layui-inline""><input type=""checkbox"" name=""IsAffirm"" id=""IsAffirm"" value=""1"" title=""已确认""></div>" & vbCrlf
		Response.Write "		<div class=""layui-inline""><input type=""checkbox"" name=""IsRetreat"" id=""IsRetreat"" value=""1"" title=""已退回""></div>" & vbCrlf
		Response.Write "		<div class=""layui-inline searchBtn""><button class=""layui-btn layui-bg-black"" type=""button"" data-type=""reload"" id=""SearchBtn""><i class=""layui-icon"">&#xe615;</i>搜索</button></div>" & vbCrlf
		Response.Write "		<div class=""layui-btn-group searchBtn"">" & vbCrlf
		Response.Write "			<button class=""layui-btn layui-bg-cyan"" type=""button"" data-type=""batchaffirm"" id=""BatchAffirm"" title=""确认提交""><i class=""hr-icon"">&#xebc5;</i>确认提交</button>" & vbCrlf
		Response.Write "			<button class=""layui-btn oneAffirm"" type=""button"" data-type=""oneAffirm"" id=""oneAffirm"" title=""一键确认""><i class=""hr-icon"">&#xf046;</i>一键确认</button>" & vbCrlf
		Response.Write "		</div>" & vbCrlf
		Response.Write "	<input type=""hidden"" name=""ItemID"" value=""" & tItemID & """>" & vbCrlf
		Response.Write "	</form>" & vbCrlf

		Response.Write "	<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "Course/jsonList.html?ItemID=" & tItemID & "&SearchWord=" & tSearchWord & "',toolbar: '#tabbtn',height:'full-210',page:true,limit:20,limits:[10,15,20,30,50,100,200],text:{none:'暂时还没有相关课程进度'},id:'TableList',done:function(res){$('.reSum').html(res.sumVA3);} }"" lay-filter=""TableList"">"
		Response.Write "		<thead><tr>" & vbCrlf
		Response.Write "			<th lay-data=""{type:'checkbox',unresize:true,align:'center',width:60}""></th>" & vbCrlf
		Response.Write "			<th lay-data=""{field:'VA0',align:'center', width:70,sort:true}"">序号</th>" & vbCrlf
		If tStuType <> "" Then Response.Write "			<th lay-data=""{field:'StudentType', width:80}"">类别</th>" & vbCrlf
		Response.Write "			<th lay-data=""{field:'VA1',unresize:true, width:80,sort:true}"">工号</th>" & vbCrlf
		Response.Write "			<th lay-data=""{field:'VA2',width:90,sort:true}"">" & arrHead(2) & "</th>" & vbCrlf
		Response.Write "			<th lay-data=""{field:'KSMC',width:120,sort:true}"">科室</th>" & vbCrlf
		Response.Write "			<th lay-data=""{field:'PRZC',width:100}"">职称</th>" & vbCrlf
		Response.Write "			<th lay-data=""{field:'VA3',unresize:true,sort:true, width:60,templet:'#tplUnit'}"">" & arrHead(3) & "</th>" & vbCrlf
		For i = 4 To Ubound(arrHead)
			Response.Write "			<th lay-data=""{field:'VA" & i & "',minWidth:100,sort:true}"">" & arrHead(i) & "</th>" & vbCrlf
			If tTemplate = "TempTableA" Then
				If i = 7 Then Response.Write "			<th lay-data=""{field:'Time',minWidth:100,sort:true}"">时　间</th>" & vbCrlf
			End If
		Next
		Response.Write "			<th lay-data=""{field:'AppendTime',align:'center',width:160}"">上传时间</th>" & vbCrlf
		Response.Write "			<th lay-data=""{field:'Attach',unresize:true,sort:true,align:'center',width:60}"">附件</th>" & vbCrlf
		Response.Write "			<th lay-data=""{field:'Affirm',unresize:true,sort:true,align:'center',width:65}"">确认</th>" & vbCrlf
		Response.Write "			<th lay-data=""{field:'Retreat',unresize:true,sort:true,align:'center',width:65}"">退回</th>" & vbCrlf
		Response.Write "			<th lay-data=""{align:'center',unresize:true,width:260, toolbar: '#barTable'}"">操作</th>" & vbCrlf
		Response.Write "		</tr></thead>" & vbCrlf
		Response.Write "	</table>" & vbCrlf
		Response.Write "	<script type=""text/html"" id=""tabbtn"">" & vbCrlf		'表头模板
		Response.Write "		<div class=""hr-rows tpltools"">" & vbCrlf
		Response.Write "			<div class=""layui-btn-group hr-tablebtn"">" & vbCrlf
		If IsAdd Then
			Response.Write "				<button type=""button"" class=""layui-btn hr-btn_deon"" lay-event=""addNew"" title=""新增课程""><i class=""hr-icon"">&#xecfb;</i></button>" & vbCrlf
		Else
			Response.Write "				<button type=""button"" class=""layui-btn layui-btn-disabled"" title=""新增课程已关闭""><i class=""hr-icon"">&#xecfb;</i></button>" & vbCrlf
		End If
		Response.Write "				<button type=""button"" class=""layui-btn hr-btn_fuch"" lay-event=""batchdel"" title=""批量删除""><i class=""hr-icon"">&#xea64;</i></button>" & vbCrlf
		Response.Write "				<button type=""button"" class=""layui-btn hr-btn_skyblue"" lay-event=""refresh"" title=""刷新""><i class=""hr-icon"">&#xebbb;</i></button>" & vbCrlf
		Response.Write "			</div>" & vbCrlf
		Response.Write "			<div class=""sumbar"">总" & arrHead(3) & "：<b class=""reSum"">0.0</b></div>" & vbCrlf		'学时汇总
		Response.Write "		</div>" & vbCrlf
		Response.Write "	</script>" & vbCrlf
		Response.Write "	<script type=""text/html"" id=""tplUnit"">" & vbCrlf
		Response.Write "		{{d.VA3}}{{d.Unit}}" & vbCrlf
		Response.Write "	</script>" & vbCrlf
		Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
		Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
		Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""details"" title=""查看""><i class=""hr-icon"">&#xefb9;</i></a>" & vbCrlf
		Response.Write "			{{#  if(d.isEdit===true){ }}" & vbCrlf
		Response.Write "				<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
		Response.Write "			{{#  }else{ }}" & vbCrlf
		Response.Write "				<a class=""layui-btn layui-btn-sm layui-btn-disabled"" title=""禁止编辑(已审核或无权限)""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
		Response.Write "			{{#  } }}" & vbCrlf

		Response.Write "			{{#  if(d.isDel===true){ }}" & vbCrlf
		Response.Write "				<a class=""layui-btn layui-btn-sm layui-btn-danger"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
		Response.Write "			{{#  }else{ }}" & vbCrlf
		Response.Write "				<a class=""layui-btn layui-btn-sm layui-btn-disabled"" title=""禁止删除(已审核或无权限)""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
		Response.Write "			{{#  } }}" & vbCrlf
		Response.Write "			<a class=""layui-btn layui-btn-sm layui-bg-cyan"" lay-event=""apply"" title=""申请修改""><i class=""hr-icon"">&#xf044;</i></a>" & vbCrlf
		Response.Write "			{{#  if(d.isAffirm===true){ }}" & vbCrlf
		Response.Write "				<a class=""layui-btn layui-btn-sm layui-btn-disabled"" title=""已经确认或已审核""><i class=""hr-icon"">&#xebc5;</i></a>" & vbCrlf
		Response.Write "			{{#  }else{ }}" & vbCrlf
		Response.Write "				<a class=""layui-btn layui-btn-sm oneAffirm"" lay-event=""affirm"" title=""确认提交""><i class=""hr-icon"">&#xebc5;</i></a>" & vbCrlf
		Response.Write "			{{#  } }}" & vbCrlf
		Response.Write "		</div>" & vbCrlf
		Response.Write "	</script>" & vbCrlf

	Response.Write "</div>" & vbCrlf

	Dim soStime, soEtime
	soStime = year(Now())-1 & "-07-01"
	soEtime = year(Now()) & "-06-30"

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.hr.serialize.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""laytpl"", ""form"", ""upload"", ""laydate"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, upload = layui.upload, laytpl = layui.laytpl, laydate = layui.laydate;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element, form = layui.form;" & vbCrlf

	tmpHtml = tmpHtml & "		laydate.render({ elem:""#Period"",range:""～""});" & vbCrlf		'日期选择
	tmpHtml = tmpHtml & "		$("".searchBtn button"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var btnEvent = $(this).data(""type"");" & vbCrlf
	tmpHtml = tmpHtml & "			if(btnEvent==""reload""){" & vbCrlf
	tmpHtml = tmpHtml & "				var arrForm = $(""#SearchForm"").hr_serialize(), postStr={};" & vbCrlf
	tmpHtml = tmpHtml & "				$.each(arrForm, function(key, val){ postStr[this.name]=this.value; });" & vbCrlf		'表单序列转json
	tmpHtml = tmpHtml & "				console.log(postStr);" & vbCrlf
	tmpHtml = tmpHtml & "				table.reload(""TableList"", {" & vbCrlf
	tmpHtml = tmpHtml & "					url:'" & ParmPath & "Course/jsonList.html', where:postStr" & vbCrlf
	tmpHtml = tmpHtml & "					,done: function(res, curr, count){" & vbCrlf
	tmpHtml = tmpHtml & "						$("".reSum"").html(res.sumVA3);" & vbCrlf
	tmpHtml = tmpHtml & "					}" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "				return false;" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""refresh""){" & vbCrlf
	tmpHtml = tmpHtml & "				location.reload();" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""addNew""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2,content:""" & ParmPath & "Course/AddNew.html?ItemID=" & tItemID & """,title:[""添加课程"",""font-size:16""],area:[""760px"", ""82%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""batchaffirm""){" & vbCrlf	'批量确认提交
	tmpHtml = tmpHtml & "				var arrID = """", chkStatus = table.checkStatus(""TableList"");" & vbCrlf
	tmpHtml = tmpHtml & "				if(chkStatus.data.length==0){layer.tips(""请选择您要提交的课程记录！"",""#BatchAffirm"",{tips: [3, ""#F30""]});return false;}" & vbCrlf
	tmpHtml = tmpHtml & "				for(var i=0; i<chkStatus.data.length; i++){" & vbCrlf
	tmpHtml = tmpHtml & "					if(i > 0){arrID = arrID + "",""}" & vbCrlf
	tmpHtml = tmpHtml & "					arrID = arrID + chkStatus.data[i].ID;" & vbCrlf
	tmpHtml = tmpHtml & "				}" & vbCrlf
	
	tmpHtml = tmpHtml & "				layer.confirm(""您确定选中的 "" + chkStatus.data.length + "" 条课程业绩没问题？"",{icon: 3, title:""重要提示""},function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "CourseProof/Affirm.html"",{ItemID:" & tItemID & ", ID:arrID}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.msg(reData.reMessge,{btn:""关闭"",time:0}); table.reload(""TableList"");" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""oneAffirm""){" & vbCrlf	'一键确认提交
	tmpHtml = tmpHtml & "				layer.confirm(""您确定所有考核项目的课程业绩没问题？"",{icon: 3, title:""重要提示""},function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "CourseProof/oneAffirm.html"", function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.msg(reData.reMessge,{btn:""关闭"",time:0}); table.reload(""TableList"");" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(btnEvent==""delete""){" & vbCrlf
	tmpHtml = tmpHtml & "				var checkStatus = table.checkStatus(""TableList"");" & vbCrlf
	tmpHtml = tmpHtml & "				if(checkStatus.data.length==0){layer.tips(""请选择您要删除的课程！"",""#BatchDel"",{tips: [3, ""#F60""]});return false;}" & vbCrlf
	tmpHtml = tmpHtml & "				layer.confirm(""确认要删除选中的“"" + checkStatus.data.length + ""”条课程记录？<br />删除后无法恢复。"",{icon: 3, title:""重要提示""},function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					var arrID = """";" & vbCrlf
	tmpHtml = tmpHtml & "					for(var i=0;i<checkStatus.data.length;i++){" & vbCrlf
	tmpHtml = tmpHtml & "						if(i > 0){arrID = arrID + "",""}" & vbCrlf
	tmpHtml = tmpHtml & "						arrID = arrID + checkStatus.data[i].ID;" & vbCrlf
	tmpHtml = tmpHtml & "					}" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "Course/Delete.html"",{ItemID:" & tItemID & ",ID:arrID}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.msg(strForm.reMessge,{btn:""关闭"",time:0},function(){table.reload(""TableList"");});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					return false;" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	tmpHtml = tmpHtml & "			var data = obj.data;" & vbCrlf
	tmpHtml = tmpHtml & "			if(obj.event === ""details""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2,content:'" & ParmPath & "Course/Preview.html?ItemID='+ data.CourseID +'&ID='+data.ID,title:[""查看课程信息"",""font-size:16""],area:[""700px"", ""80%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""add""||obj.event === ""edit""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:2,content:'" & ParmPath & "Course/Edit.html?ItemID=' + data.CourseID + '&ID='+data.ID,title:[""编辑课程"",""font-size:16""],area:[""760px"", ""90%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""del""){" & vbCrlf
	tmpHtml = tmpHtml & "				layer.confirm('真的删除选中的课程吗？<br />相关的数据将同步删除而且无法恢复！', {icon:3,title: ""删除提醒""}, function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "Course/Delete.html"",{ItemID:data.CourseID,ID:data.ID}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						var icon = 2;" & vbCrlf
	tmpHtml = tmpHtml & "						if(reData.Return){icon = 1}" & vbCrlf
	tmpHtml = tmpHtml & "						layer.alert(reData.reMessge, {icon:icon,title: ""删除结果提示""},function(layero, index){" & vbCrlf
	tmpHtml = tmpHtml & "							if(reData.Return){ layer.close(layer.index);table.reload(""TableList""); }" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""apply""){" & vbCrlf		'申请修改
	tmpHtml = tmpHtml & "				layer.open({type:2,id:""applyWin"",content:""" & ParmPath & "Course/applyModify.html?ItemID=" & tItemID & "&ID="" + data.ID, title:[""申请修改"",""font-size:16""], area:[""630px"", ""350px""]});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""backModi""){" & vbCrlf		'退回修改
	tmpHtml = tmpHtml & "				var str1 = ""<div class=\""hr-workZones hr-shrink-x10\"">"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<fieldset class=\""layui-elem-field layui-field-title\"" style=\""margin-top:1px;\""><legend>退回课程业绩</legend></fieldset>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<div class=\""layui-form layui-form-pane\"">"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<div class=\""layui-form-item layui-form-text\""><label class=\""layui-form-label\"">退回理由：</label>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<div class=\""layui-input-block\""><textarea name=\""Explain\"" id=\""Explain\"" placeholder=\""请填写退回理由\"" class=\""layui-textarea\""></textarea></div>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""</div>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""<div class=\""searchBtn\""><button class=\""layui-btn layui-btn-sm\"" type=\""button\"" id=\""backPost\"" title=\""发送\""><i class=\""hr-icon hr-icon-top\"">&#xec58;</i>发送</button></div>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""</div>"";" & vbCrlf
	tmpHtml = tmpHtml & "				str1 += ""</div>"";" & vbCrlf
	tmpHtml = tmpHtml & "				layer.open({type:1,id:""backWin"",content:str1, title:[""退回修改"",""font-size:16""], area:[""630px"", ""350px""]});" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#backPost"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "					var tExplain = $(""#Explain"").val();" & vbCrlf
	tmpHtml = tmpHtml & "					if(tExplain == """"){layer.msg(""请输入退回理由"",{icon:0,btn:""关闭""});return false;}" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "CourseProof/backSave.html"",{ItemID:" & tItemID & ", ID:data.ID, ygdm:data.VA1, userid:" & UserID & ", Explain:tExplain}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.msg(reData.reMessge,{icon:0,btn:""关闭"",time:0},function(){layer.closeAll();table.reload(""TableList"");});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""affirm""){" & vbCrlf		'确认提交
	tmpHtml = tmpHtml & "				layer.confirm(""您确定选中的课程业绩没问题？"", {icon:3,title: ""系统提醒""}, function(index){" & vbCrlf
	tmpHtml = tmpHtml & "					$.getJSON(""" & ParmPath & "CourseProof/Affirm.html"",{ItemID:" & tItemID & ", ID:data.ID}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "						layer.alert(reData.reMessge,{icon:1,title:""确认提交""});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}else if(obj.event === ""verify""){" & vbCrlf
	tmpHtml = tmpHtml & "				$.getJSON(""" & ParmPath & "Course/Passed.html"",{ID:data.ID,ItemID:data.CourseID,type:1}, function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "					layer.msg(reData.reMessge,{time: 20000,btn:""关闭"",icon:6" & vbCrlf
	tmpHtml = tmpHtml & "						,yes:function(index, layero){layer.close(index);table.reload(""TableList"");}" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		table.on(""toolbar(TableList)"", function(obj){" & vbCrlf		'监听表头工具
	tmpHtml = tmpHtml & "			var data = table.checkStatus(obj.config.id).data;" & vbCrlf
	tmpHtml = tmpHtml & "			switch(obj.event){" & vbCrlf
	tmpHtml = tmpHtml & "				case ""addNew"":" & vbCrlf
	tmpHtml = tmpHtml & "					layer.open({type:2, content:'" & ParmPath & "Course/AddNew.html?ItemID=" & tItemID & "',title:[""添加新课程"",""font-size:16""],area:[""760px"",""90%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "					break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""batchdel"":" & vbCrlf
	tmpHtml = tmpHtml & "					if(data.length==0){layer.tips(""请选择您要删除的课程！"","".laytable-cell-checkbox"",{tips: [1, ""#F60""]});return false;}" & vbCrlf
	tmpHtml = tmpHtml & "					var arrID=[]; for(var i=0;i<data.length;i++){ arrID.push(data[i].ID); }" & vbCrlf
	tmpHtml = tmpHtml & "					layer.confirm(""确认要删除选中的 "" + data.length + "" 条课程记录？<br />删除后将无法恢复。"",{icon:3, title:[""删除警告"",""background-color:#f30""]},function(index){" & vbCrlf
	tmpHtml = tmpHtml & "						$.getJSON(""" & ParmPath & "Course/Delete.html"",{ItemID:" & tItemID & ",ID:arrID.join()}, function(reJson){" & vbCrlf
	tmpHtml = tmpHtml & "							layer.msg(reJson.reMessge,{title:""删除结果"",btn:""关闭"",time:0},function(){ table.reload(""TableList""); });" & vbCrlf
	tmpHtml = tmpHtml & "						});" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "					break;" & vbCrlf
	tmpHtml = tmpHtml & "				case ""refresh"":table.reload(""layList"");break;" & vbCrlf
	tmpHtml = tmpHtml & "			};" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot("Desktop", 1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub GetJsonList()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpYear : tmpYear = DefYear : If HR_Clng(Request("SchoolYear")) > 2000 Then tmpYear = HR_Clng(Request("SchoolYear"))
	Dim vMSG, tmpJson, tmpData, rsGet, sqlGet, isErr : isErr = False

	Dim tSoVA8 : tSoVA8 = Trim(ReplaceBadChar(Request("Course")))
	Dim tTitle : tTitle = Trim(ReplaceBadChar(Request("Title")))
	Dim soStuType : soStuType = Trim(ReplaceBadChar(Request("StuType")))

	Dim soPass : soPass = HR_CBool(Request("IsPass"))					'审核状态
	Dim soAffirm : soAffirm = HR_CBool(Request("IsAffirm"))					'是否查看已确认
	Dim soRetreat : soRetreat = HR_CBool(Request("IsRetreat"))					'是否查看已退回

	Dim soCampus : soCampus = Trim(ReplaceBadChar(Request("Campus")))		'搜索校区
	Dim arrPeriod, soStartT, soEndT, soPeriod : soPeriod = Trim(ReplaceBadChar(Request("Period")))		'时间段
	If Instr(soPeriod, "～") > 0 Then
		soPeriod = Replace(soPeriod, " ", "")
		arrPeriod = Split(soPeriod, "～")
		If Ubound(arrPeriod) = 1 Then
			If IsDate(arrPeriod(0)) Then soStartT = ConvertDateToNum(arrPeriod(0))+2
			If IsDate(arrPeriod(1)) Then soEndT = ConvertDateToNum(arrPeriod(1))+2
		End If
	End If
	Dim soSort : soSort = Trim(ReplaceBadChar(Request("soSort")))			'排序方式
	Dim hrDate : hrDate = False			'是否有时间字段

	Dim tCourse, tSheetName, tFieldLen, tTempTable, tAttach, tUnit, tKSMC, tTime, tYGXB, tPRZC
	'----- 取表结构
	Set rsTmp = Conn.Execute("Select Top 1 ClassName,Unit,SheetName,FieldLen,Template From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tCourse = rsTmp("ClassName")
			tUnit = rsTmp("Unit")
			tSheetName = Trim(rsTmp("SheetName"))
			tTempTable = Trim(rsTmp("Template"))
			tFieldLen = HR_Clng(rsTmp("FieldLen"))
			If Not(ChkTable(tSheetName)) Then	'----- 检查表是否存在
				Conn.Execute("Select * Into " & tSheetName & " From HR_" & tTempTable & " Where 1=0")
				Conn.Execute("Alter Table " & tSheetName & " Add Primary Key (ID)")		'设置主键
				Conn.Execute("Alter Table " & tSheetName & " Add Default(" & tCourseID & ") For ItemID")		'设置默认值
				Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For ID")		'设置默认值
				Conn.Execute("Alter Table " & tSheetName & " Add Default('') For StudentType")		'设置默认值
				Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For VA0")		'设置默认值
				Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For VA1")
				Conn.Execute("Alter Table " & tSheetName & " Add Default('') For VA2")
				Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For VA3")
				Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For Passed")
				Conn.Execute("Alter Table " & tSheetName & " Add Default('') For Explain")
				Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For UserID")		'设置默认值
				Conn.Execute("Alter Table " & tSheetName & " Add Default(getdate()) For AppendTime")		'设置默认值
				Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For State")
				Conn.Execute("Alter Table " & tSheetName & " Add Default('') For KSMC")
				Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For KSDM")
				Conn.Execute("Alter Table " & tSheetName & " Add Default('') For YGXB")
				Conn.Execute("Alter Table " & tSheetName & " Add Default('') For PRZC")		'设置默认值
				Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For scYear")	'设置默认值
				Conn.Execute("Alter Table " & tSheetName & " Add Default(0) For scTerm")	'设置默认值
				Conn.Execute("Delete From " & tSheetName)
			End If
		Else
			Response.Write "{""code"":400,""msg"":""数据表“" & tSheetName & "”不存在"",""count"":0,""data"":[]}"
			Exit Sub
		End If
	Set rsTmp = Nothing


	If tTempTable = "TempTableA" Or tTempTable = "TempTableC" Or tTempTable = "TempTableD" Or tTempTable = "TempTableE" Then hrDate = True

	Dim sqlWhere : sqlWhere = ""
	sqlWhere = sqlWhere & " And VA1=" & HR_Clng(UserYGDM)		'仅显示本人工号
	If soStartT > 0 And soEndT > 0 And hrDate Then		'按日期搜索
		sqlWhere = sqlWhere & " And VA4 Between " & soStartT & " And " & soEndT
	End If
	If soStuType <> "" Then sqlWhere = sqlWhere & " And StudentType='" & soStuType & "'"	'搜索学生类别

	If soAffirm Then sqlWhere = sqlWhere & " And State=1"									'查看教师已确认
	If soRetreat Then sqlWhere = sqlWhere & " And Retreat=1"								'查看退回

	If soPass Then sqlWhere = sqlWhere & " And Passed=" & HR_False							'搜索未审
	If tTempTable = "TempTableA" Then	'搜索课程或课程内容
		If tSoVA8 <> "" Then sqlWhere = sqlWhere & " And VA8 like '%" & tSoVA8 & "%'"
		If tTitle <> "" Then sqlWhere = sqlWhere & " And VA9 like'%" & tTitle & "%'"	'搜索课程内容
		If soCampus <> "" Then sqlWhere = sqlWhere & " And VA11='" & soCampus & "'"		'搜索校区
	ElseIf tTempTable = "TempTableC" Or tTempTable = "TempTableD" Or tTempTable = "TempTableE" Then
		If tTitle <> "" Then sqlWhere = sqlWhere & " And VA6 like'%" & tTitle & "%'"
	ElseIf tTempTable = "TempTableB" Or tTempTable = "TempTableF" Or tTempTable = "TempTableG" Then
		If tTitle <> "" Then sqlWhere = sqlWhere & " And VA5 like'%" & tTitle & "%'"
	End If

	sqlGet = "Select * From " & tSheetName & " Where scYear=" & tmpYear & " And ItemID=" & tItemID & sqlWhere
	
	If soSort = "xhUP" Then		'排序
		sqlGet = sqlGet & " Order By VA0 ASC"
	ElseIf soSort = "xhDown" Then
		sqlGet = sqlGet & " Order By VA0 DESC"
	ElseIf soSort = "importTimeUP" Then
		sqlGet = sqlGet & " Order By AppendTime ASC"
	ElseIf soSort = "importTimeDown" Then
		sqlGet = sqlGet & " Order By AppendTime DESC"
	Else
		sqlGet = sqlGet & " Order By VA0 DESC"
	End If
	

	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			i = 0 : CurrentPage = 1
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

			Dim tVA2, tVA3, tVA4, tVA7, stTime, enTime, tmpTime, isEdit, isDel, isPass, isBack, isAffirm, strAffirm, passStyle
			Dim tRetreat, rsYG, sumVA3
			sumVA3 = 0
			Do While Not rsGet.EOF
				isEdit = "false" : isDel = "false" : isPass = "false" : isBack = "false"
				passStyle = "pass-false"
				'Conn.Execute("alter table HR_TempTableF add KSMC nvarchar(50) DEFAULT(''),KSDM int DEFAULT(0),YGXB int DEFAULT(0),PRZC nvarchar(50) DEFAULT('')")	'建字段

				tAttach = GetCountAttach(rsGet("Explain"))		'统计附件
				tVA2 = Trim(rsGet("VA2"))
				tVA3 = HR_CDbl(rsGet("VA3"))
				sumVA3 = sumVA3 + tVA3
				tVA4 = Trim(rsGet("VA4"))

				Set rsYG = Conn.Execute("Select * From HR_Teacher Where Cast(YGDM As Int)=" & HR_Clng(rsGet("VA1")))
					If Not(rsYG.BOF And rsYG.EOF) Then
						tKSMC = Trim(rsYG("KSMC"))
						tPRZC = Trim(rsYG("PRZC"))
						tYGXB = Trim(rsYG("YGXB"))
						If Trim(rsYG("YGXM")) <> tVA2 Then tVA2 = tVA2 & "<span class=\""layui-badge-dot\""></span>"
					End If
				Set rsYG = Nothing

				If tTempTable = "TempTableA" Then
					tVA7 = Trim(rsGet("VA7"))
					tmpTime = GetPeriodTime(Trim(rsGet("VA11")), tVA7, 0)		'计算节次时间
				End If
				If tTempTable = "TempTableA" Or tTempTable = "TempTableC" Or tTempTable = "TempTableD" Or tTempTable = "TempTableE" Then
					If HR_Clng(tVA4) > 0 Then tVA4 = FormatDate(ConvertNumDate(tVA4), 2)
				End If

				If i > 0 Then tmpData = tmpData & ","
				tmpData = tmpData & "{""ID"":" & rsGet("ID") & ",""CourseID"":""" & tItemID & """,""Course"":""" & Trim(tCourse) & """,""StudentType"":""" & Trim(rsGet("StudentType")) & """,""KSMC"":""" & tKSMC & """"
				tmpData = tmpData & ",""PRZC"":""" & Trim(tPRZC) & """,""YGXB"":""" & arrSex(HR_Clng(tYGXB)) & """,""Time"":""" & Trim(tmpTime) & """"
				tmpData = tmpData & ",""VA0"":""" & HR_Clng(rsGet("VA0")) & """,""VA1"":""" & Trim(rsGet("VA1")) & """,""VA2"":""" & tVA2 & """,""VA3"":""" & FormatNumber(tVA3, 1, -1) & """,""VA4"":""" & tVA4 & """"
				For m = 8 To rsGet.Fields.Count - 11
					tmpData = tmpData & ",""" & rsGet.Fields(m).name & """:""" & HR_HTMLEncode(rsGet.Fields(m).value) & """"
				Next

				'判断员工是否核对本记录
				strAffirm = "未确认" : isAffirm = "true"
				If HR_CBool(rsGet("State")) Then strAffirm = "已确认"
				If HR_CBool(rsGet("State")) = False Then isAffirm = "false"

				'是否退回
				tRetreat = "" : isBack = "true"
				If HR_Clng(rsGet("Retreat")) = 1 Then tRetreat = "是" : isBack = "false"
				If HR_CBool(rsGet("Passed")) Then isBack = "false"

				If HR_Clng(rsGet("VA1")) = HR_Clng(UserYGDM) Then
					If HR_CBool(rsGet("Passed")) = False Then isEdit = "true" : isDel = "true"
				End If
				
				tmpData = tmpData & ",""Attach"":" & HR_Clng(tAttach) & ",""Unit"":""" & tUnit & """,""UserID"":""" & HR_Clng(rsGet("UserID")) & """,""AppendTime"":""" & FormatDate(rsGet("AppendTime"), 1) & """,""Passed"":" & LCase(CSTR(rsGet("Passed"))) & ""
				tmpData = tmpData & ",""scYear"":" & HR_Clng(rsGet("scYear")) & ",""scTerm"":" & HR_Clng(rsGet("scTerm"))
				tmpData = tmpData & ",""isEdit"":" & isEdit & ",""isDel"":" & isDel & ",""isPass"":" & isPass & ",""isBack"":" & isBack & ",""isAffirm"":" & isAffirm
				tmpData = tmpData & ",""Retreat"":""" & tRetreat & """,""Affirm"":""" & strAffirm & """,""passStyle"":""" & passStyle & """}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
	Set rsGet = Nothing

	tmpJson = "{""code"":0,""msg"":""课程查询成功！"",""count"":" & HR_Clng(TotalPut) & ",""sumVA3"":""" & sumVA3 & """,""data"":[" & tmpData
	tmpJson = tmpJson & "],""limit"":" & MaxPerPage & ",""page"":" & CurrentPage & "}"
	Response.Write tmpJson
End Sub

Sub EditBody()
	Dim xlsUrl, tItemName, tSheetName, tTemplate, tUnit, tFieldLen, tFieldHead, arrHeadKey
	Dim tTeacher, strStuType, arrStuType, tmpExtname, tAttachFile, tArrFile, tAttach
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))	'考核项序号
	Dim tmpID : tmpID = HR_Clng(Request("ID"))			'课程序号
	Dim IsModify : IsModify = False
	SubButTxt = "添加" : ErrMsg = ""

	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tSheetName = Trim(rsTmp("SheetName"))
			strStuType = Trim(rsTmp("StudentType"))
			tTemplate = Trim(rsTmp("Template"))
			tUnit = rsTmp("Unit")
			tFieldLen = HR_Clng(rsTmp("FieldLen"))
			tFieldHead = Trim(rsTmp("FieldHead"))
		Else
			ErrMsg = ErrMsg & "考核项目不存在或已删除！<br />"
			Response.Write GetErrBody(2) : Exit Sub
		End If
	Set rsTmp = Nothing

	If IsNull(tSheetName) Or Trim(tSheetName) = "" Then tSheetName = "HR_Sheet_" & tItemID

	If tFieldHead <> "" Then		'标题文字
		arrHeadKey = Split(tFieldHead, ",")
		If Ubound(arrHeadKey) <> tFieldLen-1 Then Redim Preserve arrHeadKey(tFieldLen - 1)
	Else
		Redim arrHeadKey(tFieldLen)
	End If

	If HR_IsNull(tTemplate) Then
		ErrMsg = ErrMsg & tItemName & " 数据模板未设置，请联系管理员！<br />"
		Response.Write GetErrBody(2) : Exit Sub
	End If

	If ChkTable(tSheetName) = False Then
		ErrMsg = ErrMsg & tItemName & " 数据表未建立，请联系管理员！<br />"
		Response.Write GetErrBody(2) : Exit Sub
	End If

	Dim strField, arrField : Redim arrField(tFieldLen)
	Dim stuTypeID, tStuType, tPassed

	sqlTmp = "Select * From " & tSheetName & " Where ID=" & tmpID
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			IsModify = True
			SubButTxt = "修改"
			tStuType = Trim(rsTmp("StudentType"))
			tAttachFile = Trim(rsTmp("Explain"))
			tAttachFile = FilterArrNull(tAttachFile, "|")
			tPassed = HR_CBool(rsTmp("Passed"))
			For i = 0 To tFieldLen-1
				arrField(i) = rsTmp("VA" & i)
			Next
		End If
	Set rsTmp = Nothing
	If UserYGDM <> "" And UserRank=0 Then arrField(1) = UserYGDM
	If UserYGXM <> "" And UserRank=0 Then arrField(2) = UserYGXM
	If HR_IsNull(tAttachFile) = False Then		'取附件及图标
		tArrFile = Split(tAttachFile, "|")
		For i = 0 To Ubound(tArrFile)
			tmpExtname = Right(Trim(tArrFile(i)), Len(Trim(tArrFile(i))) - inStr(Trim(tArrFile(i)), "."))
			tAttach = tAttach & "<em class=""fileItem""><span title=""" & Trim(tArrFile(i)) & """><i class=""hr-icon"">" & GetAttachIcon(tmpExtname) & "</i></span><tt>删除</tt></em>"
		Next
	End If
	
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding: 5px 8px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .width_80 {width:80px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.tipTxt {padding-right:5px;cursor: pointer;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-form-item .attachBox {width:auto;}" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-inline .tips {padding-left:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-slider {flex-grow:1;}" & vbCrlf
	tmpHtml = tmpHtml & "		.slider {box-sizing:border-box;padding:1px 5px 0 5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		#AttachBar {line-height:37px;min-height:38px;display:flex;align-items:center;flex-wrap:wrap;width:390px;border:1px solid #ddd;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		#AttachBar em {min-height:60px; line-height:50px; cursor: pointer; padding:15px 0 0 15px;color:#39c;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		#AttachBar em i {font-size:46px;position:relative;top:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		#AttachBar em tt {display:none;}" & vbCrlf

	tmpHtml = tmpHtml & "		.listBox {line-height:25px;display:flex;align-items:center;flex-wrap:wrap;}" & vbCrlf
	tmpHtml = tmpHtml & "		.listBox em {line-height:25px; cursor: pointer; padding:10px 0 0 10px;color:#39c;box-sizing: border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.listBox em:hover {color:#900;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead("Desktop", 1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	tmpHtml = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	tmpHtml = tmpHtml & "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:5px;"">" & vbCrlf
	tmpHtml = tmpHtml & "<legend>" & SubButTxt & "课程【" & tItemName & "】</legend>"
	tmpHtml = tmpHtml & "<form class=""layui-form layui-form-pane"" id=""FloatForm"" name=""FloatForm"" lay-filter=""FloatForm"" action="""">" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""layer-hr-box"" id=""xlsBox"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">教　师：</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline"">"
	tmpHtml = tmpHtml & "<input type=""text"" name=""YGXM"" id=""soYGXM"" value=""" & UserYGXM & """ lay-verify=""required"" autocomplete=""on"" title=""请点击选择按钮查找教师"" class=""layui-input txt1"" readonly>"
	tmpHtml = tmpHtml & "			</div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf

	tmpHtml = tmpHtml & "		<div class=""layui-inline"">" & vbCrlf
	tmpHtml = tmpHtml & "			<label class=""layui-form-label"">工　　号：</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><input type=""text"" name=""YGDM"" id=""soYGDM"" lay-verify=""required"" value=""" & UserYGDM & """ class=""layui-input txt1"" readonly></div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	If strStuType <> "" Then
		tmpHtml = tmpHtml & "	<div class=""layui-form-item"" pane><label class=""layui-form-label"">学生类别：</label><div class=""layui-input-block"">" & vbCrlf
		strStuType = FilterArrNull(strStuType, ",") : arrStuType = Split(strStuType, ",")
		For i = 0 To Ubound(arrStuType)
			tmpHtml = tmpHtml & "<input type=""radio"" name=""StudentType"" value=""" & Trim(arrStuType(i)) & """ title=""" & arrStuType(i) & """ lay-skin=""primary"""
			If tStuType = arrStuType(i) Then
				tmpHtml = tmpHtml & " checked"
			Else
				If tmpID>0 And IsModify Then tmpHtml = tmpHtml & " disabled"
			End If
			tmpHtml = tmpHtml & ">"
		Next
		tmpHtml = tmpHtml & "</div></div>" & vbCrlf
	End If

	Dim tVA4 : tVA4 = Trim(arrField(4))
	Dim dateType : dateType = False
	If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Then
		tVA4 = FormatDate(ConvertNumDate(tVA4), 2)			'转为日期
		dateType = True
	End If

	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">" & arrHeadKey(3) & "</label>" & vbCrlf	'学时
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline width_80""><input type=""text"" name=""VA3"" value=""" & arrField(3) & """ lay-verify=""number"" autocomplete=""on"" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-form-mid"">" & tUnit & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">" & arrHeadKey(4) & "</label>" & vbCrlf	'日期
	If dateType Then
		tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><input type=""text"" name=""VA4"" value=""" & tVA4 & """ id=""date1"" lay-verify=""date"" class=""layui-input""></div>" & vbCrlf
	Else
		tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><select name=""VA4"" title=""VA4"">" & GetSemesterOption(0, tVA4) & "</select></div>" & vbCrlf
	End If
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf

	If tTemplate = "TempTableB" Then
		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(5) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-block""><input type=""text"" name=""VA5"" value=""" & arrField(5) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item layui-form-text"">"
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(6) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-block""><textarea name=""VA6"" id=""VA6"" placeholder=""" & arrHeadKey(6) & """ lay-verify=""content"" class=""layui-textarea"">" & arrField(6) & "</textarea></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf
		
	ElseIf tTemplate = "TempTableC" Then
		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(5) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA5"" value=""" & arrField(5) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA5"" title=""VA5"" lay-search=""""><option value="""">搜索/选择</option>" & GetSemesterOption(2, "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(6) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA6"" value=""" & arrField(6) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA6"" title=""VA6"" lay-search=""""><option value="""">选择/搜索</option>" & GetFieldOption(tSheetName, "VA6", "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item layui-form-text"">"
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(7) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-block""><textarea name=""VA7"" id=""VA7"" placeholder=""备注"" lay-verify=""content"" class=""layui-textarea"">" & arrField(7) & "</textarea></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf
	ElseIf tTemplate = "TempTableD" Then
		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(5) & "</label>" & vbCrlf		'学期
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA5"" value=""" & arrField(5) & """ lay-verify=""required"" autocomplete=""off"" class=""layui-input txt1"" readonly></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA5"" title=""VA5"" lay-search=""""><option value="""">选择学期</option>" & GetSemesterOption(2, "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf		'项目名称
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(6) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA6"" value=""" & arrField(6) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA6"" title=""VA6"" lay-search=""""><option value="""">选择/搜索</option>" & GetFieldOption(tSheetName, "VA6", "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(7) & "</label>" & vbCrlf		'级别
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA7"" value=""" & arrField(7) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA7"" title=""VA7"" lay-search=""""><option value="""">选择/搜索</option>" & GetSubmoduleOption(tItemID, arrField(7)) & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item layui-form-text"">"
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(8) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-block""><textarea name=""VA8"" id=""VA8"" placeholder=""备注"" lay-verify=""content"" class=""layui-textarea"">" & arrField(8) & "</textarea></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf
		
	ElseIf tTemplate = "TempTableE" Then
		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(5) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA5"" value=""" & arrField(5) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA5"" title=""VA5""><option value="""">选择学期</option>" & GetSemesterOption(2, "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(6) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA6"" value=""" & arrField(6) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA6"" title=""VA6"" lay-search=""""><option value="""">选择/搜索</option>" & GetFieldOption(tSheetName, "VA6", "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">" & arrHeadKey(7) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><select name=""VA7"" title=""VA7"" id=""LevelMenu"" lay-filter=""LevelMenu""><option value="""">选择级别</option>" & GetSubmoduleOption(tItemID, arrField(7)) & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "		</div>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">" & arrHeadKey(8) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "			<div class=""layui-input-inline"" id=""GradeBox""><select name=""VA8"" title=""VA8"" id=""GradeMenu"" lay-filter=""GradeMenu""><option value="""">无等级</option>" & GetItemGradeOption(tItemID, arrField(7), arrField(8)) & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "		</div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item layui-form-text"">"
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(9) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-block""><textarea name=""VA9"" id=""VA9"" placeholder=""备注"" lay-verify=""content"" class=""layui-textarea"">" & arrField(9) & "</textarea></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

	ElseIf tTemplate = "TempTableF" Then
		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(5) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA5"" value=""" & arrField(5) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA5"" title=""VA5"" lay-search=""""><option value="""">搜索/选择</option>" & GetFieldOption(tSheetName, "VA5", "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">" & arrHeadKey(6) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "			<div class=""layui-input-inline""><select name=""VA6"" title=""VA6"" id=""LevelMenu"" lay-filter=""LevelMenu""><option value="""">选择级别</option>" & GetSubmoduleOption(tItemID, arrField(6)) & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "		</div>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">" & arrHeadKey(7) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "			<div class=""layui-input-inline"" id=""GradeBox""><select name=""VA7"" title=""VA7"" id=""GradeMenu"" lay-filter=""GradeMenu""><option value="""">无等级</option>" & GetItemGradeOption(tItemID, arrField(6), arrField(7)) & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "		</div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item layui-form-text"">"
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(8) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-block""><textarea name=""VA8"" id=""VA8"" placeholder=""备注"" lay-verify=""content"" class=""layui-textarea"">" & arrField(8) & "</textarea></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf
	ElseIf tTemplate = "TempTableG" Then
		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(5) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA5"" value=""" & arrField(5) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA5"" title=""VA5"" lay-search=""""><option value="""">搜索/选择</option>" & GetFieldOption(tSheetName, "VA5", "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(6) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA6"" value=""" & arrField(6) & """ autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA6"" title=""VA6""><option value="""">选择" & arrHeadKey(6) & "</option>" & GetSubmoduleOption(tItemID, arrField(6)) & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item layui-form-text"">"
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(7) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-block""><textarea name=""VA7"" id=""VA7"" placeholder=""备注"" lay-verify=""content"" class=""layui-textarea"">" & arrField(7) & "</textarea></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf
	ElseIf tTemplate = "TempTableA" Then
		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">" & arrHeadKey(5) & "</label>" & vbCrlf		'周次
		tmpHtml = tmpHtml & "			<div class=""layui-input-inline width_80""><input type=""text"" name=""VA5"" value=""" & arrField(5) & """ lay-verify=""number"" autocomplete=""on"" class=""layui-input""></div>" & vbCrlf
		tmpHtml = tmpHtml & "			<div class=""layui-form-mid"">周　</div>" & vbCrlf
		tmpHtml = tmpHtml & "		</div>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">" & arrHeadKey(6) & "</label>" & vbCrlf		'星期
		tmpHtml = tmpHtml & "			<div class=""layui-input-inline width_80""><input type=""text"" id=""week"" name=""VA6"" value=""" & arrField(6) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input""></div>" & vbCrlf
		tmpHtml = tmpHtml & "		</div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf		'课程
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(8) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA8"" value=""" & arrField(8) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA8"" title=""VA8"" lay-search=""""><option value="""">选择/搜索" & arrHeadKey(8) & "</option>" & GetCourseOption(arrField(8), 0) & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf		'授课内容
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(9) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA9"" value=""" & arrField(9) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>"
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA9"" title=""VA9"" lay-search=""""><option value="""">选择/搜索" & arrHeadKey(9) & "</option>" & GetFieldOption(tSheetName, "VA9", "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf		'授课对象
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(10) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA10"" value=""" & arrField(10) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA6"" title=""VA10"" lay-search=""""><option value="""">选择/搜索" & arrHeadKey(10) & "</option>" & GetFieldOption(tSheetName, "VA10", "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf		'校(院)区【与节次联动】
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(11) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA11"" value=""" & arrField(11) & """ lay-verify=""required"" autocomplete=""on"" class=""layui-input txt1""></div>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA11"" title=""VA11"" id=""CampusMenu"" lay-filter=""CampusMenu"" lay-search=""""><option value="""">选择/搜索" & arrHeadKey(11) & "</option>" & GetCampusOption(arrField(11), 0) & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item hr-rows"">" & vbCrlf		'节次
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(7) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline width_80""><input type=""text"" name=""VA7"" id=""VA7"" value=""" & arrField(7) & """ lay-verify=""required"" autocomplete=""off"" class=""layui-input txt1""></div>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""hr-slider""><div id=""slide7"" class=""slider""></div></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf

		tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf		'上课地点
		tmpHtml = tmpHtml & "		<label class=""layui-form-label"">" & arrHeadKey(12) & "</label>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""VA12"" value=""" & arrField(12) & """ autocomplete=""on"" class=""layui-input txt1""></div>" & vbCrlf
		tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""SelectVA12"" title=""VA12"" lay-search=""""><option value="""">选择/搜索" & arrHeadKey(12) & "</option>" & GetFieldOption(tSheetName, "VA12", "") & "</select></div>" & vbCrlf
		tmpHtml = tmpHtml & "	</div>" & vbCrlf
	End If
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf		'附件
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><label class=""layui-form-label"">附件</label>" & vbCrlf
	tmpHtml = tmpHtml & "			<div class=""layui-input-inline attachBox""><div id=""AttachBar"">" & tAttach & "</div></div>" & vbCrlf
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "		<input type=""hidden"" name=""UploadAttach"" id=""UploadAttach"" value=""" & tAttachFile & """>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><button type=""button"" class=""layui-btn"" id=""UploadBtn"">添加附件</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-item"">" & vbCrlf		'附件说明
	tmpHtml = tmpHtml & "		<div class=""layui-inline"">1、上传的附件最大字节不要超过2M；<br>2、上传论文封面、目录、文章；</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf

	tmpHtml = tmpHtml & "	<input type=""hidden"" name=""SheetName"" value=""" & tSheetName & """><input type=""hidden"" name=""ItemID"" value=""" & tItemID & """><input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""ItemName"" value=""" & tItemName & """>" & vbCrlf
	If IsModify Then tmpHtml = tmpHtml & "	<input type=""hidden"" name=""Modify"" value=""True"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""hr-pop-fix"">" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-inline""><button type=""button"" class=""layui-btn layui-btn-sm"" id=""FloatPost"">" & SubButTxt & "</button><button type=""reset"" class=""layui-btn layui-btn-sm layui-btn-primary"">重置</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</form>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""xlsData"" id=""xlsData""></div>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""hr-shrink-x20""></div>" & vbCrlf
	Response.Write tmpHtml

	tmpHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form"", ""laydate"",""upload"", ""element"", ""slider""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, laydate = layui.laydate, upload = layui.upload, slider = layui.slider;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	
	If tTemplate = "TempTableA" Then
		tmpHtml = tmpHtml & "		var val7 = $(""#VA7"").val(), arrVal=[3,5];if(val7!=""""){arrVal=val7.split(""-"")}" & vbCrlf
		tmpHtml = tmpHtml & "		var slider7 = slider.render({" & vbCrlf
		tmpHtml = tmpHtml & "			elem:""#slide7"",range: true,max: 20,theme:""#809"",value:[arrVal[0],arrVal[1]]," & vbCrlf
		tmpHtml = tmpHtml & "			change: function(value){$(""#VA7"").val(value[0] + ""-"" + value[1])}" & vbCrlf
		tmpHtml = tmpHtml & "		});" & vbCrlf
	End If
	tmpHtml = tmpHtml & "		$(""#FloatPost"").on(""click"", function(){" & vbCrlf
	'tmpHtml = tmpHtml & "			console.log($(""#FloatForm"").serialize());return false;" & vbCrlf
	tmpHtml = tmpHtml & "			$.post(""" & ParmPath & "Course/SaveEdit.html"", $(""#FloatForm"").serialize(), function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "				var reData = eval(""("" + rsStr + "")""),icon=2;" & vbCrlf
	tmpHtml = tmpHtml & "				if(reData.Return){icon = 1}" & vbCrlf
	tmpHtml = tmpHtml & "				layer.alert(reData.reMessge, {icon:icon},function(){" & vbCrlf
	tmpHtml = tmpHtml & "					if(reData.Return){" & vbCrlf
	tmpHtml = tmpHtml & "						parent.layer.closeAll(); " & vbCrlf
	tmpHtml = tmpHtml & "						parent.layui.table.reload(""TableList"");" & vbCrlf		'重构列表
	tmpHtml = tmpHtml & "					}else{layer.close(layer.index);}" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "			return false;" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		upload.render({" & vbCrlf		'附件上传
	tmpHtml = tmpHtml & "			elem: '#UploadBtn',url: '" & InstallDir & "API/UploadFile.htm?UploadDir=Attach', accept:'file'" & vbCrlf
	tmpHtml = tmpHtml & "			,multiple: true,exts:'zip|rar|jpg|jpeg|png|doc|docx|xls|xlsx|gif|pdf|txt'" & vbCrlf
	tmpHtml = tmpHtml & "			,done: function(res, index){" & vbCrlf		'//上传完毕
	tmpHtml = tmpHtml & "				console.log(res.data);$(""#AttachBar"").append(UpfileToIcon(res.data.src));" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#UploadAttach"").val($(""#UploadAttach"").val() + ""|"" + res.data.src);" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#AttachBar em span"").on(""click"",function(){" & vbCrlf		'预览附件
	tmpHtml = tmpHtml & "					parent.layer.open({type:2,content:""" & ParmPath & "Course/viewAttach.html?url="" + $(this).attr(""title""),title:[""预览附件"",""font-size:16""],area:[""80%"", ""86%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "			,error: function (index, upload){console.log(index);}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#AttachBar em"").on(""click"",function(){" & vbCrlf		'预览附件
	tmpHtml = tmpHtml & "			alert($(this).attr(""title""));" & vbCrlf
	tmpHtml = tmpHtml & "			parent.layer.open({type:2,content:""" & ParmPath & "Course/viewAttach.html?url="" + $(this).attr(""title""),title:[""预览附件"",""font-size:16""],area:[""80%"", ""86%""],maxmin:true});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		var levelData=[];" & vbCrlf			'准备级别和等级数据【等级联动】
	tmpHtml = tmpHtml & "		$.getJSON(""" & ParmPath & "Course/levelData.html"",{item:" & tItemID & "}, function(rsStr){ levelData=rsStr.data;});" & vbCrlf		'异步取值
	tmpHtml = tmpHtml & "		form.on(""select(LevelMenu)"", function(data){" & vbCrlf			'监听级别下拉
	tmpHtml = tmpHtml & "			for(var i=0;i<levelData.length;i++){" & vbCrlf
	tmpHtml = tmpHtml & "				if(levelData[i].LevelName==data.value){;" & vbCrlf			'准备更新等级下拉
	tmpHtml = tmpHtml & "					var GradeOption=""<select name=\""VA8\"" id=\""GradeMenu\"" lay-filter=\""GradeMenu\"">"";" & vbCrlf
	tmpHtml = tmpHtml & "					for(var j=0;j<levelData[i].Grade.length;j++){" & vbCrlf
	tmpHtml = tmpHtml & "						GradeOption +=""<option value=\"""" + levelData[i].Grade[j].Grade +""\"">"" + levelData[i].Grade[j].Grade +""</option>"";" & vbCrlf
	tmpHtml = tmpHtml & "					}" & vbCrlf
	tmpHtml = tmpHtml & "					GradeOption +=""</select>"";" & vbCrlf
	tmpHtml = tmpHtml & "					$(""#GradeBox"").html(GradeOption);" & vbCrlf
	tmpHtml = tmpHtml & "					form.render(""select"");" & vbCrlf		'刷新等级下拉
	tmpHtml = tmpHtml & "				}" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		var CampusData=[];" & vbCrlf			'准备校区和节次数据【联动】
	tmpHtml = tmpHtml & "		$.getJSON(""" & ParmPath & "Course/CampusData.html"", function(rsStr){ CampusData=rsStr.data;});" & vbCrlf		'异步取值
	tmpHtml = tmpHtml & "		form.on(""select(CampusMenu)"", function(data){" & vbCrlf			'监听级别下拉
	tmpHtml = tmpHtml & "			for(var i=0; i<CampusData.length; i++){" & vbCrlf
	tmpHtml = tmpHtml & "				if(CampusData[i].Campus == data.value){;" & vbCrlf			'准备更新等级下拉
	tmpHtml = tmpHtml & "					var PeriodOption=""<select name=\""SelectVA7\"" id=\""PeriodMenu\"" title=\""VA7\"" lay-filter=\""PeriodMenu\"">"";" & vbCrlf
	tmpHtml = tmpHtml & "					for(var j=0;j<CampusData[i].Items.length;j++){" & vbCrlf
	tmpHtml = tmpHtml & "						PeriodOption +=""<option value=\"""" + CampusData[i].Items[j].Period +""\"">"" + CampusData[i].Items[j].Period +""</option>"";" & vbCrlf
	tmpHtml = tmpHtml & "					}" & vbCrlf
	tmpHtml = tmpHtml & "					PeriodOption +=""</select>"";" & vbCrlf
	tmpHtml = tmpHtml & "					$(""#PeriodBox"").html(PeriodOption);" & vbCrlf
	tmpHtml = tmpHtml & "					form.render(""select"");" & vbCrlf		'刷新等级下拉
	tmpHtml = tmpHtml & "				}" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf

	tmpHtml = tmpHtml & "		$(""#soYGXM"").bind(""input propertychange"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			var soKey = $(this).val(), tipsContent = """", that = $(this);" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "Teacher/SearchData.html"",{soWord:soKey}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "				for(var i = 0; i<rsStr.data.length; i++){" & vbCrlf
	tmpHtml = tmpHtml & "					tipsContent += ""<span class='tipTxt' title='"" + rsStr.data[i].YGDM + ""["" + rsStr.data[i].KSMC + ""]' name='"" + rsStr.data[i].YGDM + ""'>"" + rsStr.data[i].YGXM + ""</span>"";" & vbCrlf
	tmpHtml = tmpHtml & "				}" & vbCrlf
	tmpHtml = tmpHtml & "				layer.tips(tipsContent, that,{time:0,tips: [3, ""#F60""]});" & vbCrlf
	tmpHtml = tmpHtml & "				$("".tipTxt"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "					$(""#soYGXM"").val($(this).text());$(""#soYGDM"").val($(this).attr(""name""));layer.closeAll('tips');" & vbCrlf
	tmpHtml = tmpHtml & "				});" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		form.on(""select"", function(data){" & vbCrlf
	tmpHtml = tmpHtml & "			var el = data.elem.title;" & vbCrlf
	tmpHtml = tmpHtml & "			$("".txt1"").each(function(){" & vbCrlf
	tmpHtml = tmpHtml & "				if($(this).attr(""name"")==el){$(this).val(data.value)};" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		laydate.render({elem: ""#date1"",value:"""",done:function(value, date, endDate){" & vbCrlf
	tmpHtml = tmpHtml & "				var today = new Array('日','一','二','三','四','五','六'), day = new Date(value);" & vbCrlf
	tmpHtml = tmpHtml & "				var week = today[day.getDay()];$(""#week"").val(week);" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	tmpHtml = tmpHtml & "	function UpfileToIcon(fUrl){" & vbCrlf
	tmpHtml = tmpHtml & "		if(fUrl){" & vbCrlf
	tmpHtml = tmpHtml & "			var extName = fUrl.substr(fUrl.lastIndexOf("".""));extName =extName.replace(""."","""");" & vbCrlf
	tmpHtml = tmpHtml & "			switch(extName){" & vbCrlf
	tmpHtml = tmpHtml & "				case ""jpg"": case ""jpeg"":case ""png"":case ""bmp"":case ""gif"": return ""<em class=\""fileItem\""><span title=\"""" + fUrl + ""\""><i class=\""hr-icon\"">&#xf1c5;</i></span><tt>删除</tt></em>"";" & vbCrlf
	tmpHtml = tmpHtml & "				case ""xls"":case ""xlsx"": return ""<em class=\""fileItem\""><span title=\"""" + fUrl + ""\""><i class=\""hr-icon\"">&#xf1c3;</i></span><tt>删除</tt></em>"";" & vbCrlf
	tmpHtml = tmpHtml & "				case ""pdf"": return ""<em class=\""fileItem\""><span title=\"""" + fUrl + ""\""><i class=\""hr-icon\"">&#xf1c1;</i></span><tt>删除</tt></em>"";" & vbCrlf
	tmpHtml = tmpHtml & "				case ""doc"":case ""docx"": return ""<em class=\""fileItem\""><span title=\"""" + fUrl + ""\""><i class=\""hr-icon\"">&#xf1c2;</i></span><tt>删除</tt></em>"";" & vbCrlf
	tmpHtml = tmpHtml & "				case ""txt"": return ""<em class=\""fileItem\""><span title=\"""" + fUrl + ""\""><i class=\""hr-icon\"">&#xf0f6;</i></span><tt>删除</tt></em>"";" & vbCrlf
	tmpHtml = tmpHtml & "				case ""rar"":case ""zip"": return ""<em class=\""fileItem\""><span title=\"""" + fUrl + ""\""><i class=\""hr-icon\"">&#xec1c;</i></span><tt>删除</tt></em>"";" & vbCrlf
	tmpHtml = tmpHtml & "				default: return ""<em class=\""fileItem\""><span title=\"""" + fUrl + ""\""><i class=\""hr-icon\"">&#xec15;</i></span><tt>删除</tt></em>"";" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		}else{return fUrl;}" & vbCrlf
	tmpHtml = tmpHtml & "	}" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot("Desktop", 1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub SaveEdit()
	Dim tmpJson, tItemName, tSheetName, tTempTable, numField : numField = 13
	Dim tFieldID, tUnit, tYGXM, tKSMC, tKSDM, tYGXB, tXZZW, tPRZC
	Dim rsAdd, sqlAdd, tData, iCount
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim TempID : TempID = HR_Clng(Request("ID"))
	Dim IsModify : IsModify = HR_CBool(Request("Modify"))
	Dim tYGDM : tYGDM = HR_Clng(Request("YGDM"))
	Dim StudentType : StudentType = Trim(ReplaceBadChar(Request("StudentType")))
	Dim tAttach : tAttach = Trim(ReplaceBadUrl(Request("UploadAttach")))
	Dim subModule : subModule = Trim(Request("Submodule"))
	ErrMsg = "" : SubButTxt = "添加"
	If tAttach <> "" Then tAttach = FilterArrNull(tAttach, "|")
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassID=" & tItemID)
		If rsTmp.BOF And rsTmp.EOF Then
			ErrMsg = "业绩项目 " & tItemID & "不存在！"
		Else
			tItemName = Trim(rsTmp("ClassName"))
			tSheetName = Trim(rsTmp("SheetName"))
			tTempTable = Trim(rsTmp("Template"))
			tUnit = rsTmp("Unit")
			numField = HR_Clng(rsTmp("FieldLen"))
		End If
	Set rsTmp = Nothing

	If Not(ChkTable(tSheetName)) Then ErrMsg = ErrMsg & "数据表 " & tSheetName & " 不存在！<br>"
	iCount = 0
	
	Set rsTmp = Conn.Execute("Select * From HR_Teacher Where YGDM='" & tYGDM & "'")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tYGDM = rsTmp("YGDM")
			tYGXM = rsTmp("YGXM")
			tKSMC = rsTmp("KSMC")
			tKSDM = rsTmp("KSDM")
			tYGXB = rsTmp("YGXB")
			tXZZW = rsTmp("XZZW")
			tPRZC = rsTmp("PRZC")
		Else
			ErrMsg = ErrMsg & "未选择教师或教师不存在[" & tYGDM & "]！<br>"
		End If
	Set rsTmp = Nothing
	If ErrMsg <> "" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}"
		Exit Sub
	End If
	
	If UserYGDM <> "" And UserRank=0 Then tYGDM = UserYGDM		'非管理员限制仅添加本人
	If UserYGXM <> "" And UserRank=0 Then tYGXM = UserYGXM

	ErrMsg = ""
	Dim SentMsg, tVA3, tVA4, tmpYear, scTermUP, tUpKPI, tIsDate :tIsDate = False

	tVA3 = HR_CDbl(Request("VA3"))		'学时
	tVA4 = Trim(Request("VA4"))		'日期/学年
	tmpYear = GetSchoolYear(tVA4, 2)		'取学年年度
	scTermUP = GetSchoolYear(tVA4, 3)		'取【1为上学期2下学期】

	If tTempTable = "TempTableA" Or tTempTable = "TempTableC" Or tTempTable = "TempTableD" Or tTempTable = "TempTableE" Then
		If HR_IsNull(tVA4) Then
			ErrMsg = "日期还没有填写"
		ElseIf IsDate(tVA4) Then
			tVA4 = ConvertDateToNum(tVA4) + 2		'处理时间戳误差(非导入时转时间戳必须减2)
			tIsDate = True
		Else
			ErrMsg = "日期格式不正确！【" & tVA4 & "】"
		End If
	Else
		If HR_IsNull(tVA4) Then ErrMsg = "学期（学年）还没有填写！"
	End If
	sqlAdd = "Select * From " & tSheetName & " Where ItemID=" & tItemID & " And ID=" & TempID
	If StudentType <> "" Then sqlAdd = sqlAdd & " And StudentType='" & StudentType & "'"
	'判断数据是否重复
	Dim rsChk, sqlChk
	sqlChk = "Select * From " & tSheetName & " Where ItemID=" & tItemID & ""
	If HR_IsNull(StudentType) = False Then sqlChk = sqlChk & " And StudentType='" & Trim(StudentType) & "'"
	sqlChk = sqlChk & " And VA1=" & HR_Clng(tYGDM) & " And VA2='" & Trim(tYGXM) & "' And VA3=" & tVA3		'判断工号、姓名、学时值
	If HR_IsNull(tVA4) = False Then		'判断日期/学年
		If tIsDate Then
			sqlChk = sqlChk & " And VA4=" & tVA4
		Else
			sqlChk = sqlChk & " And VA4='" & tVA4 & "'"
		End If
	End If
	Select Case tTempTable
		Case "TempTableA"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA5")) & "' And VA7='" & Trim(Request("VA7")) & "' And VA8='" & Trim(Request("VA8")) & "'"		'判断周次、节次、课程名称
			If HR_IsNull(Request("VA9")) = False Then sqlChk = sqlChk & " And VA9='" & Trim(Request("VA9")) & "'"		'判断内容
			sqlChk = sqlChk & " And VA11='" & Trim(Request("VA11")) & "'"		'判断校区
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA7")) Or HR_IsNull(Request("VA8")) Or HR_IsNull(Request("VA11")) Then
				ErrMsg = ErrMsg & "周次、节次、课程名称、校区都不能为空！<br>"
			End If
		Case "TempTableB"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA5")) & "'"		'判断项目名称
			If Trim(Request("VA6")) <> "" Then sqlChk = sqlChk & " And Cast(VA6 As nvarchar)='" & Trim(Request("VA6")) & "'"
			If HR_IsNull(Request("VA5")) Then
				ErrMsg = ErrMsg & "项目名称不能为空！<br>"
			End If
		Case "TempTableC"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA5")) & "' And VA6='" & Trim(Request("VA6")) & "'"
			If Trim(Request("VA7")) <> "" Then sqlChk = sqlChk & " And Cast(VA7 As nvarchar)='" & Trim(Request("VA7")) & "'"		'判断工号、姓名、学期、项目名称、备注
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA6")) Then
				ErrMsg = ErrMsg & "学年(学期)或项目名称不能为空！<br>"
			End If
		Case "TempTableD"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA5")) & "' And VA6='" & Trim(Request("VA6")) & "' And VA7='" & Trim(Request("VA7")) & "'"		'判断学期、教材、级别
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA6"))  Or HR_IsNull(Request("VA7")) Then
				ErrMsg = ErrMsg & "学期、教材名称或级别不能为空！<br>"
			End If
		Case "TempTableE"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA5")) & "' And VA6='" & Trim(Request("VA6")) & "' And VA7='" & Trim(Request("VA7")) & "'"		'判断学期、项目名称、级别
			If Trim(Request("VA8")) <> "" Then sqlChk = sqlChk & " And VA8='" & Trim(Request("VA8")) & "'"	'判断等级
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA6")) Or HR_IsNull(Request("VA7")) Then
				ErrMsg = ErrMsg & "学年(学期)、项目名称及级别不能为空！<br>"
			End If
		Case "TempTableF"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA5")) & "' And VA6='" & Trim(Request("VA5")) & "'"		'判断项目名称、级别
			If Trim(Request("VA7")) <> "" Then sqlChk = sqlChk & " And VA7='" & Trim(Request("VA7")) & "'"	'判断等级
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA6")) Then
				ErrMsg = ErrMsg & "项目名称或级别不能为空！<br>"
			End If
		Case "TempTableG"
			sqlChk = sqlChk & " And VA5='" & Trim(Request("VA5")) & "' And VA6='" & Trim(Request("VA6")) & "' And Cast(VA7 as nvarchar(255))='" & Trim(Request("VA7")) & "'"		'判断学期、项目名称、级别
			If HR_IsNull(Request("VA5")) Or HR_IsNull(Request("VA6")) Then
				ErrMsg = ErrMsg & "学年(学期)、项目名称不能为空！<br>"
			End If
		Case Else
			ErrMsg = ErrMsg & "您填写的数据与系统所有模型都不匹配！“" & tItemName & "”": ChkPass = False
	End Select
	If TempID > 0 Then sqlChk = sqlChk & " And ID Not In(" & TempID & ")"
	Set rsAdd = Conn.Execute(sqlChk)
		If Not(rsAdd.BOF And rsAdd.EOF) Then
			ErrMsg = ErrMsg & "教师 " & tYGXM & "[" & tYGDM & "] 的该条数据已经存在，不能重复添加！<br>"
		End If
	Set rsAdd = Nothing
	If ErrMsg <> "" Then
		Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """,""ReStr"":""操作失败！""}"
		Exit Sub
	End If
	Dim tPXXH
	Set rsAdd = Server.CreateObject("ADODB.RecordSet")
		rsAdd.Open(sqlAdd), Conn, 1, 3
		If rsAdd.BOF And rsAdd.EOF Then
			rsAdd.AddNew
			TempID = GetNewID(tSheetName, "ID")
			tPXXH = GetNewID(tSheetName, "VA0")
			rsAdd("ID") = TempID
			rsAdd("ItemID") = tItemID
			rsAdd("StudentType") = StudentType
			rsAdd("UserID") = UserID
			rsAdd("AppendTime") = Now()
			rsAdd("VA0") = tPXXH
			rsAdd("KSMC") = tKSMC
			rsAdd("KSDM") = tKSDM
			rsAdd("YGXB") = tYGXB
			rsAdd("PRZC") = tPRZC
			rsAdd("State") = 0						'添加时状态0
		Else
			IsModify = True
			SubButTxt = "修改"
			If HR_Clng(rsAdd("UserID")) = 0 Then rsAdd("UserID") = UserID
			tPXXH = rsAdd("VA0")
			If UserYGXM <> "" And UserID=0 Then	rsAdd("Retreat") = 0						'员工修改时重置退回状态
		End If
		rsAdd("VA1") = HR_Clng(UserYGDM)
		rsAdd("VA2") = Trim(UserYGXM)
		rsAdd("VA3") = tVA3
		rsAdd("VA4") = tVA4
		rsAdd("scYear") = tmpYear
		rsAdd("scTerm") = scTermUP
		rsAdd("Passed") = False
		
		For i = 5 To numField-1
			rsAdd("VA" & i) = Trim(Request("VA" & i))
		Next
		rsAdd("Explain") = tAttach	'保存附件
		rsAdd.Update

	Set rsAdd = Nothing

	tUpKPI = UpdateKPIField()		'此处更新业绩表字段
	tUpKPI = ChkTeacherKPI(tYGDM)	'添加员工信息至业绩表
	tUpKPI = UpdateTeacherKPI(tItemID, tYGDM, StudentType)	'更新本项目员工统计数据
	tUpKPI = UpdateTeacherTotalKPI(tYGDM)	'更新员工总计数据
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""教师 " & Trim(tYGXM) & " 课程进度" & SubButTxt & "成功！"",""ReStr"":""操作成功！""}"
	Response.Write tmpJson
End Sub

Function GetCourseTitleOption(tItemID, fTitle)
	on error resume next
	Dim rsFun, strFun, tSheetName
	If HR_Clng(tItemID) > 0 Then tSheetName = "HR_Sheet_" & HR_Clng(tItemID)
	Set rsFun = Server.CreateObject("ADODB.RecordSet")
		rsFun.open "Select VA5 From " & tSheetName & " Group By VA5", Conn, 1, 1
		If Not Err.Number=0 Then Err.Clear :Exit Function
		If Not(rsFun.BOF And rsFun.EOF) Then
			Do While Not rsFun.EOF
				strFun = strFun & "<option value=""" & rsFun("VA5") & """>" & rsFun("VA5") & "</option>"
				rsFun.MoveNext
			Loop
		End If
		rsFun.Close
	Set rsFun = Nothing 
	GetCourseTitleOption = strFun
End Function

Function GetPeriodTime(fCampus, fPeriod, fType)
	Dim strFun, rsFun, fArr, strArr, stTime, enTime
	fCampus = Trim(fCampus) : fPeriod = Trim(fPeriod)
	If fCampus <> "" And fPeriod <> "" Then
		If HR_Clng(fPeriod) > 0 Then
			Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fPeriod))
				If Not(rsFun.BOF And rsFun.EOF) Then
					strFun = Trim(rsFun("StartTime")) & " - " & Trim(rsFun("EndTime"))
				End If
			Set rsFun = Nothing
		ElseIf Instr(fPeriod, "-") Then
			fArr = Split(fPeriod, "-")
			If Ubound(fArr) = 1 Then
				Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fArr(0)))
					If Not(rsFun.BOF And rsFun.EOF) Then
						stTime = Trim(rsFun("StartTime"))
					End If
				Set rsFun = Nothing
				Set rsFun = Conn.Execute("Select Top 1 * From HR_Period Where Campus='" & fCampus & "' And Period=" & HR_Clng(fArr(1)))
					If Not(rsFun.BOF And rsFun.EOF) Then
						enTime = Trim(rsFun("EndTime"))
					End If
				Set rsFun = Nothing
				strFun = Trim(stTime) & " - " & Trim(enTime)
			End If
		End If
	End If
	GetPeriodTime = strFun
End Function
%>