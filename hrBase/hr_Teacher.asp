<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incCNtoPY.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./incKPI.asp"-->
<!--#include file="./hr_TeacherInc.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim scriptCtrl, strParm, arrParm
strParm = Trim(Request("Parm")) : arrParm = Split(strParm, "/")
Dim arrField : arrField = Split("YGDM,YGNM,YGXM,YGXB,RYRQ,YGZT,PDZC,PDRQ,PRZC,PRRQ,YGXW,YGXL,YGXZ,BYXX,BYZY,RXRQ,BYRQ,KSDM,KSMC,CSRQ,JG,ZJH,GZRQ,XMJP,ZZMM,SJHM,DH,HLHSKSSJ,HL,PYJG,XZZW,RMRQ,RZJSRQ", ",")
Dim arrFieldName : arrFieldName = Split("员工代码,员工内码,员工姓名,性别,入院原日期,员工状态,评定职称,评定日期,聘任职称,聘任日期,学位,学历,学制,毕业学校,毕业专业,入学日期,毕业日期,科室代码,科室名称,出生日期,籍贯,证件号,工作日期,姓名简拼,政治面貌,手机号码,短号,HLHSKSSJ,HL,聘用机构,职务,任职时间,免职时间", ",")
Dim arrSex : arrSex = Split(XmlText("Config", "Sex", ""), "|")
Dim arrDegrees : arrDegrees = Split(XmlText("Common", "Degrees", ""), "|")		'学位


'在调用ConvertCnToPy()前请预加载拼音数据, 数组变量名为：arrPinYin
Dim arrPinYin, strPinyin : strPinyin = GetHttpPage(apiHost & "/Static/js/PinyinData.js", 1)
If HR_IsNull(strPinyin) = False Then strPinyin = Replace(strPinyin, "var pinyin=""","") : strPinyin = Replace(strPinyin, """;","") : arrPinYin = Split(strPinyin, ",")

Dim SubButTxt : SiteTitle = "教师管理"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index", "List" Call MainBody()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveForm" Call SaveForm()
	Case "AllData" Call getList()
	Case "Preview" Call Preview()
	Case "Delete" Call Delete()
	Case "ResetPass" Call ResetPass()	'重置密码

	Case "Import" Call ImportData()
	Case "ImportView" Call ImportView()
	Case "ImportPost" Call ImportPost()
	Case "ImportAll" Call ImportAll()
	Case "PostImportAll" Call PostImportAll()

	Case "SearchData" Call SearchData()
	Case "ViewAll" Call ViewAll()
	Case "UpdateKPI" Call UpdateKPI()
	Case "UpdatePY" Call UpdatePinyin()
	
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tKSDM : tKSDM = HR_Clng(Request("ks"))

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-width_100 {width:120px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.tit {color:#f30;padding:20px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.list-row {display:flex;align-items:center;flex-wrap:wrap;padding:20px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.list-row li {width:33.33%;padding-right:15px;box-sizing:border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "		.list-row li.load {text-align:center;font-size:18px; width:100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.list-row li.load i {font-size:22px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	tmpHtml = "<a href=""" & ParmPath & "Teacher/List.html"">" & SiteTitle & "</a><a><cite>所有教师</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form soBox""><div class=""layui-inline"">搜索员工：</div><div class=""layui-inline hr-width_100""><input class=""layui-input"" name=""SearchWord"" id=""SearchWord"" placeholder=""搜索姓名/工号"" autocomplete=""off"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><select name=""KSMC"" id=""KSMC"" lay-verify=""required"" lay-search=""""><option value="""">选择/搜索科室</option>" & GetDeptOption(0, tKSDM, False) & "</select></div>" & vbCrlf
	Response.Write "		<div class=""layui-inline""><select name=""DateForm"" id=""DateForm""><option value="""">选择数据来源</option><option value=""3"">医院HIS</option><option value=""2"">全院员工</option><option value=""0"">本系统</option></select></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn""><button class=""layui-btn layui-btn-normal"" data-type=""reload"" id=""SearchBtn""><i class=""layui-icon"">&#xe615;</i>搜索</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-normal"" data-type=""delete"" id=""BatchDel"" title=""批量删除""><i class=""layui-icon"">&#xe640;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_peru"" data-type=""refresh"" id=""refresh"" title=""刷新""><i class=""hr-icon"">&#xebbb;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-normal"" data-type=""addNew"" id=""addNew"" title=""新增员工""><i class=""hr-icon"">&#xe7fe;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-normal"" data-type=""import"" id=""import"" title=""员工导入""><i class=""hr-icon"">&#xe7f0;</i>导入</button>" & vbCrlf		'8月更新接口
	'Response.Write "			<button class=""layui-btn hr-btn_peru"" data-type=""import2"" id=""import2"" title=""导入第二次详细资料""><i class=""hr-icon"">&#xec51;</i></button>" & vbCrlf
	'Response.Write "			<button class=""layui-btn layui-bg-cyan"" data-type=""importAll"" id=""importAll"" title=""导入所有员工""><i class=""hr-icon"">&#xeae4;</i></button>" & vbCrlf
	'Response.Write "			<button class=""layui-btn layui-btn-normal"" data-type=""viewTxt"" id=""viewTxt"" title=""测试""><i class=""hr-icon"">&#xf2bc;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_oran"" data-type=""updatePY"" id=""updatePY"" title=""拼音""><i class=""hr-icon"">&#xf2bc;</i>拼音</button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "Teacher/AllData.html?ks=" & tKSDM & "',text:{none:'暂未找到您需要的教师数据！'},height:'full-130',page:true,limit:20,limits:[10,15,20,30,50,100],id:'TableList'}"" lay-filter=""TableList"">"
	Response.Write "		<thead><tr>" & vbCrlf
	Response.Write "			<th lay-data=""{fixed:'left',type:'checkbox',width:40}""></th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'TeacherID',unresize:true, align:'center',width:70,sort: true}"">排序</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'YGDM',unresize:true, width:100}"">员工代码</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'YGXM',width:120}"">员工姓名</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'KSMC',minWidth:120}"">科室</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'YGZT',align:'center',width:90}"">员工状态</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'PRZC',width:100}"">职称</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'XZZW',width:100}"">职务</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'XMJP',align:'center',width:70}"">拼音</th>" & vbCrlf
	Response.Write "			<th lay-data=""{fixed:'right',align:'center',unresize:true,width:255, toolbar: '#barTable'}"">操作</th>" & vbCrlf
	Response.Write "		</tr></thead>" & vbCrlf
	Response.Write "	</table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm"" lay-event=""detail"" title=""查看详情""><i class=""hr-icon"">&#xf35f;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-warm"" lay-event=""view"" title=""查看业绩""><i class=""hr-icon"">&#xea82;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-danger"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-bg-cyan"" lay-event=""reset"" title=""密码重置""><i class=""hr-icon"">&#xec7f;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""laydate"", ""layedit"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table, layedit = layui.layedit, laydate = layui.laydate;" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form;" & vbCrlf
	
	strHtml = strHtml & "		$(""#SearchBtn"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var key1 = $(this).val(), key3 = $(""#KSMC"").val(), key4 = $(""#DateForm"").val();" & vbCrlf
	strHtml = strHtml & "			table.reload(""TableList"", {" & vbCrlf
	strHtml = strHtml & "				url:""" & ParmPath & "Teacher/AllData.html"",where: {SearchWord:key1, ks:key3, DateForm:key4}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			return false;" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#SearchWord"").bind(""input propertychange"",function(){" & vbCrlf
	strHtml = strHtml & "			var key1 = $(""#SearchWord"").val(), key3 = $(""#KSMC"").val(), key4 = $(""#DateForm"").val();" & vbCrlf
	strHtml = strHtml & "			table.reload(""TableList"", {" & vbCrlf
	strHtml = strHtml & "				url:""" & ParmPath & "Teacher/AllData.html"",where: {SearchWord:key1, ks:key3, DateForm:key4}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#refresh"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			table.reload(""TableList"", {" & vbCrlf
	strHtml = strHtml & "				url:""" & ParmPath & "Teacher/AllData.html"",where: {SearchWord:"""", ks:""""}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#updatePY"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:2,id:""UpdateWin"",content:""" & ParmPath & "Teacher/UpdatePY.html"",title:[""更新教师姓名拼音"", ""font-size:16""], area:[""680px"", ""460px""],maxmin:true});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		$(""#import"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:2, content:""" & ParmPath & "ImportTmp/Teacher.html"",id:""ImportBox"",title:[""导入员工数据"", ""font-size:16""], area:[""680px"", ""360px""],moveOut:true,maxmin:true});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		$(""#importAll"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:2,content:""" & ParmPath & "Teacher/ImportAll.html"",id:""ImportBox"",title:[""导入员工数据【含退休等全体员工】"", ""font-size:16""], area:[""680px"", ""360px""],maxmin:true});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		$(""#BatchDel"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var checkStatus = table.checkStatus(""TableList"");" & vbCrlf
	strHtml = strHtml & "			if(checkStatus.data.length==0){layer.tips(""请选择您要删除的员工！"",""#BatchDel"",{tips: [3, ""#F60""]});return false;}" & vbCrlf
	strHtml = strHtml & "			layer.confirm(""确认要删除选中的“"" + checkStatus.data.length + ""”名员工？<br />删除后无法恢复。"",{icon: 3, title:""重要提示""},function(index){" & vbCrlf
	strHtml = strHtml & "				var arrID = """";" & vbCrlf
	strHtml = strHtml & "				for(var i=0;i<checkStatus.data.length;i++){" & vbCrlf
	strHtml = strHtml & "					if(i > 0){arrID = arrID + "",""}" & vbCrlf
	strHtml = strHtml & "					arrID = arrID + checkStatus.data[i].TeacherID;" & vbCrlf
	strHtml = strHtml & "				}" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "Teacher/Delete.html?ID="" + arrID, function(strForm){" & vbCrlf
	strHtml = strHtml & "					layer.msg(strForm.reMessge,function(){table.reload(""TableList"");});" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "				return false;" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		$(""#addNew"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:1,id:""popBody"",title:[""添加员工"",""font-size:16""],area:[""720px"", ""90%""],maxmin:true },function(index){ });var loadTips = layer.load(1);" & vbCrlf
	strHtml = strHtml & "			$.get(""" & ParmPath & "Teacher/AddNew.html"",{Modify:false,ID:0}, function(strForm){" & vbCrlf
	strHtml = strHtml & "				$(""#popBody"").html(strForm);" & vbCrlf
	strHtml = strHtml & "				$("".layDate"").each(function(){ laydate.render({elem:this}); }); form.render();" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "			layer.close(loadTips);" & vbCrlf
	strHtml = strHtml & "			form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "				$.post(""" & ParmPath & "Teacher/SaveForm.html?ID="", PostData.field, function(result){" & vbCrlf
	strHtml = strHtml & "					var reData = eval(""("" + result + "")""), icon=2;" & vbCrlf
	strHtml = strHtml & "					if(reData.Return){ icon=1; }" & vbCrlf
	strHtml = strHtml & "					layer.alert(reData.reMessge, {icon:icon},function(layero, index){" & vbCrlf
	strHtml = strHtml & "						if(reData.Return){layer.closeAll();window.location.reload();}" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "				});return false;" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		$(""#viewTxt"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var soYGDM = $(""#SearchWord"").val(), soKSDM = $(""#KSMC"").val();" & vbCrlf
	strHtml = strHtml & "			layer.open({type:2,id:""kpiWin"",content:""" & ParmPath & "Teacher/UpdateKPI.html?limit=" & HR_Clng(Request("limit")) & "&word="" + soYGDM + ""&ksdm="" + soKSDM, title:[""更新业绩报表"",""font-size:16""],area:[""760px"", ""460px""],maxmin:true });" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	strHtml = strHtml & "			var data = obj.data;" & vbCrlf
	strHtml = strHtml & "			if(obj.event === ""detail""){" & vbCrlf
	strHtml = strHtml & "				$.get(""" & ParmPath & "Teacher/Preview.html?ID="" + data.TeacherID, function(strForm){" & vbCrlf
	strHtml = strHtml & "					layer.open({type:1,content:strForm,title:[""查看员工信息"",""font-size:16""],area:[""700px"", ""80%""],maxmin:true});" & vbCrlf
	'strHtml = strHtml & "					form.render();" & vbCrlf
	'strHtml = strHtml & "					$("".layui-layer-content"").niceScroll();" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""view""){" & vbCrlf
	strHtml = strHtml & "				window.open(""" & ParmPath & "Tab/ExportExcel.html?teacher="" + data.YGDM);" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""edit""){" & vbCrlf
	strHtml = strHtml & "				var loadTips = layer.load(1);" & vbCrlf
	strHtml = strHtml & "				layer.open({type:1, id:""popBody"", title:[""修改员工信息"",""font-size:16""], area:[""720px"", ""90%""], maxmin:true },function(index){ });" & vbCrlf
	strHtml = strHtml & "				$.get(""" & ParmPath & "Teacher/Edit.html"", {ID:data.TeacherID,ParentID:data.KSDM,Eve:obj.event}, function(strForm){" & vbCrlf
	strHtml = strHtml & "					$(""#popBody"").html(strForm); form.render();" & vbCrlf
	strHtml = strHtml & "					lay("".layDate"").each(function(){;" & vbCrlf
	strHtml = strHtml & "						laydate.render({ elem: this });" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "				$(""#popBody"").niceScroll();" & vbCrlf
	strHtml = strHtml & "				layer.close(loadTips);" & vbCrlf

	strHtml = strHtml & "				form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	strHtml = strHtml & "					$.post(""" & ParmPath & "Teacher/SaveForm.html"",PostData.field, function(result){" & vbCrlf
	strHtml = strHtml & "						var reData = eval(""("" + result + "")"");" & vbCrlf
	strHtml = strHtml & "						if(reData.Return){" & vbCrlf
	strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:1},function(layero, index){layer.closeAll();table.reload(""TableList"");});" & vbCrlf
	strHtml = strHtml & "						}else{" & vbCrlf
	strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:2});" & vbCrlf
	strHtml = strHtml & "						}" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "					return false;" & vbCrlf
	strHtml = strHtml & "					" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""reset""){" & vbCrlf		'重置密码
	strHtml = strHtml & "				layer.confirm(""您确认要重置该员工的密码吗？<br />重置后请用新密码登陆！"", {icon:3,title: ""重置密码提醒""}, function(index){" & vbCrlf
	strHtml = strHtml & "					$.getJSON(""" & ParmPath & "Teacher/ResetPass.html?ID="" + data.TeacherID, function(reData){" & vbCrlf
	strHtml = strHtml & "						layer.alert(reData.reMessge, {icon:1, title: ""系统提示""},function(layero, index){layer.closeAll(); });" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""del""){" & vbCrlf
	strHtml = strHtml & "				layer.confirm('真的删除选中的员工吗？<br />相关的数据将同步删除而且无法恢复！', {icon:3,title: ""删除提醒""}, function(index){" & vbCrlf
	strHtml = strHtml & "					$.getJSON(""" & ParmPath & "Teacher/Delete.html?ID="" + data.TeacherID, function(reData){" & vbCrlf
	strHtml = strHtml & "						if(reData.Return){;" & vbCrlf
	strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:1,title: ""系统提示""},function(layero, index){layer.closeAll();table.reload(""TableList"");});" & vbCrlf
	strHtml = strHtml & "						}else{" & vbCrlf
	strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:2,title: ""系统提示""});" & vbCrlf
	strHtml = strHtml & "						}" & vbCrlf
	strHtml = strHtml & "					});" & vbCrlf
	strHtml = strHtml & "					layer.close(index);" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		form.on(""select"", function(data){" & vbCrlf
	strHtml = strHtml & "			var el = data.elem.title;" & vbCrlf
	strHtml = strHtml & "			$("".txt1"").each(function(){" & vbCrlf
	strHtml = strHtml & "				if($(this).attr(""name"")==el){$(this).val(data.value)};" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		form.verify({" & vbCrlf
	strHtml = strHtml & "			pass: [/^[\S]{6,12}$/,'密码必须6到12位，且不能出现空格']" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	'strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strHtml = Replace(getPageFoot(1), "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHtml)

End Sub

Sub getList()

	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tPage : tPage = HR_Clng(Trim(Request("page")))
	If tPage = 0 Then tPage = 1
	Dim tLimit : tLimit = HR_Clng(Trim(Request("limit")))
	Dim vCount, vMSG, tmpJson, tmpData, rsGet, sqlGet, tmpField, tDataForm, tFirstWord
	Dim soYGDM : soYGDM = Trim(ReplaceBadChar(Request("SearchWord")))
	Dim soKSDM : soKSDM = HR_Clng(ReplaceBadChar(Request("ks")))
	Dim soForm : soForm = Trim(ReplaceBadChar(Request("DateForm")))
	
	Dim sqlCount : sqlCount = "Select Count(TeacherID) From HR_Teacher Where TeacherID>0"

	sqlGet = "Select Top " & tLimit & " TeacherID,ApiType,YGDM,YGXM,KSMC,ZZMM,YGZT,PRZC,XZZW,XMJP From HR_Teacher Where"
	If tPage > 1 Then
		sqlGet = sqlGet & " TeacherID NOT IN(Select Top " & (tPage-1) * tLimit & " TeacherID From HR_Teacher)"
	Else
		sqlGet = sqlGet & " TeacherID>0"
	End If
	tFirstWord = UCase(Left(soYGDM, 1))
	If HR_IsNull(soYGDM) = False Then
		If Asc(tFirstWord) > 64 And Asc(tFirstWord) < 91 Then
			sqlGet = sqlGet & " And XMJP like('" & soYGDM & "%')"
			sqlCount = sqlCount & " And XMJP like('" & soYGDM & "%')"
		ElseIf HR_Clng(soYGDM) > 0 Then
			sqlGet = sqlGet & " And YGDM like('" & soYGDM & "%')"
			sqlCount = sqlCount & " And YGDM like('" & soYGDM & "%')"
		Else
			sqlGet = sqlGet & " And YGXM like '%" & soYGDM & "%'"
			sqlCount = sqlCount & " And YGXM like '%" & soYGDM & "%'"
		End If
	End If
	If soKSDM > 0 Then
		sqlGet = sqlGet & " And KSDM='" & soKSDM & "'"
		sqlCount = sqlCount & " And KSDM='" & soKSDM & "'"
	End If
	If soForm <> "" Then
		sqlGet = sqlGet & " And ApiType=" & HR_Clng(soForm)
		sqlCount = sqlCount & " And ApiType=" & HR_Clng(soForm)
	End If
	sqlGet = sqlGet & " Order By YGDM ASC"

	Set rsGet = Server.CreateObject("ADODB.RecordSet")
		rsGet.Open sqlGet, Conn, 1, 1
		If Not(rsGet.BOF And rsGet.EOF) Then
			i = 0
			Do While Not rsGet.EOF
				tDataForm = "本系统"
				If HR_Clng(rsGet("ApiType")) = 1 Then tDataForm = "医院HIS"
				If i > 0 Then tmpData = tmpData & ","
				tmpData = tmpData & "{""TeacherID"":" & rsGet("TeacherID") & ",""DataForm"":""" & tDataForm & """"
				tmpData = tmpData & ",""YGDM"":""" & rsGet("YGDM") & """,""YGXM"":""" & ReplaceAPIStr(rsGet("YGXM")) & """,""KSMC"":""" & ReplaceAPIStr(rsGet("KSMC")) & """,""ZZMM"":""" & ReplaceAPIStr(rsGet("ZZMM")) & """"
				tmpData = tmpData & ",""YGZT"":""" & rsGet("YGZT") & """,""PRZC"":""" & rsGet("PRZC") & """,""XZZW"":""" & rsGet("XZZW") & """,""XMJP"":""" & Trim(rsGet("XMJP")) & """}"
				rsGet.MoveNext
				i = i + 1
			Loop
		End If
		'tmpField = ",""Fields"":["
		'For m = 1 To rsGet.Fields.Count - 1
		'	If m > 1 Then tmpField = tmpField & ","
		'	tmpField = tmpField & "{""FieldID"":""" & m & """,""FieldName"":""" & rsGet.Fields(m).Name & """}"
		'Next
		'tmpField = tmpField & "]"
	Set rsGet = Nothing
	Set rsGet = Conn.Execute(sqlCount)
		vCount = HR_Clng(rsGet(0))
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""教师信息"",""count"":" & vCount & ",""limit"":" & tLimit & ",""page"":" & tPage & ",""timer"":""" & Timer - BeginTime & """,""data"":[" & tmpData
	tmpJson = tmpJson & "]}"
	Response.Write tmpJson
End Sub
Sub SearchData()
	Dim rsGet, tmpJson, tmpData, vCount, soWord : soWord = Trim(ReplaceBadChar(Request("soWord")))
	If soWord <> "" Then
		Set rsGet = Server.CreateObject("ADODB.RecordSet")
			rsGet.Open("Select Top 50 a.*,b.KSMC From HR_Teacher a Inner Join HR_Department b On a.KSDM=b.KSDM Where YGXM like '%" & soWord & "%'"), Conn, 1, 1
			vCount = rsGet.Recordcount
			If Not(rsGet.BOF And rsGet.EOF) Then
				Do While Not rsGet.EOF
					If i > 0 Then tmpData = tmpData & ","
					tmpData = tmpData & "{""KSMC"":""" & rsGet("KSMC") & """"
					For m = 1 To rsGet.Fields.Count - 2
						tmpData = tmpData & ",""" & rsGet.Fields(m).Name & """:""" & Trim(rsGet.Fields(m).Value) & """"
					Next
					tmpData = tmpData & "}"
					rsGet.MoveNext
					i = i + 1
				Loop
			End If
		Set rsGet = Nothing
	End If
	tmpJson = "{""code"":0,""msg"":""暂无数据"",""count"":" & vCount & ",""data"":[" & tmpData & "]}"
	Response.Write tmpJson
End Sub

Sub Preview()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim rsShow, strChk, iTeacher, tDepartment, viewYGDM, tUpKPI
	Set rsShow = Conn.Execute("Select * From HR_Teacher Where TeacherID=" & tmpID )
		If rsShow.BOF And rsShow.EOF Then
			strHtml = "<div class=""layui-row"" style=""text-align:center;padding:20px 0"">"
			strHtml = strHtml & "<a class=""layui-btn layui-btn-lg layui-btn-danger""><i class=""layui-icon"">&#xe69c;</i> 您要查看的员工信息【ID：" & tmpID & "】不存在！</a></div>"
			Response.Write strHtml
			Exit Sub
		Else
			tDepartment = GetTypeName("HR_Department", "KSMC", "KSDM", rsShow("KSDM"))
			viewYGDM = HR_Clng(rsShow("YGDM"))
			strHtml = "<fieldset class=""layui-elem-field layui-field-title""><legend>员工 " & rsShow("YGXM") & " 预览</legend>"
			strHtml = strHtml & "<div class=""layui-form layer-hr-box""><table class=""layui-table"" lay-skin=""line"">"
			strHtml = strHtml & "<colgroup><col width=""120""><col><col width=""120""><col></colgroup>"
			strHtml = strHtml & "<tbody>"

			strHtml = strHtml & "<tr><td style=""text-align:right;"">员工代码：</td><td>" & rsShow("YGDM") & "</td>"
			strHtml = strHtml & "<td style=""text-align:right;"">姓　名：</td><td>" & Trim(rsShow("YGXM")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">排序序号：</td><td>" & Trim(rsShow("TeacherID")) & "</td>"
			strHtml = strHtml & "<td style=""text-align:right;"">拼音代码：</td><td>" & Trim(rsShow("XMJP")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">政治面貌：</td><td>" & Trim(rsShow("ZZMM")) & "</td><td style=""text-align:right;"">性　别：</td><td>" & Trim(rsShow("YGXB")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">员工状态：</td><td>" & rsShow("YGZT") & "</td>"
			strHtml = strHtml & "<td style=""text-align:right;"">科　室：</td><td>" & Trim(tDepartment) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">职　称：</td><td>" & Trim(rsShow("PRZC")) & "</td>"
			strHtml = strHtml & "<td style=""text-align:right;"">职　务：</td><td>" & Trim(rsShow("XZZW")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">学　位：</td><td>" & Trim(rsShow("YGXW")) & "</td><td style=""text-align:right;"">学　历：</td><td>" & Trim(rsShow("YGXL")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">毕业学校：</td><td colspan=""3"">" & Trim(rsShow("BYXX")) & "</td></tr>"
			strHtml = strHtml & "</tbody>"
			strHtml = strHtml & "</table></div>"  & vbCrlf
			strHtml = strHtml & "</fieldset>" & vbCrlf
			tUpKPI = UpdateTeacherTotalKPI(viewYGDM)	'更新员工总计数据
			Response.Write tUpKPI
		End If
	Set rsShow = Nothing
	Response.Write strHtml
End Sub

Sub EditBody()

	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim isModify : isModify = HR_CBool(Request("Modify"))
	Dim tmpHtml, tPass, arr1, arr2, str2
	Dim tChecked
	tPass = "12345678" : SubButTxt = "添加"
	sqlTmp = "Select * From HR_Teacher Where TeacherID=" & tmpID

	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		Redim arr1(rsTmp.Fields.count-1)
		Redim arr2(rsTmp.Fields.count-1)
		If rsTmp.BOF And rsTmp.EOF Then
			If tmpID > 0 Then
				ErrMsg = "您要修改的员工【ID：" & tmpID & "】不存在！"
				Response.Write GetErrBody(1) : Exit Sub
			End If
		Else
			SubButTxt = "修改"
			For i = 1 To rsTmp.Fields.count-1
				arr1(i) = rsTmp.Fields(i).Value
				str2 = str2 & i & ":" & rsTmp.Fields(i).Name & "<br>"
			Next
		End If
	Set rsTmp = Nothing

	tmpHtml = "<div class=""layer-hr-box"">" & vbCrlf
	tmpHtml = tmpHtml & "<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">员工代码：</label>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""YGDM"" value=""" & arr1(3) & """ placeholder=""员工代码不能为空"" lay-verify=""number"" autocomplete=""off"" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">员工姓名：</label>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""YGXM"" value=""" & arr1(4) & """ placeholder=""员工姓名不能为空"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	
	tmpHtml = tmpHtml & "<div class=""layui-form-item""><label class=""layui-form-label"">登陆密码:</label>" & vbCrlf
	If tmpID > 0 Then
		tmpHtml = tmpHtml & "<div class=""layui-input-inline""><input type=""password"" name=""Pass"" value="""" placeholder=""密码不能修改"" autocomplete=""off"" class=""layui-input"" readonly></div>"
		tmpHtml = tmpHtml & "<div class=""layui-form-mid layui-word-aux"">密码修改请到个人中心自行修改</div>"
	Else
		tmpHtml = tmpHtml & "<div class=""layui-input-inline""><input type=""password"" name=""Pass"" value=""" & tPass & """ placeholder=""员工登陆密码"" lay-verify=""pass"" autocomplete=""off"" class=""layui-input""></div>"
		tmpHtml = tmpHtml & "<div class=""layui-form-mid layui-word-aux"">密码必须8到12位，初始密码为：" & tPass & "</div>"
	End If
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<label class=""layui-form-label"">性　别：</label>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-block"">"
	For i = 1 To Ubound(arrSex)
		tmpHtml = tmpHtml & "<input type=""radio"" name=""YGXB"" lay-skin=""primary"" title=""" & arrSex(i) & """ value=""" & arrSex(i) & """"
		If Trim(arr1(6)) = arrSex(i) Then tmpHtml = tmpHtml & " checked"
		tmpHtml = tmpHtml & ">"
	Next
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<label class=""layui-form-label"">员工状态：</label>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-inline""><input type=""text"" name=""YGZT"" value=""" & arr1(7) & """ class=""layui-input txt1""></div>"
	tmpHtml = tmpHtml & "	<div class=""layui-input-inline""><select name=""SelectYGZT"" title=""YGZT""><option value="""">搜索/选择员工状态</option>" & GetFieldOption("HR_Teacher", "YGZT", arr1(7)) & "</select></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">科　　室:</label>"
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><select name=""KSDM"" lay-verify=""required"" lay-search=""""><option value="""">搜索/选择科室</option>" & GetDepartmentOption(0, arr1(8), False) & "</select></div>"
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">拼音代码:</label>"
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""XMJP"" value=""" & arr1(22) & """ placeholder=""拼音代码"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "	</div>"
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<label class=""layui-form-label"">聘任职称：</label>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-inline""><input type=""text"" name=""PRZC"" value=""" & arr1(12) & """ class=""layui-input txt1""></div>"
	tmpHtml = tmpHtml & "	<div class=""layui-input-inline""><select name=""SelectPRZC"" title=""PRZC"" lay-search=""""><option value="""">搜索/选择聘任职称</option>" & GetFieldOption("HR_Teacher", "PRZC", arr1(12)) & "</select></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<label class=""layui-form-label"">行政职务：</label>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-inline""><input type=""text"" name=""XZZW"" value=""" & arr1(14) & """ class=""layui-input txt1""></div>"
	tmpHtml = tmpHtml & "	<div class=""layui-input-inline""><select name=""SelectXZZW"" title=""XZZW"" lay-search=""""><option value="""">搜索/选择行政职务</option>" & GetFieldOption("HR_Teacher", "XZZW", arr1(14)) & "</select></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item""><label class=""layui-form-label"">个性签名：</label>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-block""><textarea name=""Explain"" id=""Explain"" placeholder=""备注"" class=""layui-textarea"">" & arr1(10) & "</textarea></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	If tmpID > 0 Then tmpHtml = tmpHtml & "<input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""Modify"" value=""True"">"
	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-input-block""><button class=""layui-btn"" lay-submit lay-filter=""SubPost"">" & SubButTxt & "</button><button type=""reset"" class=""layui-btn layui-btn-primary"">重置</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</form>"
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	Response.Write tmpHtml
End Sub
Sub SaveForm()
	Dim tmpJson, tmpID : tmpID = HR_Clng(Request("ID"))

	SubButTxt = "修改"
	If UserRank < 2 Then
		ErrMsg = "{""Return"":false,""Err"":500,""reMessge"":""您没有添加员工的权限"",""ReStr"":[]}"
		Response.Write ErrMsg : Exit Sub
	End If
	Dim tYGDM : tYGDM = HR_Clng(ReplaceBadChar(Request("YGDM")))
	Dim tYGXM : tYGXM = Trim(ReplaceBadChar(Request("YGXM")))
	Dim tKSMC, tKSDM : tKSDM = HR_Clng(ReplaceBadChar(Request("KSDM")))	'科室代码
	Dim isModify : isModify = HR_CBool(Request("Modify"))
	'判断员工代码是否存在

	If HR_IsNull(tYGDM) Or HR_IsNull(tYGDM) Then
		ErrMsg = "{""Return"":false,""Err"":500,""reMessge"":""员工代码/员工姓名 没有填写，请重新输入！"",""ReStr"":[]}"
		Response.Write ErrMsg : Exit Sub
	End If

	If isModify = False Then
		Set rsTmp = Conn.Execute("Select * From HR_Teacher Where YGDM='" & tYGDM & "'")
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				ErrMsg = "{""Return"":false,""Err"":500,""reMessge"":""员工代码 " & tYGDM & " 已存在，不能重复添加！【推荐工号：" & GetNewID("HR_Teacher", "YGDM") & "】"",""ReStr"":[]}"
				Response.Write ErrMsg : Exit Sub
			End If
		Set rsTmp = Nothing
	End If
	If tKSDM > 0 Then tKSMC = GetTypeName("HR_Department", "KSMC", "KSDM", tKSDM)


	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_Teacher Where TeacherID=" & tmpID), Conn, 1, 3
		If rsTmp.BOF And rsTmp.EOF Then
			rsTmp.AddNew
			rsTmp("TeacherID") = GetNewID("HR_Teacher", "TeacherID")
			rsTmp("LoginPass") = "83aa400af464c76d"
			rsTmp("YGDM") = tYGDM
			rsTmp("YGNM") = tYGDM
			rsTmp("ApiType") = 0					'本系统添加，非HIS接口来源
			rsTmp("PXXH") = GetNewID("HR_Teacher", "PXXH")
			rsTmp("ImportTime") = Now()
			SubButTxt = "添加"
		End If
		rsTmp("YGXM") = tYGXM
		rsTmp("KSDM") = tKSDM						'科室代码
		rsTmp("KSMC") = tKSMC								'科室名称
		rsTmp("YGXB") = Trim(Request("YGXB"))				'员工姓别
		rsTmp("YGZT") = Trim(Request("YGZT"))				'员工状态
		rsTmp("XMJP") = Trim(Request("XMJP"))				'姓名简拼
		rsTmp("PRZC") = Trim(Request("PRZC"))				'职称
		rsTmp("XZZW") = Trim(Request("XZZW"))				'行政职务
		rsTmp("Explain") = Trim(Request("Explain"))
		rsTmp("UpdateTime") = Now()
		rsTmp.Update
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""提示：员工" & tYGDM & " 信息" & SubButTxt & "成功！"",""ReStr"":""操作成功！""}"

		Dim tUpKPI : tUpKPI = ChkTeacherKPI(tYGDM)	'添加员工信息至业绩表

		rsTmp.Close
	Set rsTmp = Nothing
	Response.Write tmpJson
End Sub

Sub ImportData()
	Server.ScriptTimeout = 900
	Dim apiDataType : apiDataType = Trim(ReplaceBadChar(Request("apiAction")))

	Dim tmpHtml, xlsUrl, getStr, jsonObj, st1
	getStr = GetHttpPage(apiHost & "/API/API.htm?Type=GetRyxxForJXGL", 1)

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.im-box {min-height:180px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.normal dt {color:#060;} .normal dd h4{color:#060;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = Replace(getPageHead(1), "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)

	tmpHtml = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	tmpHtml = tmpHtml & "	<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	tmpHtml = tmpHtml & "		<legend>远程接口数据【员工】</legend>" & vbCrlf
	'tmpHtml = tmpHtml & "		<div class=""hr-shrink-x10 hr-align_c""><button class=""layui-btn layui-btn-sm"" type=""button"" name=""ImportPost"" id=""ImportPost"">导入</button><button type=""button"" class=""layui-btn layui-btn-sm"" name=""PrevBtn"" id=""PrevBtn"">查看</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""im-box"" id=""ImportData"">"
	If Instr(getStr, "reMessge") > 0 Then
		Set jsonObj = parseJSON(getStr)
			tmpHtml = tmpHtml & "<dl class=""hr-tips_dl normal""><dt><i class=""hr-icon"">&#xef8a;</i></dt><dd><h4>" & jsonObj.reMessge & "</h4><p>" & jsonObj.ReStr & "</p></dd></dl>"
		Set jsonObj = Nothing
	Else
		tmpHtml = tmpHtml & "<dl class=""hr-tips_dl""><dt><i class=""hr-icon"">&#xef61;</i></dt><dd><h4>导入中止！</h4><p>与远程数据通讯时发生错误</p></dd></dl>"
		tmpHtml = tmpHtml & "<div class=""err-msg"">" & getStr & "</div>" & vbCrlf
	End If

	tmpHtml = tmpHtml & "	</div>" & vbCrlf

	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""hr-pop-fix"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><button type=""button"" class=""layui-btn layui-btn-sm hr-btn_darkgreen"" id=""import""><i class=""hr-icon hr-icon-top"">&#xef12;</i>导入</button>"
	tmpHtml = tmpHtml & "<button type=""button"" class=""layui-btn layui-btn-sm hr-btn_peru"" id=""prevdata""><i class=""hr-icon hr-icon-top"">&#xef17;</i>查看</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""hr-shrink-x20""></div>"
	Response.Write tmpHtml

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var  element = layui.element, form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#prevdata"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			parent.layer.open({type:1,content:""<ul class=\""list-row\""><li class=\""load\""><i class=\""layui-icon layui-anim layui-anim-rotate layui-anim-loop\"">&#xe63d;</i>读取中…</li></ul>"",id:""ViewWin"",title:[""导入缓存预览"",""font-size:16""],area:[""90%"", ""80%""],success:function(layero, index){" & vbCrlf
	tmpHtml = tmpHtml & "					$.get(""" & ParmPath & "Teacher/ImportView.html"",{apiAction:""""}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "						parent.$(""#ViewWin"").html(strForm);" & vbCrlf
	tmpHtml = tmpHtml & "					});" & vbCrlf
	tmpHtml = tmpHtml & "				},maxmin:true" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#import"").on(""click"", function(){" & vbCrlf
	tmpHtml = tmpHtml & "			layer.load(); $(""#ImportData"").html(""正在导入数据，请稍候…"");" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "Teacher/ImportPost.html"",{apiAction:""GetRyxxForJXGL""}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "				$(""#ImportData"").html(strForm.reMessge);layer.closeAll(""loading"");" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf


	'strHtml = strHtml & "			$.get(""" & ParmPath & "Teacher/Import.html"", function(strForm){" & vbCrlf
	'strHtml = strHtml & "				$(""#ImportBox"").html(strForm);layer.close(loadTips);" & vbCrlf
	'strHtml = strHtml & "				$(""#PrevBtn"").on(""click"", function(){" & vbCrlf
	'strHtml = strHtml & "					layer.open({type:1,content:"""",id:""ShowWin"",title:[""导入缓存预览"",""font-size:16""],area:[""90%"", ""80%""],maxmin:true }); var loadTips1 = layer.load(1)" & vbCrlf
	'strHtml = strHtml & "					$.get(""" & ParmPath & "Teacher/ViewAll.html"",{apiAction:""GetAllRyxx""}, function(strForm){" & vbCrlf
	'strHtml = strHtml & "						$(""#ShowWin"").html(strForm);layer.close(loadTips1);" & vbCrlf
	'strHtml = strHtml & "						$(""#ShowWin"").niceScroll();" & vbCrlf
	'strHtml = strHtml & "					});" & vbCrlf

	'strHtml = strHtml & "				});" & vbCrlf
	'strHtml = strHtml & "				$(""#ImportPost"").on(""click"", function(){" & vbCrlf
	'strHtml = strHtml & "					layer.load();" & vbCrlf
	'strHtml = strHtml & "					$.getJSON(""" & ParmPath & "Teacher/ImportPost.html"",{apiAction:""GetRyxxForJXGL""}, function(strForm){" & vbCrlf
	'strHtml = strHtml & "						$(""#ImportData"").html(strForm.reMessge);layer.closeAll(""loading"");" & vbCrlf
	'strHtml = strHtml & "					});" & vbCrlf
	'strHtml = strHtml & "				});" & vbCrlf
	'strHtml = strHtml & "			});" & vbCrlf

	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = Replace(getPageFoot(1), "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub
Sub ImportView()
	If UserRank < 1 Then
		ErrMsg = "您没有 导入的权限 的权限！"
		Response.Write GetErrBody(1) : Exit Sub
	End If
	Server.ScriptTimeout = 500
	Dim tmpHtml, xlsUrl, getStr, jsonObj, st1, tYGDM, iAdd
	Dim inData : inData = "[Out]"
	Dim dataFile : dataFile = "Teacher08.txt"
	Dim apiDataType : apiDataType = Trim(ReplaceBadChar(Request("apiAction")))
	If apiDataType = "GetAllRyxx" Then dataFile = "AllTeacher.txt"
	getStr = GetHttpPage(apiHost & "/Upload/" & dataFile, 1)
	If Instr(getStr, "reData") = 0 Then			'没有数据
		Response.Write "<div class=""tit"">没有员工数据</div>" : Exit Sub
	End If
	Set jsonObj = parseJSON(getStr)
		tmpHtml = tmpHtml & "<h3 class=""tit"">共有数据：" & jsonObj.reData.length & "条</h3>" & vbCrlf
		tmpHtml = tmpHtml & "<ul class=""list-row"">" & vbCrlf
		For m = 0 To jsonObj.reData.length - 1
			tYGDM = jsonObj.reData.get(m).YGDM
			Set rs = Conn.Execute("Select Count(0) From HR_Teacher Where YGDM='" & tYGDM & "'")
				If Not(rs.BOF And rs.EOF) Then
					inData = ""
				Else
					iAdd = iAdd + 1
				End If
			Set rs = Nothing
			tmpHtml = tmpHtml & "<li>工号：" & tYGDM & "　姓名：" & jsonObj.reData.get(m).YGXM & "</li>" & vbCrlf
		Next
		tmpHtml = tmpHtml & "</ul>" & vbCrlf
		If HR_CLng(iAdd) > 0 Then tmpHtml = tmpHtml & "<h4>共有" & HR_CLng(iAdd) & "位员工需要导入！</h4>" & vbCrlf
	Set jsonObj = Nothing
	Response.Write tmpHtml
End Sub
Sub ImportPost()
	Server.ScriptTimeout = 1200			'超时20分钟
	Dim tmpJson, rsSave, getStr, jsonObj, tmpLog, j, k
	Dim tUpKPI, tYGDM, tYGXM, tYGXB
	Dim dataFile : dataFile = "Teacher08.txt"
	Dim apiDataType : apiDataType = Trim(ReplaceBadChar(Request("apiAction")))
	If apiDataType = "GetAllRyxx" Then dataFile = "AllTeacher.txt"
	If apiDataType = "GetRyxxForJXGL" Then dataFile = "Teacher08.txt"
	getStr = GetHttpPage(apiHost & "/Upload/" & dataFile, 1)
	If getStr = "" Or isNull(getStr) Then
		Response.Write "{""Return"":false,""Err"":800,""reMessge"":""<p class=\""resultTxt\"">获取员工缓存数据失败！</p>"",""ReStr"":""操作失败！""}"
		Exit Sub
	End If
	Set jsonObj = parseJSON(getStr)
	j = 0 : k = 0
	Dim tmpNow : tmpNow = Now()
	For m=0 To jsonObj.reData.length-1
		tYGDM = Trim(jsonObj.reData.get(m).YGDM)
		tYGXM = Trim(jsonObj.reData.get(m).YGXM)
		If HR_IsNull(tYGDM) = False And HR_IsNull(tYGXM) = False Then
		Set rsSave = Server.CreateObject("ADODB.RecordSet")
			rsSave.Open("Select Top 1 * From HR_Teacher Where YGXM='" & tYGXM & "' And YGDM='" & tYGDM & "'"), Conn, 1, 3
			If rsSave.BOF And rsSave.EOF Then
				rsSave.AddNew
				rsSave("TeacherID") = GetNewID("HR_Teacher", "TeacherID")
				rsSave("YGDM") = tYGDM
				rsSave("YGXM") = tYGXM
				rsSave("PXXH") = GetNewID("HR_Teacher", "PXXH")
				rsSave("LoginPass") = "83aa400af464c76d"
				rsSave("ImportTime") = tmpNow
				j = j + 1
			Else
				k = k + 1
			End If
			rsSave("Explain") = Trim(Request("Explain"))
			rsSave("UpdateTime") = tmpNow
			rsSave("KSDM") = HR_Clng(jsonObj.reData.get(m).KSDM)		'科室代码

			If apiDataType = "GetRyxxForJXGL" Then
				rsSave("ApiType") = 3
				rsSave("KSMC") = Trim(jsonObj.reData.get(m).KSMC)			'科室名称
				rsSave("XMJP") = Trim(jsonObj.reData.get(m).XMJP)			'姓名简拼
				rsSave("YGZT") = Trim(jsonObj.reData.get(m).YGZT)			'员工状态
				rsSave("YGXB") = Trim(jsonObj.reData.get(m).YGXB)			'员工性别
				rsSave("ZCBM") = HR_Clng(jsonObj.reData.get(m).ZCBM)		'职称编码
				rsSave("PRZC") = Trim(jsonObj.reData.get(m).PRZC)
				rsSave("ZWBM") = HR_Clng(jsonObj.reData.get(m).ZWBM)		'职务编码
				rsSave("XZZW") = Trim(jsonObj.reData.get(m).XZZW)			'职务

				rsSave("YGNM") = tYGDM
			ElseIf apiDataType = "GetAllRyxx" Then
				rsSave("ApiType") = 1
				rsSave("YGNM") = Trim(jsonObj.reData.get(m).YGNM)
				tYGXB = HR_Clng(jsonObj.reData.get(m).YGXB)
				rsSave("YGXB") = arrSex(tYGXB)		'员工性别
				rsSave("KSMC") = Trim(jsonObj.reData.get(m).KSMC)			'科室名称
				rsSave("RYRQ") = FormatAPIDate(jsonObj.reData.get(m).RYRQ, 0)		'入院原日期
				rsSave("YGZT") = Trim(jsonObj.reData.get(m).YGZT)
				rsSave("PDZC") = Trim(jsonObj.reData.get(m).PDZC)
				rsSave("PDRQ") = FormatAPIDate(jsonObj.reData.get(m).PDRQ, 0)
				rsSave("PRZC") = Trim(jsonObj.reData.get(m).PRZC)
				rsSave("PRRQ") = FormatAPIDate(jsonObj.reData.get(m).PRRQ, 0)
				rsSave("YGXW") = Trim(jsonObj.reData.get(m).YGXW)
				rsSave("YGXL") = Trim(jsonObj.reData.get(m).YGXL)
				rsSave("YGXZ") = Trim(jsonObj.reData.get(m).YGXZ)
				rsSave("BYXX") = Trim(jsonObj.reData.get(m).BYXX)
				rsSave("BYZY") = Trim(jsonObj.reData.get(m).BYZY)
				rsSave("RXRQ") = FormatAPIDate(jsonObj.reData.get(m).RXRQ, 0)
				rsSave("BYRQ") = FormatAPIDate(jsonObj.reData.get(m).BYRQ, 0)		'毕业日期
				rsSave("CSRQ") = FormatAPIDate(jsonObj.reData.get(m).CSRQ, 0)
				rsSave("JG") = Trim(jsonObj.reData.get(m).JG)
				rsSave("ZJH") = Trim(jsonObj.reData.get(m).ZJH)
				rsSave("GZRQ") = FormatAPIDate(jsonObj.reData.get(m).GZRQ, 0)
				rsSave("XMJP") = Trim(jsonObj.reData.get(m).XMJP)
				rsSave("ZZMM") = Trim(jsonObj.reData.get(m).ZZMM)
				rsSave("SJHM") = Trim(jsonObj.reData.get(m).SJHM)
				rsSave("DH") = Trim(jsonObj.reData.get(m).DH)
				rsSave("HLHSKSSJ") = Trim(jsonObj.reData.get(m).HLHSKSSJ)
				rsSave("HL") = Trim(jsonObj.reData.get(m).HL)
				rsSave("PYJG") = Trim(jsonObj.reData.get(m).PYJG)
				rsSave("XZZW") = Trim(jsonObj.reData.get(m).XZZW)
				rsSave("RMRQ") = FormatAPIDate(jsonObj.reData.get(m).RMRQ, 0)
				rsSave("RZJSRQ") = Trim(jsonObj.reData.get(m).RZJSRQ)
			Else
				rsSave("ApiType") = 2
				rsSave("YGZT") = Trim(jsonObj.reData.get(m).YGLB)
				rsSave("XMJP") = Trim(jsonObj.reData.get(m).PYDM)
				rsSave("ZFPB") = HR_Clng(jsonObj.reData.get(m).ZFPB)
				rsSave("SIGN") = Trim(jsonObj.reData.get(m).SIGN)
				rsSave("PRZC") = Trim(jsonObj.reData.get(m).YGZC)
				rsSave("XZZW") = Trim(jsonObj.reData.get(m).YGZW)
			End If
			rsSave.Update
			tUpKPI = ChkTeacherKPI(tYGDM)	'添加员工信息至业绩表
		Set rsSave = Nothing
		End If
	Next
	Set jsonObj = Nothing
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""<p class=\""resultTxt\"">共有 " & j & " 名新员工导入成功！有 " & k & " 名员工已经存在(资料已更新)！</p>"",""ReStr"":""操作成功！""}"
	Response.Write tmpJson
End Sub

Sub ViewAll()
	Dim tmpHtml
	tmpHtml = "<fieldset class=""layui-elem-field site-demo-button"" style=""margin:10px;"">" & vbCrlf
	tmpHtml = tmpHtml & "	<legend>远程接口数据【员工】</legend>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""ImportData"" id=""ImportData""><p class=""resultTxt""></p>"


	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
	Response.Write tmpHtml
End Sub

Sub UpdateKPI()
	Server.ScriptTimeout=900

	Dim tItemYGDM
	Dim upYGDM : upYGDM = HR_Clng(Request("ygdm"))
	Dim upCount : upCount = HR_Clng(Request("Count"))
	Dim tApiType : tApiType = HR_Clng(Request("ApiType"))
	Dim upKSDM : upKSDM = HR_Clng(Request("ksdm"))
	Dim upSort : upSort = HR_Clng(Request("sort"))

	If upCount = 0 Then upCount = 20

	sqlTmp = "Select Top " & upCount & " * From HR_Teacher Where TeacherID>0"
	If tApiType > 0 Then sqlTmp = sqlTmp & " And ApiType=" & tApiType
	If upYGDM > 0 Then sqlTmp = sqlTmp & " And YGDM='" & upYGDM & "'"
	If upKSDM > 0 Then sqlTmp = sqlTmp & " And KSDM=" & upKSDM

	If upSort = 1 Then
		sqlTmp = sqlTmp & " Order By TeacherID DESC"
	ElseIf upSort = 2 Then
		sqlTmp = sqlTmp & " Order By LoginTime DESC"
	Else
		sqlTmp = sqlTmp & " Order By TeacherID ASC"
	End If

	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			Do While Not rsTmp.EOF
				If i>0 Then tItemYGDM = tItemYGDM & ","
				tItemYGDM = tItemYGDM & rsTmp("YGDM")
				rsTmp.MoveNext
				i = i + 1
			Loop
		End If
	Set rsTmp = Nothing
	tItemYGDM = FilterArrNull(tItemYGDM, ",")

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.ImportTips {text-align:center;line-height:50px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.ImportTips b {color:#f30}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		layui.use([""layer""], function(){ layer.load(1); });" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "<fieldset class=""layui-elem-field site-demo-button"">" & vbCrlf
	Response.Write "	<legend>更新员工业绩报表</legend>" & vbCrlf
	Response.Write "	<div class=""hr-shrink-x10"">" & vbCrlf
	Response.Write "		<div><button class=""layui-btn layui-btn-sm"" id=""refresh"" title=""更新""><i class=""hr-icon"">&#xeeaa;</i></button><button class=""layui-btn layui-btn-sm"" id=""stop"" title=""停止""><i class=""hr-icon"">&#xf28d;</i></button></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-shrink-x10"">" & vbCrlf
	Response.Write "		<div class=""layui-progress layui-progress-big"" lay-showpercent=""true"" lay-filter=""demo"">" & vbCrlf
	Response.Write "			<div class=""layui-progress-bar layui-bg-red"" lay-percent=""0%""></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""ImportTips"" id=""ImportTips"">生成员工业绩报表，可能会持续十至二十分钟　<b>数据准备中，请稍候…</b></div>" & vbCrlf
	Response.Write "</fieldset>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$(document).ready(function(){ });" & vbCrlf
	strHtml = strHtml & "	var arrYgdm = [" & tItemYGDM & "];" & vbCrlf
	strHtml = strHtml & "	layui.use([""form"", ""table"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var form = layui.form, table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element;" & vbCrlf
	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	
	strHtml = strHtml & "		$(""#refresh"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "			updateKPI(0);" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	function updateKPI(iNum){" & vbCrlf			'更新全部员工
	strHtml = strHtml & "		var iBegin=iNum, iEnd = iBegin + 50, iArr=[];console.log(arrYgdm.length);" & vbCrlf
	strHtml = strHtml & "		for(var i=iBegin;i<iEnd;i++){" & vbCrlf
	strHtml = strHtml & "			iArr.push(arrYgdm[i]);" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "		console.log(iArr.join());" & vbCrlf
	strHtml = strHtml & "		$.getJSON(""" & ParmPath & "Tab/Reload.html"",{YGDM:iArr.join()}, function(reData){" & vbCrlf
	strHtml = strHtml & "			if(reData.Return){" & vbCrlf
	strHtml = strHtml & "				$(""#ImportTips b"").html(iBegin + "" / "" + arrYgdm.length);" & vbCrlf
	strHtml = strHtml & "				var p1 = (iBegin/arrYgdm.length)*100;" & vbCrlf
	strHtml = strHtml & "				element.progress('demo', p1.toFixed(2) + '%');" & vbCrlf
	strHtml = strHtml & "				if(iEnd>=arrYgdm.length){" & vbCrlf
	strHtml = strHtml & "					layer.closeAll(""loading""); $(""#ImportTips b"").html(""完成！"");return false;" & vbCrlf
	strHtml = strHtml & "				}else{" & vbCrlf
	strHtml = strHtml & "					updateKPI(iEnd);layer.load(1);" & vbCrlf
	strHtml = strHtml & "				}" & vbCrlf
	strHtml = strHtml & "			}else{" & vbCrlf
	strHtml = strHtml & "				$(""#ImportTips b"").html(reData.reMessge);" & vbCrlf
	strHtml = strHtml & "			}" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	}" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml
End Sub

Sub Delete()
	If UserRank < 2 Then
		ErrMsg = "{""Return"":false,""Err"":400,""reMessge"":""您没有删除员工权限"",""ReStr"":[]}"
		Response.Write ErrMsg : Exit Sub
	End If

	Dim tmpJson, rsDel, sqlDel, strDel, arrDel, iDel, tmpErr, tDelYGDM
	strDel = Trim(ReplaceBadChar(Request("ID")))
	strDel = DelRightComma(strDel)
	arrDel = Split(strDel, ",")
	iDel = 0
	For i = 0 To Ubound(arrDel)
		Set rsDel = Server.CreateObject("ADODB.RecordSet")
			rsDel.Open("Select * From HR_Teacher Where TeacherID=" & HR_Clng(arrDel(i))), Conn, 1, 3
			If Not(rsDel.BOF And rsDel.EOF) Then
				tDelYGDM = HR_Clng(rsDel("YGDM"))
				rsDel.Delete
				iDel = iDel + 1
				rsDel.Close
				Conn.Execute("Delete From HR_KPI_SUM Where YGDM=" & tDelYGDM)		'删除KPI
				Conn.Execute("Delete From HR_KPI Where YGDM=" & tDelYGDM)		'删除KPI
				'删除员工该员工在所有考核项目中的数据
			End If
		Set rsDel = Nothing
	Next
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & iDel & "/" & Ubound(arrDel) + 1 & " 名员工删除成功！" & tmpErr & """,""ReStr"":""操作成功！""}"
	Response.Write tmpJson
End Sub
%>