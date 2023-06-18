<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim scriptCtrl
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim SubButTxt : SiteTitle =  "科室管理"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index", "List" Call MainBody()
	Case "AddNew", "Edit" Call EditBody()
	Case "SaveForm" Call SaveForm()
	Case "AllData" Call getList()
	Case "Preview" Call Preview()
	Case "Delete" Call Delete()
	Case "Import" Call ImportData()
	Case "ImportPost" Call ImportPost()
	Case "DeptSort" Call DeptSort()
	Case Else Response.Write GetErrBody(0)
End Select

Sub ListTree()	'树形菜单模式
	Dim rsDept
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	'tmpHtml = tmpHtml & "		.hr-navmenu-main u {display: inline;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin {background-color:#D4E7F0;width:100%;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-item {border-bottom:0px solid #bcd8e6;line-height:35px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-item a {color:#222;padding:0 10px;height:35px;line-height:35px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-item a:hover {background-color:#3EAFE0}" & vbCrlf

	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-child dd:hover {background-color:#B7D5DF}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-child a:hover {color:#f60;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-child dd {border-bottom:1px solid #edf9ff;border-right:1px solid #d4e7f0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-child dd a {padding-left:10px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-itemed>a, .hr-nav-skin .layui-nav-title a, .hr-nav-skin .layui-nav-title a:hover {color: #f00!important;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-itemed .layui-nav-child .layui-nav-itemed dd {border-right:0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-itemed .layui-nav-child .layui-nav-itemed dd a {padding-left:20px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-itemed>.layui-nav-child {background-color: #fff!important;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-more {border-top-color:#79b;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-mored, .hr-nav-skin .layui-nav-itemed > a .layui-nav-more {border-color: transparent transparent #036;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-child dd.layui-this, .hr-nav-skin .layui-nav-child dd.layui-this a, .hr-nav-skin .layui-this, .hr-nav-skin .layui-this>a, .hr-nav-skin .layui-this>a:hover {background-color:transparent;color:#fff;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-nav-skin .layui-nav-child dd.layui-this {border-left:3px solid #f30;background-color:#3EAFE0;background-image:url(/Static/admin/images/sj.png);margin-right:-1px;background-position:right center;background-repeat:no-repeat;}" & vbCrlf

	tmpHtml = tmpHtml & "		.hr-navmenu-main {border-bottom:1px solid #edf9ff;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-navmenu-main i {font-size:15px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-navmenu-main u {display:inline;padding:0 5px;font-style:normal;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	tmpHtml = "<a href=""" & ParmPath & "Branch/List.html"">" & SiteTitle & "</a><a><cite>查看科室</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-side-scroll"">" & vbCrlf
	Response.Write "		<ul class=""layui-nav layui-nav-tree hr-nav-skin"" lay-shrink=""all"" id=""right-menu"" lay-filter=""right-side-menu"">" & vbCrlf
	Set rsDept = Conn.Execute("Select * From HR_Department Where ParentID=0 And RootID>0 Order By RootID ASC, OrderID ASC")
		If Not(rsDept.BOF And rsDept.EOF) Then
			Do While Not rsDept.EOF
				Response.Write "			<li data-name=""home"" class=""layui-nav-item"">" & vbCrlf
				Response.Write "				<a href=""javascript:void(0);"" class=""hr-navmenu-main""><i class=""hr-icon"">&#xe1b2;</i><u>" & rsDept("KSMC") & "</u></a>" & vbCrlf
				Response.Write "				<dl class=""layui-nav-child"">" & ShowNavMenu(rsDept("KSDM")) & "</dl>" & vbCrlf
				Response.Write "			</li>" & vbCrlf
				rsDept.MoveNext
			Loop
		End If
	Set rsDept = Nothing

	Response.Write "		</ul>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table;" & vbCrlf
	strHtml = strHtml & "		element = layui.element;" & vbCrlf

	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml

End Sub

Sub MainBody()
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.layui-elem-field legend {font-size: 16px;}" & vbCrlf
	tmpHtml = tmpHtml & "		#ImportData {padding:0 10px;box-sizing:border-box;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
	tmpHtml = "<a href=""" & ParmPath & "Branch/List.html"">" & SiteTitle & "</a><a><cite>查看科室</cite></a>"
	strHtml = strHtml & getFrameNav(1)
	strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Response.Write "	<div class=""layui-form soBox"">搜索科室：<div class=""layui-inline""><input class=""layui-input"" name=""SearchWord"" id=""SearchWord"" autocomplete=""off"" /></div>" & vbCrlf
	Response.Write "		<div class=""layui-btn-group searchBtn""><button class=""layui-btn layui-btn-normal"" data-type=""reload"" id=""SearchBtn""><i class=""layui-icon"">&#xe615;</i>搜索</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-normal"" data-type=""delete"" id=""BatchDel"" title=""批量删除""><i class=""layui-icon"">&#xe640;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn hr-btn_peru"" data-type=""addNew"" id=""addNew"" title=""新增一级科室""><i class=""layui-icon"">&#xe654;</i></button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-bg-green"" data-type=""import"" id=""import"" title=""科室导入""><i class=""hr-icon"">&#xecb8;</i>导入</button>" & vbCrlf
	Response.Write "			<button class=""layui-btn layui-btn-normal"" data-type=""sort"" id=""sort"" title=""整理排序""><i class=""hr-icon"">&#xf160;</i></button>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<table class=""layui-table"" lay-data=""{url:'" & ParmPath & "Branch/AllData.html',height:'full-130',page:true,limit:20,limits:[10,15,20,30,50,100],id:'TableList'}"" lay-filter=""TableList"">"
	Response.Write "		<thead><tr>" & vbCrlf
	Response.Write "			<th lay-data=""{fixed:'left',type:'checkbox'}""></th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'PXBH',unresize:true, align:'center',width:70,sort: true}"">排序</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'KSDM',unresize:true, width:80}"">科室代码</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'KSMC',width:180}"">科室名称</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'PYDM',unresize:true,width:70}"">拼音代码</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'ParentDept',width:120}"">上级科室</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'CountYG',unresize:true,width:70,sort: true}"">员工数</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'KSDD'}"">科室地点</th>" & vbCrlf
	Response.Write "			<th lay-data=""{field:'KSWM',width:160}"">科室位码</th>" & vbCrlf
	Response.Write "			<th lay-data=""{fixed:'right',align:'center',unresize:true,width:250, toolbar: '#barTable'}"">操作</th>" & vbCrlf
	Response.Write "		</tr></thead>" & vbCrlf
	Response.Write "	</table>" & vbCrlf
	Response.Write "	<script type=""text/html"" id=""barTable"">" & vbCrlf
	Response.Write "		<div class=""layui-btn-group"">" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm"" lay-event=""detail"" title=""查看详情""><i class=""hr-icon"">&#xf35f;</i></a>" & vbCrlf
	Response.Write "			{{#  if(d.Depth<""3""){ }}" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-bg-cyan layui-anim-scale"" lay-event=""add"" title=""添加下级科室""><i class=""hr-icon"">&#xe3ba;</i></a>" & vbCrlf
	Response.Write "			{{#  }else{ }}" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-disabled"" lay-event=""addChil"" title=""不能添加下级科室""><i class=""hr-icon"">&#xe611;</i></a>" & vbCrlf
	Response.Write "			{{#  } }}" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-normal"" lay-event=""edit"" title=""编辑""><i class=""hr-icon"">&#xebf7;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-warm"" lay-event=""list"" title=""查看员工""><i class=""hr-icon"">&#xeeed;</i></a>" & vbCrlf
	Response.Write "			<a class=""layui-btn layui-btn-sm layui-btn-danger"" lay-event=""del"" title=""删除""><i class=""hr-icon"">&#xebcf;</i></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</script>" & vbCrlf
	Response.Write "</div>" & vbCrlf


	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	layui.use([""table"", ""form"", ""layedit"", ""upload"", ""element""], function(){" & vbCrlf
	strHtml = strHtml & "		var table = layui.table, layedit = layui.layedit, upload = layui.upload;" & vbCrlf
	strHtml = strHtml & "		element = layui.element, form = layui.form;" & vbCrlf

	strHtml = strHtml & "		$(""#SearchBtn"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var key1 = $(""#SearchWord"").val();" & vbCrlf
	strHtml = strHtml & "			table.reload(""TableList"", {" & vbCrlf
	strHtml = strHtml & "				url:""" & ParmPath & "Branch/AllData.html"",where: {SearchWord:key1}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		$(""#sort"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "		layer.load();" & vbCrlf
	strHtml = strHtml & "			$.getJSON(""" & ParmPath & "Branch/DeptSort.html"", function(strForm){" & vbCrlf
	strHtml = strHtml & "				layer.msg(strForm.reMessge, {icon:1,time:0,btn:[""确定""]},function(index){" & vbCrlf
	strHtml = strHtml & "					window.location.reload();" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "				layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		$(""#import"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:2, id:""ImportWin"", content:""" & ParmPath & "Branch/Import.html"", title:[""导入科室数据"",""font-size:16""],area:[""80%"", ""85%""],maxmin:true});" & vbCrlf
	'strHtml = strHtml & "			$.get(""" & ParmPath & "Branch/Import.html"", function(strForm){" & vbCrlf
	'strHtml = strHtml & "				layer.open({type:1,content:strForm,title:[""导入科室数据"",""font-size:16""],area:[""760px"", ""90%""],offset:[""70px"",""100px""],maxmin:true});" & vbCrlf
	'strHtml = strHtml & "				form.on(""submit(ImportPost)"", function(PostData){" & vbCrlf
	'strHtml = strHtml & "					$.getJSON(""" & ParmPath & "Branch/ImportPost.html"", function(strForm){" & vbCrlf
	'strHtml = strHtml & "						layer.alert(strForm.reMessge, {icon:1,title: ""系统提示"",area:""500px""},function(layero, index){window.location.reload();layer.closeAll();});" & vbCrlf
	'strHtml = strHtml & "					});" & vbCrlf
	'strHtml = strHtml & "					return false;" & vbCrlf
	'strHtml = strHtml & "				});" & vbCrlf
	'strHtml = strHtml & "				form.render();" & vbCrlf
	'strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		$(""#addNew"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			layer.open({type:2,id:""EditWin"",content:""" & ParmPath & "Branch/AddNew.html?AddNew=True"",title:[""添加一级科室"",""font-size:16""],area:[""780px"", ""460px""],maxmin:true });" & vbCrlf
	'strHtml = strHtml & "			$.get(""" & ParmPath & "Branch/AddNew.html?AddNew=True"", function(strForm){" & vbCrlf
	'strHtml = strHtml & "				form.render();layer.close(loadTips);" & vbCrlf
	'strHtml = strHtml & "				form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	'strHtml = strHtml & "					$.post(""" & ParmPath & "Branch/SaveForm.html"",PostData.field, function(result){" & vbCrlf
	'strHtml = strHtml & "						var reData = eval(""("" + result + "")"");" & vbCrlf
	'strHtml = strHtml & "						if(reData.Return){" & vbCrlf
	'strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:1},function(layero, index){layer.closeAll();window.location.reload();});" & vbCrlf
	'strHtml = strHtml & "						}else{" & vbCrlf
	'strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:2});" & vbCrlf
	'strHtml = strHtml & "						}" & vbCrlf
	'strHtml = strHtml & "					});" & vbCrlf
	'strHtml = strHtml & "					return false;" & vbCrlf
	'strHtml = strHtml & "				});" & vbCrlf
	'strHtml = strHtml & "				$("".layui-layer-content"").niceScroll();" & vbCrlf
	'strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#BatchDel"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var checkStatus = table.checkStatus(""TableList"");" & vbCrlf
	strHtml = strHtml & "			if(checkStatus.data.length==0){layer.tips(""请选择您要删除的科室！"",""#BatchDel"",{tips: [3, ""#F60""]});return false;}" & vbCrlf
	strHtml = strHtml & "			layer.confirm(""确认要删除选中的“"" + checkStatus.data.length + ""”个科室？<br />删除后无法恢复。"",{icon: 3, title:""重要提示""},function(index){" & vbCrlf
	strHtml = strHtml & "				var arrID = """";" & vbCrlf
	strHtml = strHtml & "				for(var i=0;i<checkStatus.data.length;i++){" & vbCrlf
	strHtml = strHtml & "					if(i > 0){arrID = arrID + "",""}" & vbCrlf
	strHtml = strHtml & "					arrID = arrID + checkStatus.data[i].DepartID;" & vbCrlf
	strHtml = strHtml & "				}" & vbCrlf
	strHtml = strHtml & "				$.getJSON(""" & ParmPath & "Branch/Delete.html?ID="" + arrID, function(strForm){" & vbCrlf
	strHtml = strHtml & "					layer.msg(strForm.reMessge,function(){table.reload(""TableList"");});" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "				return false;" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf

	strHtml = strHtml & "		table.on(""tool(TableList)"", function(obj){" & vbCrlf
	strHtml = strHtml & "			var data = obj.data;" & vbCrlf
	strHtml = strHtml & "			if(obj.event === ""prevCarousel""){" & vbCrlf
	strHtml = strHtml & "				location.href=""../Branch/Index/3.html?ItemID="" + data.DepartID;" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""detail""){" & vbCrlf
	strHtml = strHtml & "				$.get(""" & ParmPath & "Branch/Preview.html?ID="" + data.DepartID, function(strForm){" & vbCrlf
	strHtml = strHtml & "					layer.open({type:1,content:strForm,title:[""查看科室信息"",""font-size:16""],area:[""700px"", ""460px""],maxmin:true});" & vbCrlf
	strHtml = strHtml & "					form.render();" & vbCrlf
	strHtml = strHtml & "					$("".layui-layer-content"").niceScroll();" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""add""||obj.event === ""edit""){" & vbCrlf
	strHtml = strHtml & "				layer.open({type:2,id:""EditWin"",content:""" & ParmPath & "Branch/Edit.html?ID="" + data.DepartID + ""&ParentID="" + data.KSDM + ""&Eve="" + obj.event,title:[""编辑科室资料"",""font-size:16""],area:[""780px"", ""460px""],maxmin:true});" & vbCrlf

	'strHtml = strHtml & "				$.get(""" & ParmPath & "Branch/Edit.html"",{ID:data.DepartID,ParentID:data.KSDM,Eve:obj.event}, function(strForm){" & vbCrlf
	'strHtml = strHtml & "					layer.open({type:1,content:strForm,title:[""编辑科室资料"",""font-size:16""],area:[""700px"", ""80%""],maxmin:true});" & vbCrlf
	'strHtml = strHtml & "					form.render();" & vbCrlf
	'strHtml = strHtml & "					form.on(""submit(SubPost)"", function(PostData){" & vbCrlf
	'strHtml = strHtml & "					$.post(""" & ParmPath & "Branch/SaveForm.html"",PostData.field, function(result){" & vbCrlf
	'strHtml = strHtml & "						var reData = eval(""("" + result + "")"");" & vbCrlf
	'strHtml = strHtml & "						if(reData.Return){" & vbCrlf
	'strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:1},function(layero, index){layer.closeAll();window.location.reload();});" & vbCrlf
	'strHtml = strHtml & "						}else{" & vbCrlf
	'strHtml = strHtml & "							layer.alert(reData.reMessge, {icon:2});" & vbCrlf
	'strHtml = strHtml & "						}" & vbCrlf
	'strHtml = strHtml & "					});" & vbCrlf
	'strHtml = strHtml & "					return false;" & vbCrlf
	'strHtml = strHtml & "					});" & vbCrlf
	'strHtml = strHtml & "					$("".layui-layer-content"").niceScroll();" & vbCrlf
	'strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""list""){" & vbCrlf
	strHtml = strHtml & "				window.location.href=""" & ParmPath & "Teacher/List.html?ks="" + data.KSDM;" & vbCrlf
	strHtml = strHtml & "			}else if(obj.event === ""del""){" & vbCrlf
	strHtml = strHtml & "				layer.confirm('真的删除选中的部门吗？<br />相关的数据将同步删除而且无法恢复！', {icon:3,title: ""删除提醒""}, function(index){" & vbCrlf
	strHtml = strHtml & "					$.getJSON(""" & ParmPath & "Branch/Delete.html?ID="" + data.DepartID, function(reData){" & vbCrlf
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

	strHtml = strHtml & "		layer.closeAll(""loading"");" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	'strHtml = strHtml & "	var niceObj = $(""html"").niceScroll();" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1) & strHtml
	strHtml = Replace(strHtml, "[@FootScript]", "")
	Response.Write strHtml

End Sub

Sub getList()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim vCount, vMSG, tmpJson, tmpData, rsGet, sqlGet
	Dim tWord : tWord = Trim(ReplaceBadChar(Request("SearchWord")))
	Dim tParentID, ParentDept, tOrder, tDeptName

	sqlGet = "Select a.*,(Select Count(TeacherID) From HR_Teacher Where ApiType=3 And KSDM = a.KSDM) As CountYG From HR_Department a Where a.DepartmentID>0"
	If tWord <> "" Then sqlGet = sqlGet & " And KSMC like '%" & tWord & "%'"
	sqlGet = sqlGet & " Order By RootID ASC,OrderID ASC"
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
			Dim sp1
			Do While Not rsGet.EOF
				tParentID = HR_Clng(rsGet("ParentID")) : ParentDept = ""
				tDeptName = Trim(rsGet("KSMC"))
				If tParentID > 0 Then
					ParentDept = GetTypeName("HR_Department", "KSMC", "KSDM", tParentID)
					Set rsTmp = Conn.Execute("Select Max(OrderID) From HR_Department Where SJKS=" & tParentID)
						tOrder = HR_Clng(rsTmp(0))
					Set rsTmp = Nothing
					sp1 = ""
					tDeptName = "　" & tDeptName
				End If

				If i > 0 Then tmpData = tmpData & ","
				tmpData = tmpData & "{""DepartID"":" & rsGet("DepartmentID") & ",""KSDM"":""" & HR_Clng(rsGet("KSDM")) & """,""KSMC"":""" & tDeptName & """,""PYDM"":""" & Trim(rsGet("PYDM")) & """,""PXBH"":""" & HR_Clng(rsGet("PXBH")) & """,""KSDD"":""" & Trim(rsGet("KSDD")) & """"
				tmpData = tmpData & ",""SJKS"":""" & HR_Clng(rsGet("SJKS")) & """,""ParentDept"":""" & Trim(ParentDept) & """,""KSWM"":""" & Trim(rsGet("KSWM")) & """,""Depth"":" & HR_Clng(rsGet("Depth")) & ",""CountYG"":" & HR_Clng(rsGet("CountYG")) & "}"
				rsGet.MoveNext
				i = i + 1
				If i >= MaxPerPage Then Exit Do
			Loop
		End If
		vCount = rsGet.Recordcount
	Set rsGet = Nothing
	tmpJson = "{""code"":0,""msg"":""暂无数据"",""count"":" & vCount & ",""data"":[" & tmpData
	tmpJson = tmpJson & "],""limit"":" & MaxPerPage & ",""page"":" & CurrentPage & "}"
	Response.Write tmpJson
End Sub

Sub Preview()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim rsShow, strChk, iTeacher, tParentDept
	Set rsShow = Conn.Execute("Select * From HR_Department Where DepartmentID=" & tmpID )
		If rsShow.BOF And rsShow.EOF Then
			strHtml = "<div class=""layui-row"" style=""text-align:center;padding:20px 0"">"
			strHtml = strHtml & "<a class=""layui-btn layui-btn-lg layui-btn-danger""><i class=""layui-icon"">&#xe69c;</i> 您要查看的科室信息【ID：" & tmpID & "】不存在！</a></div>"
			Response.Write strHtml
			Exit Sub
		Else
			Set rsTmp = Conn.Execute("Select Count(TeacherID) From HR_Teacher Where KSDM=" & rsShow("KSDM"))
				iTeacher = HR_Clng(rsTmp(0))
			Set rsTmp = Nothing
			tParentDept = GetTypeName("HR_Department", "KSMC", "KSDM", rsShow("SJKS"))

			strHtml = "<fieldset class=""layui-elem-field layui-field-title""><legend>" & rsShow("KSMC") & "预览</legend>"
			strHtml = strHtml & "<div class=""layui-form layer-hr-box""><table class=""layui-table"" lay-skin=""line"">"
			strHtml = strHtml & "<colgroup><col width=""120""><col><col width=""120""><col></colgroup>"
			strHtml = strHtml & "<tbody>"

			strHtml = strHtml & "<tr><td style=""text-align:right;"">科室代码：</td><td>" & rsShow("KSDM") & "</td>"
			strHtml = strHtml & "<td style=""text-align:right;"">科室名称：</td><td>" & Trim(rsShow("KSMC")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">拼音代码：</td><td>" & Trim(rsShow("PYDM")) & "</td>"
			strHtml = strHtml & "<td style=""text-align:right;"">排序编号：</td><td>" & Trim(rsShow("PXBH")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">科室地点：</td><td colspan=""3"">" & Trim(rsShow("KSDD")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">上级科室：</td><td>" & Trim(tParentDept) & "</td>"
			strHtml = strHtml & "<td style=""text-align:right;"">科室位码：</td><td>" & Trim(rsShow("KSWM")) & "</td></tr>"
			strHtml = strHtml & "<tr><td style=""text-align:right;"">员工数　：</td><td>" & iTeacher & "</td>"
			strHtml = strHtml & "<td style=""text-align:right;""></td><td></td></tr>"
			strHtml = strHtml & "</tbody>"
			strHtml = strHtml & "</table></div>"  & vbCrlf
			strHtml = strHtml & "</fieldset>" & vbCrlf
		End If
	Set rsShow = Nothing
	Response.Write strHtml
End Sub

Sub EditBody()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tParentID : tParentID = HR_Clng(Request("ParentID"))
	Dim eveAction : eveAction = Trim(ReplaceBadChar(Request("Eve")))
	Dim tKSDM, tKSMC, tPYDM, tPXBH, tKSDD, tSJKS, tKSWM, tDepth
	SubButTxt = "添加"
	Dim opTitle : opTitle = "一级科室"
	If eveAction = "add" Then tmpID = 0 : opTitle = "二级科室"

	sqlTmp = "Select * From HR_Department"
	If tmpID > 0 Then
		sqlTmp = sqlTmp & " Where DepartmentID=" & tmpID : SubButTxt = "修改" : opTitle = "科室"
		Set rsTmp = Conn.Execute(sqlTmp)
			If rsTmp.BOF And rsTmp.EOF Then
				tmpHtml = "<div class=""layui-row"" style=""text-align:center;padding:20px 0"">"
				tmpHtml = tmpHtml & "<a class=""layui-btn layui-btn-lg layui-btn-danger""><i class=""layui-icon"">&#xe69c;</i> 您要修改的科室【ID：" & tmpID & "】不存在！</a></div>"
				Response.Write tmpHtml
				Exit Sub
			Else
				tKSDM = HR_Clng(rsTmp("KSDM"))
				tKSMC = Trim(rsTmp("KSMC"))
				tPYDM = Trim(rsTmp("PYDM"))
				tPXBH = HR_Clng(rsTmp("PXBH"))
				tKSDD = Trim(rsTmp("KSDD"))
				tSJKS = HR_Clng(rsTmp("SJKS"))
				tKSWM = Trim(rsTmp("KSWM"))
				tParentID = HR_Clng(rsTmp("ParentID"))
				tDepth = HR_Clng(rsTmp("Depth"))
			End If
		Set rsTmp = Nothing
	End If

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	tmpHtml = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	tmpHtml = tmpHtml & "	<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	tmpHtml = tmpHtml & "		<legend>" & SubButTxt & opTitle & "</legend>" & vbCrlf
	tmpHtml = tmpHtml & "<div class=""layer-hr-box"">"
	tmpHtml = tmpHtml & "<form class=""layui-form"" id=""EditForm"" name=""EditForm"" action="""">"

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">科室代码：</label>"
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""KSDM"" value=""" & tKSDM & """ placeholder=""必须与接口中的代码一致"" lay-verify=""number"" autocomplete=""off"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">科室名称：</label>"
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""KSMC"" value=""" & tKSMC & """ placeholder=""科室名称不能为空"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "	</div>"
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">拼音代码：</label>"
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""PYDM"" value=""" & tPYDM & """ placeholder=""拼音代码不能为空"" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">排序编号：</label>"
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""PXBH"" value=""" & tPXBH & """ placeholder="""" lay-verify=""required"" autocomplete=""off"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "	</div>"
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item""><label class=""layui-form-label"">科室地点：</label>"
	tmpHtml = tmpHtml & "<div class=""layui-input-block""><input type=""text"" name=""KSDD"" value=""" & tKSDD & """ placeholder="""" autocomplete=""off"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "</div>"

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">上级科室：</label>" & vbCrlf
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline"">" & vbCrlf
	If tDepth < 4 Then
		tmpHtml = tmpHtml & "		<select name=""ParentID"" lay-verify=""required"" lay-search=""""><option value="""">直接选择或搜索选择</option>"
		tmpHtml = tmpHtml & GetDepartmentOption(0, tParentID, False)
		tmpHtml = tmpHtml & "</select>" & vbCrlf
	Else
		tmpHtml = tmpHtml & "		<select name=""ParentID"" disabled=""""><option value="""">一级科室不能选择上级</option></select>" & vbCrlf
	End If
	tmpHtml = tmpHtml & "		</div>" & vbCrlf
	tmpHtml = tmpHtml & "	</div>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-inline""><label class=""layui-form-label"">科室位码：</label>"
	tmpHtml = tmpHtml & "		<div class=""layui-input-inline""><input type=""text"" name=""KSWM"" value=""" & tKSWM & """ placeholder="""" autocomplete=""off"" class=""layui-input""></div>"
	tmpHtml = tmpHtml & "	</div>"
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""layui-form-item"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""layui-form-mid""><em class=""hr-help""><i class=""hr-icon"">&#xecfd;</i>请确认您了解科室管理的操作流程！<br>系统不建议添加或修改科室信息，请从接口中导入</em></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	If tmpID > 0 Then tmpHtml = tmpHtml & "<input type=""hidden"" name=""ID"" value=""" & tmpID & """><input type=""hidden"" name=""Modify"" value=""True"">"
	tmpHtml = tmpHtml & "<div class=""hr-pop-fix"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""formBtn""><button class=""layui-btn layui-btn-sm layui-bg-cyan"" lay-submit lay-filter=""SubPost"">" & SubButTxt & "</button><button type=""reset"" class=""layui-btn layui-btn-sm layui-btn-primary"">重置</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</form>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
	Response.Write tmpHtml

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""table"", ""form"", ""element"", ""layedit""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var table = layui.table, form = layui.form, element = layui.element, layedit = layui.layedit;" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub
Sub SaveForm()
	Dim tmpJson, tKSMC, tRootID, tDepth, tUpKPI
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	tKSMC = Trim(ReplaceBadChar(Request("KSMC")))
	Dim tParentID : tParentID = HR_Clng(Request("ParentID"))
	If tParentID > 0 Then
		tRootID = GetTypeName("HR_Department", "RootID", "ParentID", tParentID)
		tDepth = GetTypeName("HR_Department", "Depth", "ParentID", tParentID)
	Else
		Set rsTmp = Conn.Execute("Select Max(RootID) From HR_Department Where ParentID=" & tParentID)
			tRootID = HR_Clng(rsTmp(0)) + 1
		Set rsTmp = Nothing
		tDepth = 0
	End If

	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_Department Where DepartmentID=" & tmpID), Conn, 1, 3
		If rsTmp.BOF And rsTmp.EOF Then
			rsTmp.AddNew
			rsTmp("DepartmentID") = GetNewID("HR_Department", "DepartmentID")
		End If
		rsTmp("KSDM") = HR_Clng(Request("KSDM"))
		rsTmp("KSMC") = Trim(Request("KSMC"))
		rsTmp("PYDM") = Trim(Request("PYDM"))
		rsTmp("PXBH") = HR_Clng(Request("PXBH"))
		rsTmp("KSDD") = Trim(Request("KSDD"))
		rsTmp("SJKS") = tParentID
		rsTmp("KSWM") = Trim(Request("KSWM"))
		rsTmp("SJKS") = tParentID
		rsTmp("RootID") = HR_Clng(tRootID)
		rsTmp("Depth") = tDepth

		rsTmp.Update
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""提示：科室 " & tKSMC & " 更新成功！"",""ReStr"":""操作成功！""}"
		rsTmp.Close
	Set rsTmp = Nothing

	tUpKPI = UpdateKPIField()		'更新业绩表字段
	Response.Write tmpJson
End Sub

Sub ImportData()
	Dim xlsUrl, getStr, jsonObj, st1
	getStr = GetHttpPage(apiHost & "/API/API.htm?Type=GetKsDict", 1)
	Set jsonObj = parseJSON(getStr)

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", "")
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
	
	tmpHtml = "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	tmpHtml = tmpHtml & "<fieldset class=""layui-elem-field layui-field-title"">" & vbCrlf
	tmpHtml = tmpHtml & "	<legend>远程接口数据【科室】</legend>" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""ImportData"" id=""ImportData"">" & vbCrlf

' 用法：Dim obj, scriptCtrl : Set obj = parseJSON(json)
' 长度：jsonObj.reData.length
' 子节点数组取值：jsonObj.reData(0).KSMC

	tmpHtml = tmpHtml & "	<table class=""layui-table"" lay-skin=""line""><thead><tr><th>科室代码</th><th>科室名称</th><th>拼音代码</th><th>科室地点</th><th>上级科室</th><th>排序编号</th><th>状态</th></tr></thead>" & vbCrlf
	tmpHtml = tmpHtml & "	<tbody>" & vbCrlf
	
	For m=0 To jsonObj.reData.length-1
		Set rsTmp = Conn.Execute("Select Top 1 * From HR_Department Where KSMC='" & Trim(jsonObj.reData.get(m).KSMC) & "' And KSDM=" & jsonObj.reData.get(m).KSDM & "")
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				st1 = "<span style=""color:#f30"">有</span>"
			Else
				st1 = "<span style=""color:#060"">无</span>"
			End If
		Set rsTmp = Nothing
		tmpHtml = tmpHtml & "		<tr>" & vbCrlf
		tmpHtml = tmpHtml & "		<td>" & jsonObj.reData.get(m).KSDM & "</td>" & vbCrlf
		tmpHtml = tmpHtml & "		<td>" & jsonObj.reData.get(m).KSMC & "</td>" & vbCrlf
		tmpHtml = tmpHtml & "		<td>" & jsonObj.reData.get(m).PYDM & "</td>" & vbCrlf
		tmpHtml = tmpHtml & "		<td>" & jsonObj.reData.get(m).KSDD & "</td>" & vbCrlf
		tmpHtml = tmpHtml & "		<td>" & jsonObj.reData.get(m).SJKS & "</td>" & vbCrlf
		tmpHtml = tmpHtml & "		<td>" & jsonObj.reData.get(m).KSWM & "</td>" & vbCrlf
		tmpHtml = tmpHtml & "		<td>" & st1 & "</td>" & vbCrlf
		tmpHtml = tmpHtml & "		</tr>" & vbCrlf
	Next
	Set jsonObj = Nothing
	tmpHtml = tmpHtml & "	</tbody></table>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	tmpHtml = tmpHtml & "</fieldset>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""hr-shrink-x10""></div>" & vbCrlf

	tmpHtml = tmpHtml & "<div class=""hr-pop-fix"" id=""ImportBox"">" & vbCrlf
	tmpHtml = tmpHtml & "	<div class=""formBtn""><button class=""layui-btn layui-btn-sm"" lay-submit lay-filter=""ImportPost"">导入</button><button type=""reset"" class=""layui-btn layui-btn-sm layui-btn-primary"">重置</button></div>" & vbCrlf
	tmpHtml = tmpHtml & "</div>" & vbCrlf
	Response.Write tmpHtml

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	layui.use([""form"", ""element""], function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var form = layui.form;" & vbCrlf
	tmpHtml = tmpHtml & "		element = layui.element;" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub ImportPost()
	Dim tmpJson, rsSave, getStr, jsonObj, tmpLog, j, tUpKPI
	getStr = GetHttpPage(apiHost & "/API/API.htm?Type=GetKsDict", 1)
	If getStr = "" Or isNull(getStr) Then
		Response.Write "{""Return"":false,""Err"":800,""reMessge"":""获取部门数据接口失败！"",""ReStr"":""操作成功！""}"
		Exit Sub
	End If
	Set jsonObj = parseJSON(getStr)
	j = 0
	For m=0 To jsonObj.reData.length-1
		Set rsSave = Server.CreateObject("ADODB.RecordSet")
			rsSave.Open("Select Top 1 * From HR_Department Where KSMC='" & Trim(jsonObj.reData.get(m).KSMC) & "' And KSDM=" & jsonObj.reData.get(m).KSDM & ""), Conn, 1, 3
			If rsSave.BOF And rsSave.EOF Then
				rsSave.AddNew
				rsSave("DepartmentID") = GetNewID("HR_Department", "DepartmentID")
				rsSave("KSDM") = HR_Clng(jsonObj.reData.get(m).KSDM)
				rsSave("KSMC") = Trim(jsonObj.reData.get(m).KSMC)
				rsSave("PYDM") = Trim(jsonObj.reData.get(m).PYDM)
				rsSave("PXBH") = HR_Clng(jsonObj.reData.get(m).PXBH)
				rsSave("KSDD") = Trim(jsonObj.reData.get(m).KSDD)
				rsSave("SJKS") = HR_Clng(jsonObj.reData.get(m).SJKS)
				rsSave("KSWM") = Trim(jsonObj.reData.get(m).KSWM)
				rsSave("RootID") = 0
				rsSave("Depth") = 0
				rsSave("Child") = 0
				rsSave.Update
				j = j + 1
			Else
				tmpLog = tmpLog & "<br>" & Trim(jsonObj.reData.get(m).KSMC) & "【代码：" & HR_Clng(jsonObj.reData.get(m).KSDM) & "】已存在，暂未导入！"
			End If
		Set rsSave = Nothing
	Next
	Set jsonObj = Nothing
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & HR_Clng(j) & "个部门导入成功！" & Trim(tmpLog) & """,""ReStr"":""操作成功！""}"
	Response.Write tmpJson
End Sub

Sub Delete()
	Dim tmpJson, rsDel, sqlDel, strDel, arrDel, iDel, tmpErr
	strDel = Trim(ReplaceBadChar(Request("ID")))
	strDel = DelRightComma(strDel)
	arrDel = Split(strDel, ",")
	iDel = 0
	For i = 0 To Ubound(arrDel)
		Set rsDel = Server.CreateObject("ADODB.RecordSet")
			rsDel.Open("Select * From HR_Department Where DepartmentID=" & HR_Clng(arrDel(i))), Conn, 1, 3
			If Not(rsDel.BOF And rsDel.EOF) Then
				Set rsTmp = Conn.Execute("Select Count(DepartmentID) From HR_Department Where SJKS=" & HR_Clng(rsDel("KSDM")))
					If rsTmp(0) > 0 Then
						tmpErr = tmpErr & "<br>" & rsDel("KSMC") & " 包含有下级科室，请先删除下级科室后再试！"
					Else
						rsDel.Delete
						iDel = iDel + 1
					End If
				Set rsTmp = Nothing
				rsDel.Close
			End If
		Set rsDel = Nothing
	Next
	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & iDel & "/" & Ubound(arrDel) + 1 & " 条记录删除成功！" & tmpErr & """,""ReStr"":""操作成功！""}"
	Response.Write tmpJson
End Sub

Sub DeptSort()
	Dim tmpJson
	Conn.Execute("Update HR_Department Set ParentID=0,Depth=0,OrderID=0,RootID=0,Child=0")
	'一级科室排序
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_Department Where SJKS=KSDM Order By PXBH ASC"), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 1
			Do While Not rsTmp.EOF
				Conn.Execute("Update HR_Department Set ParentID=0,Depth=0,OrderID=0,RootID=" & i & " Where KSDM=" & rsTmp("KSDM"))
				rsTmp.MoveNext
				i = i + 1
			Loop
		End If
	Set rsTmp = Nothing
	'更新二级子类的ParentID及RootID
	Dim tSJKS, tRootID, rsRoot, tDepth, rs2, j
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_Department Where SJKS=KSDM Order By RootID ASC"), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 1
			Do While Not rsTmp.EOF
				tSJKS = rsTmp("KSDM")
				tRootID = rsTmp("RootID")
				'循环二级科室
				Set rs2 = Conn.Execute("Select * From HR_Department Where SJKS<>KSDM And SJKS=" & tSJKS & " Order By PXBH ASC")
					If Not(rs2.BOF And rs2.EOF) Then
						j = 1
						Do While Not rs2.EOF
							Conn.Execute("Update HR_Department Set ParentID=" & tSJKS & ",RootID=" & tRootID & ",Depth=1,OrderID=" & j & " Where KSDM=" & rs2("KSDM"))
							rs2.MoveNext
							j = j + 1
						Loop
						If j > 1 Then Conn.Execute("Update HR_Department Set Child=" & j-1 & " Where KSDM=" & tSJKS)
					End If
					rs2.Close
				Set rs2 = Nothing
				
				rsTmp.MoveNext
			Loop
			i = i + 1
		End If
	Set rsTmp = Nothing
	'更新三级子类
	Dim tOrderID
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_Department Where Depth=1 And ParentID>0 Order By RootID ASC,OrderID ASC"), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 1
			Do While Not rsTmp.EOF
				tSJKS = rsTmp("KSDM")
				tRootID = rsTmp("RootID")
				tOrderID = rsTmp("OrderID")
				'循环三级科室
				Set rs2 = Conn.Execute("Select * From HR_Department Where SJKS=" & tSJKS & " Order By PXBH ASC")
					If Not(rs2.BOF And rs2.EOF) Then
						j = 1
						Do While Not rs2.EOF
							Conn.Execute("Update HR_Department Set OrderID=OrderID+1 Where RootID=" & tRootID & " And OrderID>" & tOrderID)
							Conn.Execute("Update HR_Department Set ParentID=" & tSJKS & ",RootID=" & tRootID & ",Depth=2,OrderID=" & tOrderID + 1 & " Where KSDM=" & rs2("KSDM"))
							rs2.MoveNext
							j = j + 1
						Loop
						If j > 1 Then Conn.Execute("Update HR_Department Set Child=" & j-1 & " Where KSDM=" & tSJKS)
					End If
					rs2.Close
				Set rs2 = Nothing

				rsTmp.MoveNext
			Loop
			i = i + 1
		End If
	Set rsTmp = Nothing

	'更新四级子类
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_Department Where Depth=2 Order By RootID ASC,OrderID ASC"), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 1
			Do While Not rsTmp.EOF
				tSJKS = rsTmp("KSDM")
				tRootID = rsTmp("RootID")
				tOrderID = rsTmp("OrderID")
				'循环四级科室
				Set rs2 = Conn.Execute("Select * From HR_Department Where SJKS=" & tSJKS & " Order By PXBH ASC")
					If Not(rs2.BOF And rs2.EOF) Then
						j = 1
						Do While Not rs2.EOF
							Conn.Execute("Update HR_Department Set OrderID=OrderID+1 Where RootID=" & tRootID & " And OrderID>" & tOrderID)
							Conn.Execute("Update HR_Department Set ParentID=" & tSJKS & ",RootID=" & tRootID & ",Depth=3,OrderID=" & tOrderID+1 & " Where KSDM=" & rs2("KSDM"))
							rs2.MoveNext
							j = j + 1
						Loop
						If j > 1 Then Conn.Execute("Update HR_Department Set Child=" & j-1 & " Where KSDM=" & tSJKS)
					End If
					rs2.Close
				Set rs2 = Nothing
				rsTmp.MoveNext
			Loop
			i = i + 1
		End If
	Set rsTmp = Nothing

	'更新五级子类
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open("Select * From HR_Department Where Depth=3 Order By RootID ASC,OrderID ASC"), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 1
			Do While Not rsTmp.EOF
				tSJKS = rsTmp("KSDM")
				tRootID = rsTmp("RootID")
				tOrderID = rsTmp("OrderID")
				'循环五级科室
				Set rs2 = Conn.Execute("Select * From HR_Department Where SJKS=" & tSJKS & " Order By PXBH ASC")
					If Not(rs2.BOF And rs2.EOF) Then
						j = 1
						Do While Not rs2.EOF
							Conn.Execute("Update HR_Department Set OrderID=OrderID+1 Where RootID=" & tRootID & " And OrderID>" & tOrderID)
							Conn.Execute("Update HR_Department Set ParentID=" & tSJKS & ",RootID=" & tRootID & ",Depth=4,OrderID=" & tOrderID+1 & " Where KSDM=" & rs2("KSDM"))
							rs2.MoveNext
							j = j + 1
						Loop
						If j > 1 Then Conn.Execute("Update HR_Department Set Child=" & j-1 & " Where KSDM=" & tSJKS)
					End If
					rs2.Close
				Set rs2 = Nothing
				rsTmp.MoveNext
			Loop
			i = i + 1
		End If
	Set rsTmp = Nothing

	tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""整理科室序列成功！"",""ReStr"":""操作成功！""}"
	Response.Write tmpJson
End Sub

Function ShowNavMenu(fSJKS)
	Dim strMenu, rsMenu, sqlMenu, fParentID, fItemName, fOrder, fIcon
	fSJKS = HR_Clng(fSJKS)

	sqlMenu = "Select * From HR_Department Where ParentID=" & fSJKS & " And RootID>0 Order By RootID ASC, OrderID ASC"
	Set rsMenu = Conn.Execute(sqlMenu)
		If Not(rsMenu.BOF And rsMenu.EOF) Then
			strMenu = strMenu & vbCrlf
			Do While Not rsMenu.EOF
				fParentID = HR_Clng(rsMenu("ParentID"))
				fItemName = Trim(rsMenu("KSMC"))
				If fParentID > 0 Then
					Set rsTmp = Conn.Execute("Select Max(OrderID) From HR_Department Where ParentID=" & fParentID)
						fOrder = HR_Clng(rsTmp(0))
					Set rsTmp = Nothing
				End If

				fIcon = "<i class=""hr-icon"">&#xf31d;</i>"

				If fParentID > 0 And HR_Clng(rsMenu("OrderID")) <> fOrder Then strMenu = strMenu & "		"
				If fParentID > 0 And HR_Clng(rsMenu("OrderID")) <= fOrder Then fIcon = "<i class=""hr-icon"">&#xf328;</i>"
				If HR_Clng(rsMenu("OrderID")) = fOrder And fParentID > 0 Then strMenu = strMenu & "		"

				strMenu = strMenu & "<dd data-name=""console"">"
				If HR_Clng(rsMenu("Child")) > 0 Then strMenu = strMenu & vbCrlf & "	"

				strMenu = strMenu & "<a "
				If HR_Clng(rsMenu("Child")) > 0 Then strMenu = strMenu & " class=""hr-navmenu-main"""
				strMenu = strMenu & "href=""javascript:void(0);"" target=""rightFrame"">" & fIcon & fItemName & "</a>"

				If HR_Clng(rsMenu("Child")) > 0 Then strMenu = strMenu & vbCrlf & "	<dl class=""layui-nav-child"">" & vbCrlf
				If HR_Clng(rsMenu("Child")) = 0 Then strMenu = strMenu & "</dd>" & vbCrlf
				If HR_Clng(rsMenu("OrderID")) = fOrder And fParentID > 0 Then strMenu = strMenu & "	</dl>" & vbCrlf & "</dd>" & vbCrlf
				rsMenu.MoveNext
			Loop
		End If
	Set rsMenu = Nothing

	ShowNavMenu = strMenu
End Function
%>