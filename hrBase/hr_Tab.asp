<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./incKPI.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim SubButTxt : SiteTitle = "员工业绩"

Dim arrStudentType : arrStudentType = Split(XmlText("Common", "StudentType", ""), "|")

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "Reload" Call Reload()
	Case "ExportExcel" Call ExportExcel()
	Case "ReloadTest" Call ReloadTest()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tPage : tPage = HR_Clng(Request("page"))
	If tPage = 0 Then tPage = 1
	Dim tArrSort : tArrSort = Split("学时数正序↑,学时数倒序↓,业绩分正序↑,业绩分倒序↓,科室排序,工号正序↑,工号倒序↓", ",")
	Dim tSort : tSort = HR_Clng(Request("sort"))
	Dim tType, tChecked : tType = HR_Clng(Request("type"))
	If tType = 2 Then tChecked = " checked"

	Dim tLimit : tLimit = HR_Clng(Request("limit"))
	Dim tSword : tSword = Trim(ReplaceBadChar(Request("word")))
	Dim tKSDM : tKSDM = HR_Clng(Request("ksdm"))

	sqlTmp = "Select * From HR_KPI_SUM Where ID>0"
	Dim pageUrl : pageUrl = "Tab.html?type=" & tType
	If tLimit > 0 Then pageUrl = pageUrl & "&limit=" & tLimit
	If HR_IsNull(tSword) = False Then
		pageUrl = pageUrl & "&word=" & tSword
	End If
	If tKSDM > 0 Then
		pageUrl = pageUrl & "&ksdm=" & tKSDM
		sqlTmp = sqlTmp & " And KSDM=" & tKSDM
	End If
	If HR_IsNull(tSort) = False Then pageUrl = pageUrl & "&sort=" & tSort

	Dim iPage, iMaxPage, iTotal : iMaxPage = 20		'取最大页
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		iTotal = HR_Clng(rsTmp.Recordcount)
		If (iTotal Mod iMaxPage) = 0 Then
			iPage = iTotal \ iMaxPage
		Else
			iPage = iTotal \ iMaxPage + 1
		End If
	Set rsTmp = Nothing
	If tPage > iPage Then tPage = iPage

	Dim tSortOption
	For i = 0 To Ubound(tArrSort)
		tSortOption = tSortOption & "<option value=""" & i + 1 & """"
		If tSort = i+1 Then tSortOption = tSortOption & " selected"
		tSortOption = tSortOption & ">" & tArrSort(i) & "</option>"
	Next
%>
<!DOCTYPE html>
<html>
<head>
	<meta charset="utf-8" />
	<title>业绩管理</title>
	<meta name="renderer" content="webkit" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1" />
	<link rel="stylesheet" type="text/css" href="/Static/layui/css/layui.css?v=layui2.2.5" />
	<!--[if lt IE 9]>
		<script src="https://cdn.staticfile.org/html5shiv/r29/html5.min.js"></script>
		<script src="https://cdn.staticfile.org/respond.js/1.4.2/respond.min.js"></script>
	<![endif]-->
	<link rel="stylesheet" type="text/css" href="/Static/Admin/css/hr.lay.css?v=1.0.1" />
	<style type="text/css">
		.box{width:auto;margin:5px auto;}
		.th1, .tr1 td {white-space:nowrap; overflow:hidden; text-overflow:ellipsis;text-align:center;}
		.layui-table td, .layui-table th {padding:3px 8px}

		.soBox {box-sizing:border-box;padding-top:8px;}
		.soBox .searchBtn {vertical-align:top}
		.soBox .layui-inline {margin-bottom:1px;}
		.soBox .layui-form-select dl {top: 31px;}
		.soBox .layui-input {height: 30px;}
		.soBox .layui-btn {height: 30px;line-height: 30px;padding: 0 12px;}
		.soBox .layui-form-select dl dd {padding: 0 5px;line-height: 30px;}
	</style>
	<script type="text/javascript" src="/Static/js/jquery.min.js?v=1.11.2"></script>
	<script type="text/javascript" src="/Static/js/jquery.nicescroll.min.js?v=3.7.6"></script>
	<script type="text/javascript" src="/Static/js/jquery.table2excel.min.js?v=1.1.1"></script>
	<script type="text/javascript" src="/Static/layui/layui.js?v=2.3.0"></script>
	<script type="text/javascript">
		$(function () { });
		layui.use(["layer", "element"], function(){
			var layer = layui.layer, element = layui.element;
			layer.config({skin:"layer-hr"});var loadInit = layer.load();
		});
	</script>

</head>
<body>
<header class="hr-rows iframe-nav">
	<nav class="hr-row navPath"><i class="hr-icon">&#xf101;</i>我的位置：</nav>
	<nav class="hr-row hr-row-fill"><hgroup class="layui-breadcrumb" lay-separator="/"><a href="/Manage/Index/Start.html">开始</a><a href="/Manage/Achieve.html">业绩报表</a><a><cite>查看业绩</cite></a></hgroup></nav>
	<nav class="hr-row navBtn"><a href="javascript:void(0);" class="navLayer"><i class="hr-icon">&#xf141;</i></a></nav>
</header>
<% If UserRank > 0 Then %>
<div class="layui-form soBox"><div class="layui-inline">筛选：</div><div class="layui-inline"><input class="layui-input" name="soTeacher" value="<%=tSword %>" id="soTeacher" placeholder="员工姓名/工号" autocomplete="off" /></div>
	<div class="layui-inline"><select name="KSMC" id="KSMC" lay-search=""><option value="">选择/搜索科室名称</option><%=GetDeptOption(0, tKSDM, 0) %></select></div>
	<div class="layui-inline"><select name="sort" id="sort"><option value="">选择排序方式</option><%=tSortOption %></select></div>
	<div class="layui-inline"><input type="checkbox" name="type" id="sumtype" value="2" title="业绩分"<%=tChecked %>></div>
	<div class="layui-btn-group searchBtn">
		<button class="layui-btn layui-btn-normal" data-type="reload" id="SearchBtn"><i class="hr-icon">&#xea67;</i>搜索</button>
		<button class="layui-btn layui-btn-normal" data-type="export" id="ExportBtn"><i class="hr-icon">&#xf34a;</i></button>
		<button class="layui-btn layui-btn-normal" data-type="prev" id="PrevPage" title="上一页"><i class="hr-icon">&#xf048;</i>上一页</button>
		<button class="layui-btn layui-btn-normal" data-type="next" id="NextPage" title="下一页"><i class="hr-icon">&#xf051;</i>下一页</button>
	</div>
</div>
<% End If %>
<div class='box'>
    <div id="myTable" class="table2excel" style="margin-bottom: 30px">
        <table class="layui-table table table-bordered" id="testTable"></table>
    </div>
</div>
</body>
</html>
<script type="text/javascript" src="/Manage/Ajax/Export.html"></script>
<script type="text/javascript">
	var getUrl;
	getUrl = "/Manage/Ajax/ExportData.html";
	var trs = [];
	start = new Date().getTime();
	foo(columns);

	function pushTrs(arr) {
	  var rank = arr[0].rank;
	  if(trs[rank]){
		$.merge( trs[rank], arr )
	  }else{
		trs[rank]=arr;
	  }
	}

	function render() {
	  var $thead = $('<thead></thead>');
	  var len = trs.length;
	  for (var i = 0; i < trs.length; i++) {
		var $tr = $('<tr></tr>');
		for (var j = 0 ; j < trs[i].length; j++) {
    		var $th = $('<th class="th1" data-field="' + trs[i][j].key + '">' + trs[i][j].title + '</th>');
		  $th.attr('colspan',trs[i][j].colspan);
		  if(trs[i][j].rowspan){
			$th.attr('rowspan',trs[i][j].rowspan);
		  }else{
			$th.attr('rowspan',len-trs[i][j].rank);
		  }
		  $tr.append($th);
		}
		$thead.append($tr);
	  }
	  $('#myTable table').append($thead);
	  $('#myTable table').append("<tbody class='list1'><tr><td colspan='8'>正在构建统计数据(约几分钟)，请稍候…</td></tr></tbody>");
	  end = new Date().getTime();
	  console.log(end - start);
	  PushList();
	}

	function foo(arr, parent) {
	  for (var i = 0; i < arr.length; i++) {
		len = arr[i].children ? arr[i].children.length : 0;
		arr[i].rank = parent ? parent.rank + 1 : 0;
		if (len > 0) {	//children 存在
		  arr[i].rowspan = 1;
		  foo(arr[i].children, arr[i]);
		} else {//children 不存在
		  arr[i].colspan = 1;
		}
		if (parent) {//parent的colspan为children的colspan总和
		  parent.colspan = parent.colspan ? parent.colspan : 0;
		  parent.colspan += arr[i].colspan;
		}
	  }

	  pushTrs(arr);

	  if(arr[0].rank == 0){//最后一次递归结束
		render();
	  }
	}
	function PushList() {
		$.getJSON("/Manage/Ajax/KPIData.html", { word:"<%=tSword %>", limit:<%=tLimit %>,page:<%=tPage %>,type:<%=tType %>,ksdm:"<%=tKSDM %>",sort:<%=tSort %>, sword:$("#soName").val() }, function (tjData) {
			strTR = "";
			for (var j = 0 ; j < tjData.length; j++) {
				strTR += "<tr class='tr1'>";
				//console.log(tjData[j]);
				$.each(tjData[j], function (idx, obj) {
					if (idx !== "id") {
						if (idx == "Items") {
							for (var m = 0; m < obj.length; m++) { strTR += "<td value='" + obj[m].ID + "'>" + obj[m].Score + "</td>" }
						} else {
							strTR += "<td value='" + idx + "'>" + obj + "</td>";
						}
					}
				});
				strTR += "</tr>";
			}
			strTR += "";
			$('.list1').html(strTR);
		});
	}
	layui.use(["table", "form", "element"], function () {
		var table = layui.table; element = layui.element, form = layui.form;
		layer.closeAll("loading");
		var soPage = <%=tPage %>;
		$(".searchBtn button").on("click", function () {
			var btnEvent = $(this).data("type");
			if (btnEvent == "reload") {
				var soWord = $("#soTeacher").val(), soKSMC = $("#KSMC").val(), soSort = $("#sort").val();
				var soType=0; if($("#sumtype").is(":checked")){soType=2;}
				location.href="Tab.html?type=" + soType + "&word=" + soWord + "&sort=" + soSort + "&ksdm=" + soKSMC;
			} else if (btnEvent == "prev") {
				location.href="<%=pageUrl %>&page=" + (soPage-1);
			} else if (btnEvent == "next") {
				location.href="<%=pageUrl %>&page=" + (soPage+1);
			}
		});
	});
	$("#ExportBtn").click(function(){
		$(".table2excel").table2excel({
			// 不被导出的表格行的CSS class类
			exclude: ".noExl",
			// 导出的Excel文档的名称
			name: "Excel Document Name",
			// Excel文件的名称
			filename: "test",
			//文件后缀名
			fileext: ".xls",
			//是否排除导出图片
			exclude_img: false,
			//是否排除导出超链接
			exclude_links: false,
			//是否排除导出输入框中的内容
			exclude_inputs: false
		});
	}); 
</script>
<%
End Sub

Sub ExportExcel()		'导出Excel报表

	Server.ScriptTimeout=1200		'缓存时间20分钟
	Dim sTimer : sTimer = Timer()	'开始时间
	Dim tExcelFileName : tExcelFileName = "PB" & FormatDate(Date(), 2) & ".xls" 
	Dim tLimit : tLimit = HR_Clng(Request("limit"))				'单页数
	If tLimit = 0 Then tLimit = 20								'单页数默认

	Dim TotalPage, TotalRecord, PrevPage, NextPage, tPage : tPage = HR_Clng(Request("page"))					'页码
	Dim tShowTotal : tShowTotal = HR_CBool(Request("total"))		'查看统计
	Dim tChecked, pageUrl
	Dim tSort : tSort = HR_Clng(Request("sort"))					'排序方式
	Dim soTeacher : soTeacher = Trim(ReplaceBadChar(Request("teacher")))	'工号或姓名
	Dim soKSDM : soKSDM = HR_Clng(Request("ksdm"))							'科室代码
	Dim toExcel : toExcel = HR_CBool(Request("excel"))						'输出为Excel
	Dim soYear : soYear = HR_Clng(Request("soyear"))
	If soYear < 2000 Then soYear = DefYear	'如果学年不正确，取系统默认学年

	Dim tSortOption, tArrSort : tArrSort = Split("学时数正序↑,学时数倒序↓,业绩分正序↑,业绩分倒序↓,科室排序,工号正序↑,工号倒序↓", ",")
	For i = 0 To Ubound(tArrSort)	'排序下拉
		tSortOption = tSortOption & "<option value=""" & i + 1 & """"
		If tSort = i+1 Then tSortOption = tSortOption & " selected"
		tSortOption = tSortOption & ">" & tArrSort(i) & "</option>"
	Next
	Dim limitOption, tArrlimit : tArrlimit = Split("20,50,100,200,300", ",")
	For i = 0 To Ubound(tArrlimit)	'排序下拉
		limitOption = limitOption & "<option value=""" & tArrlimit(i) & """"
		If tLimit = HR_Clng(tArrlimit(i)) Then limitOption = limitOption & " selected"
		limitOption = limitOption & ">" & tArrlimit(i) & "条/每页</option>"
	Next

	If Not(toExcel) Then
		tmpHtml = "<style type=""text/css"">" & vbCrlf
		tmpHtml = tmpHtml & "		.layui-table th {text-align:center;}" & vbCrlf
		tmpHtml = tmpHtml & "		.tabth {white-space:nowrap;}" & vbCrlf
		tmpHtml = tmpHtml & "		.limitBox {width:110px;}" & vbCrlf
		tmpHtml = tmpHtml & "		.navBtn a {color:#f30;}" & vbCrlf
		
		tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding:2px 5px;}" & vbCrlf
		tmpHtml = tmpHtml & "		.soBox {box-sizing:border-box;padding-top:8px;}" & vbCrlf
		tmpHtml = tmpHtml & "		.soBox .searchBtn {vertical-align:top}" & vbCrlf
		tmpHtml = tmpHtml & "		.soBox .layui-inline {margin-bottom:1px;}" & vbCrlf
		tmpHtml = tmpHtml & "		.soBox .layui-form-select dl {top: 31px;}" & vbCrlf
		tmpHtml = tmpHtml & "		.soBox .layui-input {height: 30px;}" & vbCrlf
		tmpHtml = tmpHtml & "		.soBox .layui-btn {height: 30px;line-height: 30px;padding: 0 12px;}" & vbCrlf
		tmpHtml = tmpHtml & "		.soBox .layui-form-select dl dd {padding: 0 5px;line-height: 30px;}" & vbCrlf
		tmpHtml = tmpHtml & "	</style>" & vbCrlf
		tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/js/jquery.nicescroll.min.js""></script>" & vbCrlf
		tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
		tmpHtml = tmpHtml & "		layui.use([""layer""], function(){ layer.load(1); });" & vbCrlf
		tmpHtml = tmpHtml & "	</script>" & vbCrlf

		strHtml = getPageHead(1)
		strHtml = Replace(strHtml, "[@HeadStyle]", "")
		strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
		tmpHtml = "<a href=""" & ParmPath & "Achieve/List.html"">" & SiteTitle & "</a><a><cite>查看报表</cite></a>"
		strHtml = strHtml & getFrameNav(1)
		strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
		Call ReplaceCommonLabel(strHtml)
		Response.Write strHtml
		Response.Write "<div class=""hr-workZones hr-shrink-x10"">" & vbCrlf
	Else
		Response.ContentType = "application/ms-download"
		Response.AddHeader "content-disposition", "attachment; filename=" & tExcelFileName & ""
	End If

	Dim tCols, tRows, tabStuTypeNum, iTypeCols, strThead
	strThead = "<table class=""layui-table table-bordered"" border=""1"" id=""ExcelTable""><thead>" & vbCrlf
	strThead = strThead & "<tr>" & vbCrlf
	strThead = strThead & "<th class=""tabth"" rowspan=""4"">序号</th>" & vbCrlf
	strThead = strThead & "<th class=""tabth"" rowspan=""4"">工号</th>" & vbCrlf
	strThead = strThead & "<th class=""tabth"" rowspan=""4"">姓名</th>" & vbCrlf
	strThead = strThead & "<th class=""tabth"" rowspan=""4"">科室</th>" & vbCrlf
	strThead = strThead & "<th class=""tabth"" rowspan=""4"">职称</th>" & vbCrlf
	strThead = strThead & "<th class=""tabth"" rowspan=""4"">学年</th>" & vbCrlf
	strThead = strThead & "<th class=""tabth"" rowspan=""4"">学时数</th>" & vbCrlf
	strThead = strThead & "<th class=""tabth"" rowspan=""4"">业绩分</th>" & vbCrlf
	strThead = strThead & "<th class=""tabth"" rowspan=""4"">等级</th>" & vbCrlf

	'表头输入A、B两大类
	strThead = strThead & "<th class=""tabth"""
	Set rsTmp = Conn.Execute("Select ClassID,StudentType From HR_Class Where ClassType=1 And ParentID=0")		'取A类子栏目数，用于合并列
		tCols = 0 : iTypeCols = 0
		If Not(rsTmp.BOF And rsTMp.EOF) Then
			Do While Not rsTmp.EOF
				tabStuTypeNum = GetTabColsStuType(rsTmp("StudentType"))		'取学生类别数
				iTypeCols = GetTabColsChild(rsTmp(0))	'子类数
				If iTypeCols > 0 Then
					tCols = tCols + iTypeCols
				Else
					If tabStuTypeNum > 0 Then			'有学生类别时合并单元格数
						tCols = tCols + tabStuTypeNum
					Else
						tCols = tCols + 1
					End If
				End If
				rsTmp.MoveNext
			Loop
		End If
		If tCols > 0 Then strThead = strThead & " colspan=""" & tCols & """"	'合并单元格
	Set rsTmp = Nothing
	strThead = strThead & ">基础性教学业绩</th>" & vbCrlf
	strThead = strThead & "<th class=""tabth"""
	Set rsTmp = Conn.Execute("Select ClassID,StudentType From HR_Class Where ClassType=2 And ParentID=0")		'取B类子栏目数，用于合并列
		tCols = 0 : iTypeCols = 0
		If Not(rsTmp.BOF And rsTMp.EOF) Then
			Do While Not rsTmp.EOF
				tabStuTypeNum = GetTabColsStuType(rsTmp("StudentType"))		'取学生类别数
				iTypeCols = GetTabColsChild(rsTmp(0))	'子类数
				If iTypeCols > 0 Then
					tCols = tCols + iTypeCols
				Else
					If tabStuTypeNum > 0 Then			'有学生类别时
						tCols = tCols + tabStuTypeNum
					Else
						tCols = tCols + 1
					End If
				End If
				rsTmp.MoveNext
			Loop
		End If
		If tCols > 0 Then strThead = strThead & " colspan=""" & tCols & """"
	Set rsTmp = Nothing
	strThead = strThead & ">激励性教学业绩</th>" & vbCrlf
	strThead = strThead & "</tr>" & vbCrlf

	'输出第一行【一级项目】
	strThead = strThead & "<tr>"
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassType in(1,2) And ParentID=0 Order By RootID,OrderID")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				strThead = strThead & "<th class=""tabth"""
				tabStuTypeNum = GetTabColsStuType(rsTmp("StudentType"))		'取学生类别数
				If rsTmp("Child") = 0 Then
					If tabStuTypeNum > 0 Then
						strThead = strThead & " rowspan=""2"""		'如果无二级分类，合并行
						strThead = strThead & " colspan=""" & tabStuTypeNum & """"
					Else
						strThead = strThead & " rowspan=""3"""		'如果无二级分类，合并行
					End If
				Else
					strThead = strThead & " colspan=""" & GetTabColsChild(rsTmp("ClassID")) & """"		'如果有二级分类，合并列（含学生类别）
				End If
				strThead = strThead & ">"
				strThead = strThead & rsTmp("ClassName") & "_" & rsTmp("ClassID")
				strThead = strThead & "</th>" & vbCrlf
				rsTmp.MoveNext
			Loop
		End If
	Set rsTmp = Nothing
	strThead = strThead & "</tr>" & vbCrlf

	'输出第二行
	strThead = strThead & "<tr>"
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassType in(1,2) And Depth=1 Order By RootID,OrderID")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				strThead = strThead & "<th class=""tabth"""
				tabStuTypeNum = GetTabColsStuType(rsTmp("StudentType"))		'取学生类别数
				If tabStuTypeNum > 0 Then
					strThead = strThead & " colspan=""" & tabStuTypeNum & """"
				Else
					strThead = strThead & " rowspan=""2"""
				End If
				strThead = strThead & ">"
				strThead = strThead & rsTmp("ClassName") & "_" & rsTmp("ClassID")
				'strThead = strThead & " cols:[" & tCols & "]"
				'strThead = strThead & " rows:[" & tRows & "]"
				strThead = strThead & "</th>" & vbCrlf
				rsTmp.MoveNext
			Loop
			
		End If
	Set rsTmp = Nothing

	strThead = strThead & "</tr>" & vbCrlf

	strThead = strThead & "<tr>"		'学生类别第三行
	Dim tabStuType, tabArrStu, iStu
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassType=1 And StudentType<>'' Order By RootID,OrderID")	'A类
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				tabStuType = rsTmp("StudentType")
				tabStuType = FilterArrNull(tabStuType, ",")
				tabArrStu = Split(tabStuType, ",")
				For iStu=0 To Ubound(tabArrStu)
					strThead = strThead & "<th class=""tabth"">" & tabArrStu(iStu) & "_" & GetStudentType(tabArrStu(iStu))
					strThead = strThead & "</th>" & vbCrlf
				Next
				rsTmp.MoveNext
			Loop
			
		End If
	Set rsTmp = Nothing
	strThead = strThead & "</tr>" & vbCrlf
	strThead = strThead & "</thead>" & vbCrlf

	'取数据行
	strThead = strThead & "<tbody>" & vbCrlf
	Dim sqlOrder, sqlWhere
	Select Case tSort
		Case 1 sqlOrder = " Order By SumScore ASC"
		Case 2 sqlOrder = " Order By SumScore DESC"
		Case 3 sqlOrder = " Order By TotalScore ASC"
		Case 4 sqlOrder = " Order By TotalScore DESC"
		Case 5 sqlOrder = " Order By KSDM DESC"
		Case 6 sqlOrder = " Order By YGDM ASC"
		Case Else sqlOrder = " Order By YGDM DESC"
	End Select

	If soKSDM > 0 Then sqlWhere = " And KSDM=" & soKSDM
	If soYear > 2000 Then sqlWhere = " And scYear=" & soYear
	If HR_IsNull(soTeacher) = False Then
		soTeacher = FilterArrNull(soTeacher, ",")
		If Instr(soTeacher, ",") > 0 Then
			sqlWhere = sqlWhere & " And YGDM in (" & soTeacher & ")"
		ElseIf HR_Clng(soTeacher) > 0 Then
			sqlWhere = sqlWhere & " And YGDM=" & soTeacher
		Else
			sqlWhere = sqlWhere & " And YGXM like '%" & soTeacher & "%'"
		End If
	End If

	sqlTmp = "Select top " & tLimit & " * From HR_KPI_SUM Where YGDM>0"
	If tShowTotal Then sqlTmp = "Select top " & tLimit & " * From HR_KPI Where YGDM>0"
	sqlTmp = sqlTmp & sqlWhere

	Dim rsPage
	Set rsPage = Server.CreateObject("ADODB.RecordSet")
		rsPage.Open(Replace(sqlTmp, "top " & tLimit, "")), Conn, 1, 1
		TotalRecord = rsPage.Recordcount		'总记录
	Set rsPage = Nothing

	If (TotalRecord Mod tLimit) = 0 Then		'计算总页数
		TotalPage = TotalRecord \ tLimit
	Else
		TotalPage = TotalRecord \ tLimit + 1
	End If
	If tPage = 0 Then tPage = 1
	PrevPage = tPage - 1				'取上一页
	If PrevPage = 0 Then PrevPage = 1
	NextPage = tPage + 1				'取下一页
	If NextPage >= TotalPage Then NextPage = TotalPage
	If tPage >= TotalPage Then tPage = TotalPage

	If tPage > 1 Then
		If tShowTotal Then
			sqlTmp = sqlTmp & " And ID NOT IN(Select Top " & (tPage-1) * tLimit & " ID From HR_KPI Where YGDM>0 " & sqlWhere & sqlOrder & ")"
		Else
			sqlTmp = sqlTmp & " And ID NOT IN(Select Top " & (tPage-1) * tLimit & " ID From HR_KPI_SUM Where YGDM>0 " & sqlWhere & sqlOrder & ")"
		End If
	End If
	sqlTmp = sqlTmp & sqlOrder
	If Not(toExcel) Then
		Response.Write "<div class=""layui-form soBox""><div class=""layui-inline"">筛选：</div>" & vbCrlf
		Response.Write "	<div class=""layui-inline"" style=""width:100px""><select name=""SchoolYear"" id=""SchoolYear""><option value="""">选择学年</option>" & GetYearOption(0, soYear) & "</select></div>" & vbCrlf
		Response.Write "	<div class=""layui-inline""><input class=""layui-input"" name=""teacher"" value=""" & soTeacher & """ id=""teacher"" placeholder=""员工姓名/工号"" autocomplete=""off"" /></div>" & vbCrlf
		Response.Write "	<div class=""layui-inline""><select name=""ksdm"" id=""ksdm"" lay-search=""""><option value="""">选择/搜索科室名称</option>" & GetDeptOption(0, soKSDM, 0) & "</select></div>" & vbCrlf
		Response.Write "	<div class=""layui-inline""><select name=""sort"" id=""sort""><option value="""">选择排序方式</option>" & tSortOption & "</select></div>" & vbCrlf
		Response.Write "	<div class=""layui-inline limitBox""><select name=""limit"" id=""limit"">" & limitOption & "</select></div>" & vbCrlf

		If tShowTotal Then tChecked = " checked"
		Response.Write "	<div class=""layui-inline""><input type=""checkbox"" name=""total"" id=""total"" value=""true"" title=""业绩分""" & tChecked & "></div>" & vbCrlf
		Response.Write "	<div class=""layui-btn-group searchBtn"">" & vbCrlf
		Response.Write "		<button class=""layui-btn layui-btn-normal"" data-type=""reload"" id=""SearchBtn""><i class=""hr-icon"">&#xea67;</i>搜索</button>" & vbCrlf
		Response.Write "		<button class=""layui-btn layui-btn-normal"" data-type=""export"" id=""ExportBtn"" title=""导出所有员工业绩报表""><i class=""hr-icon"">&#xf34a;</i>导出Excel</button>" & vbCrlf
		Response.Write "		<button class=""layui-btn layui-btn-normal"" data-type=""prev"" id=""PrevPage"" title=""上一页""><i class=""hr-icon"">&#xf048;</i>上一页</button>" & vbCrlf
		Response.Write "		<button class=""layui-btn layui-btn-normal"" data-type=""next"" id=""NextPage"" title=""下一页""><i class=""hr-icon"">&#xf051;</i>下一页</button>" & vbCrlf
		Response.Write "	</div>" & vbCrlf
		Response.Write "</div>" & vbCrlf
	End If
	Response.Write strThead
	Dim SumNum,TotalNum,ValueNum
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			m = 0
			Do While Not rsTmp.EOF
				SumNum = HR_CDbl(rsTmp("SumScore")) : If SumNum > 0 Then SumNum = FormatNumber(SumNum, 2, -1)			'格式化学时数
				TotalNum = HR_CDbl(rsTmp("TotalScore")) : If TotalNum > 0 Then TotalNum = FormatNumber(TotalNum, 2, -1)	'格式化总分

				Response.Write "<tr>" & vbCrlf
				Response.Write "<td>" & m+1 & "</td>"
				Response.Write "<td>" & rsTmp("YGDM") & "</td>"
				Response.Write "<td class=""tabth"">" & rsTmp("YGXM") & "</td>"
				Response.Write "<td class=""tabth"">" & rsTmp("KSMC") & "</td>"
				Response.Write "<td class=""tabth"">" & rsTmp("PRZC") & "</td>"
				Response.Write "<td>" & rsTmp("scYear") & "</td>"
				Response.Write "<td>" & SumNum & "</td>"
				Response.Write "<td>" & TotalNum & "</td>"
				Response.Write "<td>" & rsTmp.Fields(10).value & "</td>"
				For i = 12 To rsTmp.Fields.Count - 1
					ValueNum = rsTmp.Fields(i).value
					If HR_CDbl(ValueNum) > 0 Then ValueNum = FormatNumber(ValueNum, 2, -1) Else ValueNum=""
					Response.Write "<td>" & ValueNum & "</td>"
				Next
				Response.Write "</tr>" & vbCrlf
				rsTmp.MoveNext
				m = m + 1
			Loop
		Else
			Response.Write "<tr><td colspan=""" & rsTmp.Fields.Count - 1 & """><h3 class=""hr-color-false hr-shrink-x10"">暂无业绩汇总！</h3></td></tr>" & vbCrlf
		End If
	Set rsTmp = Nothing
	Response.Write "</tbody></table>" & vbCrlf
	Response.Flush
	If Not(toExcel) Then
		pageUrl = "ExportExcel.html?total=" & tShowTotal & "&teacher=" & soTeacher & "&sort=2&ksdm=" & soKSDM & "&soyear=" & soYear & "&limit=" & tLimit
		Response.Write "<br>共" & TotalRecord & "条记录　页数：" & tPage & "/" & TotalPage & "页" & vbCrlf
		Response.Write "</div>" & vbCrlf

		tmpHtml = "<script type=""text/javascript"">" & vbCrlf
		tmpHtml = tmpHtml & "	$("".navBtn a"").html(""<i class='hr-icon hr-icon-top'>&#xf351;</i>报表帮助"");" & vbCrlf
		tmpHtml = tmpHtml & "	layui.use([""table"", ""element""], function(){" & vbCrlf
		tmpHtml = tmpHtml & "		var table = layui.table" & vbCrlf
		tmpHtml = tmpHtml & "		element = layui.element;" & vbCrlf
		tmpHtml = tmpHtml & "		layer.closeAll(""loading"");" & vbCrlf

		tmpHtml = tmpHtml & "		$("".searchBtn button"").on(""click"", function () {" & vbCrlf
		tmpHtml = tmpHtml & "			var btnEvent = $(this).data(""type"");" & vbCrlf
		tmpHtml = tmpHtml & "			if (btnEvent == ""reload"") {" & vbCrlf
		tmpHtml = tmpHtml & "				var soteacher = $(""#teacher"").val(), soksdm = $(""#ksdm"").val(), soSort = $(""#sort"").val(), solimit = $(""#limit"").val(), soyear = $(""#SchoolYear"").val();" & vbCrlf
		tmpHtml = tmpHtml & "				var sototal=""false""; if($(""#total"").is("":checked"")){sototal=""true"";}" & vbCrlf
		tmpHtml = tmpHtml & "				location.href=""ExportExcel.html?total=""+ sototal +""&teacher=""+ soteacher +""&sort="" + soSort + ""&ksdm="" + soksdm + ""&limit="" + solimit + ""&soyear="" + soyear ;" & vbCrlf
		tmpHtml = tmpHtml & "			} else if (btnEvent == ""prev"") {" & vbCrlf
		tmpHtml = tmpHtml & "				location.href=""" & pageUrl & "&page=" & PrevPage & """;" & vbCrlf
		tmpHtml = tmpHtml & "			} else if (btnEvent == ""next"") {" & vbCrlf
		tmpHtml = tmpHtml & "				location.href=""" & pageUrl & "&page=" & NextPage & """;" & vbCrlf
		tmpHtml = tmpHtml & "			} else if (btnEvent == ""export"") {" & vbCrlf
		tmpHtml = tmpHtml & "				layer.confirm('本操作将导出所有员工的业绩报表！<br>导出的文件会自动保存到“我的电脑”的“下载”目录中<br>文件名为：" & tExcelFileName & "', {icon:3,title: ""导出报表""}, function(index){" & vbCrlf
		tmpHtml = tmpHtml & "					location.href=""ExportExcel.html?excel=true&total=" & tShowTotal & "&teacher=" & soTeacher & "&sort=" & tSort & "&ksdm=" & soKSDM & "&limit=10000"";" & vbCrlf
		tmpHtml = tmpHtml & "					layer.close(layer.index);" & vbCrlf
		tmpHtml = tmpHtml & "				});" & vbCrlf
		tmpHtml = tmpHtml & "			}" & vbCrlf
		tmpHtml = tmpHtml & "		});" & vbCrlf
		tmpHtml = tmpHtml & "		$("".navBtn"").on(""click"",function(){" & vbCrlf
		tmpHtml = tmpHtml & "			location.href = """ & ParmPath & "Help.html?file=helpAchieve.pdf"";" & vbCrlf
		tmpHtml = tmpHtml & "		});" & vbCrlf
		tmpHtml = tmpHtml & "	});" & vbCrlf
		tmpHtml = tmpHtml & "</script>" & vbCrlf

		strHtml = getPageFoot(1)
		strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
		Response.Write strHtml
	End If
End Sub

Function GetTabColsChild(fClassID)		'返回子类数，包括学生类别
	Dim funArrStu, rsFun, rsFunChild, funCount
	GetTabColsChild = 0 : funCount = 0
	If HR_Clng(fClassID) > 0 Then
		Set rsFun = Conn.Execute("Select ClassID,StudentType From HR_Class Where ParentID=" & fClassID)		'循环子类
			If Not(rsFun.BOF And rsFun.EOF) Then
				Do While Not rsFun.EOF
					If Trim(rsFun(1)) <> "" Then
						funCount = funCount + GetTabColsStuType(Trim(rsFun(1)))	'累计学生类别数
					Else
						funCount = funCount + 1		'子类数
					End If
					rsFun.MoveNext
				Loop
				If funCount > 0 Then GetTabColsChild = funCount
			End If
		Set rsFun = Nothing
	End If
End Function
Function GetTabColsStuType(fStuType)		'返回学生类别数
	Dim funArrStu
	GetTabColsStuType = 0
	fStuType = FilterArrNull(fStuType, ",")
	If HR_IsNull(fStuType) = False Then
		funArrStu = Split(fStuType, ",")
		If Ubound(funArrStu) > 0 Then GetTabColsStuType = Ubound(funArrStu) + 1
	End If
End Function



Sub ReloadTest()
	Server.ScriptTimeout=900
	Dim stime :stime = Timer()
	Dim arrYGDM, tUpKPI, tArrStuType, j, k
	If tYGDM <> "" Then
		tYGDM = FilterArrNull(tYGDM, ",")
		arrYGDM = Split(tYGDM, ",")

		For k = 0 To Ubound(arrYGDM)
			tUpKPI = ChkTeacherKPI(arrYGDM(k))	'添加员工信息至业绩表
			'Response.Write arrYGDM(k) & "<br>"
			Set rsTmp = Conn.Execute("Select * From HR_Class Order By ClassType,RootID,OrderID")
				If Not(rsTmp.BOF And rsTmp.EOF) Then
					Do While Not rsTmp.EOF
						If rsTmp("Child") = 0 Then	'有子类跳过
							tUpKPI = UpdateTeacherKPI(rsTmp("ClassID"), arrYGDM(k), "")
						End If
						rsTmp.MoveNext
					Loop
				End If
			Set rsTmp = Nothing
			tUpKPI = UpdateTeacherTotalKPI(arrYGDM(k))	'更新员工总计数据
		Next
		Response.Write "{""Return"":true,""Err"":0,""reMessge"":""更新完毕" & Ubound(arrYGDM) + 1 & "：" & Timer() - stime & """,""fileUrl"":""""}"
	End If
End Sub

Sub Reload()
	Dim tmpJson, arrYGDM, reStr
	'UpdateKPIField()				'更新字段
	Server.ScriptTimeout=900

	Dim rsData, str1, tArrStuType, tArrFields, strClass, arrClass
	Set rsData = Conn.Execute("Select * From HR_Class Order By ClassType,RootID,OrderID")			'取列
		If Not(rsData.BOF And rsData.EOF) Then
			Do While Not rsData.EOF
				If rsData("Child") = 0 Then	'有子类跳过
					If rsData("StudentType") <> "" Then
						tArrStuType = Split(rsData("StudentType"), ",")
						For i = 0 To Ubound(tArrStuType)
							str1 = str1 & "F" & rsData("ClassID") & "_" & GetStudentType(tArrStuType(i)) & "||"
						Next
					Else
						str1 = str1 & "F" & rsData("ClassID") & "||"
					End If
					strClass = strClass & rsData("ClassID") & "||"
				End If
				rsData.MoveNext
			Loop
		End If
	Set rsData = Nothing
	str1 = FilterArrNull(str1, "||") : strClass = FilterArrNull(strClass, "||")
	tArrFields = Split(str1, "||")
	arrClass = Split(strClass, "||")

	Dim tTableName, tStuType, tTemplate, rs2, sql2, j, k, iSum, iScore, tYGXM
	Dim iFormula, tTotal, tTotalSum
	If tYGDM <> "" Then
		tYGDM = FilterArrNull(tYGDM, ",")
		arrYGDM = Split(tYGDM, ",")
		For i = 0 To Ubound(arrYGDM)
			'检查KPI表中是否有该员工记录，无则添加
			tTotal = 0 : tTotalSum = 0
			tYGXM = strGetTypeName("HR_Teacher", "YGXM", "YGDM", arrYGDM(i))
			Set rs2 = Server.CreateObject("ADODB.RecordSet")
				rs2.Open("Select * From HR_KPI Where YGDM=" & arrYGDM(i)), Conn, 1, 3
				If rs2.BOF And rs2.EOF Then
					rs2.AddNew
					rs2("ID") = GetNewID("HR_KPI", "ID")
					rs2("YGDM") = arrYGDM(i)
				End If
				rs2("YGXM") = tYGXM
				rs2.Update
			Set rs2 = Nothing

			Set rs2 = Server.CreateObject("ADODB.RecordSet")		'更新累计表
				rs2.Open("Select * From HR_KPI_SUM Where YGDM=" & arrYGDM(i)), Conn, 1, 3
				If rs2.BOF And rs2.EOF Then
					rs2.AddNew
					rs2("ID") = GetNewID("HR_KPI_SUM", "ID")
					rs2("YGDM") = arrYGDM(i)
				End If
				rs2("YGXM") = tYGXM
				rs2.Update
			Set rs2 = Nothing

			'Response.Write "<br>" & arrYGDM(i) & "/" & tYGXM & "<br>" & vbCrlf

			'----- 循环所有的课程表
			For j = 0 To Ubound(arrClass)
				tTableName = "HR_Sheet_" & arrClass(j) : iSum = 0 : iScore = 0
				'取学生类别及模板
				tStuType = GetTypeName("HR_Class", "StudentType", "ClassID", arrClass(j))
				tStuType = FilterArrNull(tStuType, ",")
				tTemplate = GetTypeName("HR_Class", "Template", "ClassID", arrClass(j))
				If ChkTable(tTableName) Then			'检查表是否存在
					If HR_IsNull(tStuType) Then			'无学生类别时直接统计
						sql2 = "Select * From " & tTableName & " Where Passed=" & HR_True & " And VA1='" & Trim(arrYGDM(i)) & "'"
						Set rs2 = Server.CreateObject("ADODB.RecordSet")
							rs2.Open sql2, Conn, 1, 1
							iSum = 0 : iScore = 0 : iFormula = 0
							If Not(rs2.BOF And rs2.EOF) Then
								Do While Not rs2.EOF
									iFormula = GetRatioStutype(arrClass(j), "", 0)		'无学生类别系数
									If tTemplate = "TempTableD" Or tTemplate = "TempTableG" Then		'有级别无等级
										If Trim(rs2("VA7")) <> "" Then iFormula = GetRatioLevel(arrClass(j), Trim(rs2("VA7")), "", 0)
									ElseIf tTemplate = "TempTableE" Or tTemplate = "TempTableF" Then		'有级别及等级
										If Trim(rs2("VA7")) <> "" Then iFormula = GetRatioLevel(arrClass(j), Trim(rs2("VA7")), Trim(rs2("VA8")), 0)
									End If

									iSum = iSum + HR_CDbl(rs2("VA3"))
									iScore = iScore + (HR_CDbl(rs2("VA3")) * iFormula)
									rs2.MoveNext
								Loop
							End If
							tTotal = tTotal + iScore
							tTotalSum = tTotalSum + iSum
							'Response.Write arrYGDM(i) & "/" & tTableName & "：" & iFormula & "<br>" & vbCrlf
							Conn.Execute("Update HR_KPI Set F" & arrClass(j) & "=" & iScore & " Where YGDM=" & arrYGDM(i))
							Conn.Execute("Update HR_KPI_SUM Set F" & arrClass(j) & "=" & iSum & " Where YGDM=" & arrYGDM(i))
						Set rs2 = Nothing
					Else
						tArrStuType = Split(tStuType, ",")
						For k = 0 To Ubound(tArrStuType)
							sql2 = "Select * From " & tTableName & " Where Passed=" & HR_True & " And VA1='" & Trim(arrYGDM(i)) & "' And StudentType='" & tArrStuType(k) & "'"
							Set rs2 = Server.CreateObject("ADODB.RecordSet")
								rs2.Open sql2, Conn, 1, 1
								iSum = 0 : iScore = 0 : iFormula = 0
								If Not(rs2.BOF And rs2.EOF) Then
									Do While Not rs2.EOF
										iFormula = GetRatioStutype(arrClass(j), Trim(tArrStuType(k)), 0)		'有学生类别时取系数
										iSum = iSum + HR_CDbl(rs2("VA3"))
										iScore = iScore + (HR_CDbl(rs2("VA3")) * iFormula)
										rs2.MoveNext
									Loop
								End If
								tTotal = tTotal + iScore
								tTotalSum = tTotalSum + iSum

								'Response.Write arrYGDM(i) & "/" & tTableName & "_" & tArrStuType(k) & "：" & iFormula & "<br>" & vbCrlf
								Conn.Execute("Update HR_KPI Set F" & arrClass(j) & "_" & GetStudentType(tArrStuType(k)) & "=" & iScore & " Where YGDM=" & arrYGDM(i))
								Conn.Execute("Update HR_KPI_SUM Set F" & arrClass(j) & "_" & GetStudentType(tArrStuType(k)) & "=" & iSum & " Where YGDM=" & arrYGDM(i))
							Set rs2 = Nothing
						Next
					End If
				Else
					'表不存在，将学生类别赋0值
					If HR_IsNull(tStuType) Then			'无学生类别时直接统计
						'Response.Write arrYGDM(i) & "/" & tTableName & "：" & iSum & "<br>" & vbCrlf
						Conn.Execute("Update HR_KPI Set F" & arrClass(j) & "=" & iScore & " Where YGDM=" & arrYGDM(i))
						Conn.Execute("Update HR_KPI_SUM Set F" & arrClass(j) & "=" & iSum & " Where YGDM=" & arrYGDM(i))
					Else
						tArrStuType = Split(tStuType, ",")
						For k = 0 To Ubound(tArrStuType)
							'Response.Write arrYGDM(i) & "/" & tTableName & "_" & tArrStuType(k) & "：" & iSum & "<br>" & vbCrlf
							Conn.Execute("Update HR_KPI Set F" & arrClass(j) & "_" & GetStudentType(tArrStuType(k)) & "=" & iScore & " Where YGDM=" & arrYGDM(i))
							Conn.Execute("Update HR_KPI_SUM Set F" & arrClass(j) & "_" & GetStudentType(tArrStuType(k)) & "=" & iSum & " Where YGDM=" & arrYGDM(i))
						Next
					End If
				End If

			Next
			'Response.Write "总分：" & tTotal & "<br>" & vbCrlf
			'Response.Write "学时：" & tTotalSum & "<br>" & vbCrlf
			Conn.Execute("Update HR_KPI Set TotalScore=" & tTotal & ",SumScore=" & tTotalSum & ",UpdateTime=(GETDATE()) Where YGDM=" & arrYGDM(i))
			Conn.Execute("Update HR_KPI_SUM Set TotalScore=" & tTotal & ",SumScore=" & tTotalSum & ",UpdateTime=(GETDATE()) Where YGDM=" & arrYGDM(i))
		Next
		tmpJson = "{""Return"":true,""Err"":0,""reMessge"":""共有 " & Ubound(arrYGDM) + 1 & " 位教师业绩更新完毕！"",""fileUrl"":""""}"
	Else
		tmpJson = "{""Return"":false,""Err"":400,""reMessge"":""更新失败，未指定员工！"",""fileUrl"":""""}"
	End If
	Response.Write tmpJson
End Sub

%>