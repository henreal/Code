<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="incCommon.asp"-->
<!--#include file="incPurview.asp"-->
<!--#include file="../hrBase/incKPI.asp"-->
<%
SiteTitle = "我的业绩"

Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index","Mine" Call MainBody()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Server.ScriptTimeout=1800		'缓存时间30分钟
	Dim tExcelFileName : tExcelFileName = "PB" & FormatDate(Date(), 2) & ".xls" 
	Dim tLimit : tLimit = HR_Clng(Request("limit"))				'单页数
	If tLimit = 0 Then tLimit = 20								'单页数默认

	Dim TotalPage, TotalRecord, PrevPage, NextPage, tPage : tPage = HR_Clng(Request("page"))					'页码
	Dim tShowTotal : tShowTotal = HR_CBool(Request("total"))		'查看统计
	Dim tChecked, pageUrl, tmpField
	Dim tSort : tSort = HR_Clng(Request("sort"))					'排序方式
	Dim soTeacher : soTeacher = Trim(ReplaceBadChar(Request("teacher")))	'工号或姓名
	Dim soKSDM : soKSDM = HR_Clng(Request("ksdm"))							'科室代码
	Dim soYear : soYear = HR_Clng(Request("soyear"))
	If soYear < 2000 Then soYear = DefYear	'如果学年不正确，取系统默认学年

	Dim toExcel : toExcel = HR_CBool(Request("excel"))						'输出为Excel

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
		tmpHtml = tmpHtml & "		.layui-table td, .layui-table th {padding:2px 5px; white-space:nowrap;word-break:keep-all;}" & vbCrlf
		tmpHtml = tmpHtml & "	</style>" & vbCrlf
		tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
		tmpHtml = tmpHtml & "		layui.use([""layer""], function(){ layer.load(1); });" & vbCrlf
		tmpHtml = tmpHtml & "	</script>" & vbCrlf

		strHtml = getPageHead("Desktop", 1)
		strHtml = Replace(strHtml, "[@HeadStyle]", "")
		strHtml = Replace(strHtml, "[@HeadScript]", tmpHtml)
		If Action="Mine" Then
			tmpHtml = "<a href=""" & ParmPath & "Achieve/Mine.html"">" & SiteTitle & "</a><a><cite>查看报表</cite></a>"
		Else
			tmpHtml = "<a href=""" & ParmPath & "Achieve/Index.html"">业绩报表</a><a><cite>查看</cite></a>"
		End If
		strHtml = strHtml & getFrameNav(1)
		strHtml = Replace(strHtml, "[@Module_Path]", tmpHtml)
		Response.Write ReplaceCommonLabel(strHtml)
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
	If UserRank > 0 Then		'仅管理员才能查看业绩分和等级
		strThead = strThead & "<th class=""tabth"" rowspan=""4"">业绩分</th>" & vbCrlf
		strThead = strThead & "<th class=""tabth"" rowspan=""4"">等级</th>" & vbCrlf
	End If
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
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassType=1 And ParentID=0 Order By RootID,OrderID")
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
				strThead = strThead & rsTmp("ClassName") : tmpField = tmpField & "_" & rsTmp("ClassID")
				strThead = strThead & "</th>" & vbCrlf
				rsTmp.MoveNext
			Loop
		End If
	Set rsTmp = Nothing

	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassType=2 And ParentID=0 Order By RootID,OrderID")		'B类
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
				strThead = strThead & rsTmp("ClassName")
				'strThead = strThead & " cols:[" & tCols & "]"
				'strThead = strThead & " rows:[" & tRows & "]"
				strThead = strThead & "</th>" & vbCrlf
				rsTmp.MoveNext
			Loop
		End If
	Set rsTmp = Nothing
	strThead = strThead & "</tr>" & vbCrlf

	'输出第二行
	strThead = strThead & "<tr>"
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassType=1 And Depth=1 Order By RootID,OrderID")
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
				strThead = strThead & rsTmp("ClassName")
				'strThead = strThead & " cols:[" & tCols & "]"
				'strThead = strThead & " rows:[" & tRows & "]"
				strThead = strThead & "</th>" & vbCrlf
				rsTmp.MoveNext
			Loop
			
		End If
	Set rsTmp = Nothing
	Set rsTmp = Conn.Execute("Select * From HR_Class Where ClassType=2 And Depth=1 Order By RootID,OrderID")	'B类
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
				strThead = strThead & rsTmp("ClassName")
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
					strThead = strThead & "<th class=""tabth"">" & tabArrStu(iStu) & "</th>" & vbCrlf
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
	
	If Action="Mine" Then soTeacher = UserYGDM

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
		Response.Write "	<div class=""layui-inline"" style=""width:150px""><select name=""SchoolYear"" id=""SchoolYear""><option value="""">选择学年</option>" & GetYearOption(0, soYear) & "</select></div>" & vbCrlf
		Response.Write "	<div class=""layui-inline""><input class=""layui-input"" name=""teacher"" value=""" & soTeacher & """ id=""teacher"" placeholder=""员工姓名/工号"" autocomplete=""off"" /></div>" & vbCrlf
		If UserRank>0 And Action="Index" Then
			Response.Write "	<div class=""layui-inline""><select name=""ksdm"" id=""ksdm"" lay-search=""""><option value="""">选择/搜索科室名称</option>" & GetDeptOption(0, soKSDM, 0) & "</select></div>" & vbCrlf
			Response.Write "	<div class=""layui-inline""><select name=""sort"" id=""sort""><option value="""">选择排序方式</option>" & tSortOption & "</select></div>" & vbCrlf
			Response.Write "	<div class=""layui-inline limitBox""><select name=""limit"" id=""limit"">" & limitOption & "</select></div>" & vbCrlf
		End If
		If tShowTotal Then tChecked = " checked"
		If UserRank>0 Then Response.Write "	<div class=""layui-inline""><input type=""checkbox"" name=""total"" id=""total"" value=""true"" title=""业绩分""" & tChecked & "></div>" & vbCrlf
		Response.Write "	<div class=""layui-inline searchBtn""><button type=""button"" class=""layui-btn layui-bg-cyan"" data-type=""reload"" id=""SearchBtn""><i class=""hr-icon"">&#xea67;</i>搜索</button></div>" & vbCrlf
		If UserRank>0 And Action="Index" Then		'此为管理员选项
			Response.Write "	<div class=""layui-btn-group searchBtn"">" & vbCrlf
			Response.Write "		<button type=""button"" class=""layui-btn hr-btn_olive"" data-type=""export"" id=""ExportBtn"" title=""导出所有员工业绩报表""><i class=""hr-icon"">&#xf34a;</i>导出Excel</button>" & vbCrlf
			Response.Write "		<button type=""button"" class=""layui-btn"" data-type=""prev"" id=""PrevPage"" title=""上一页""><i class=""hr-icon"">&#xf048;</i>上一页</button>" & vbCrlf
			Response.Write "		<button type=""button"" class=""layui-btn"" data-type=""next"" id=""NextPage"" title=""下一页""><i class=""hr-icon"">&#xf051;</i>下一页</button>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
		End If
		Response.Write "</div>" & vbCrlf
	End If
	Response.Write strThead
	Dim arrField : arrField = Split(Trim(GetStatisTableField()), "||")
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
				Response.Write "<td>" & rsTmp("scYear")-1 & "-" & rsTmp("scYear") & "</td>"
				Response.Write "<td>" & SumNum & "</td>"
				If UserRank>0 Then
					Response.Write "<td>" & TotalNum & "</td>"
					Response.Write "<td>" & rsTmp("Grade") & "</td>"
				End If
				For i = 0 To Ubound(arrField)
					ValueNum = rsTmp(arrField(i))
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
	
	'Response.Write "<tr><td colspan=""9"">" & Ubound(arrField) & "</td>"		'所有统计字段名
	'For i=0 To Ubound(arrField)
	'	Response.Write "<td>" & arrField(i) & "</td>"
	'Next
	'Response.Write "</tr>"
	Response.Write "</tbody></table>" & vbCrlf
	Response.Flush
	If Not(toExcel) Then
		pageUrl = Action & ".html?total=" & tShowTotal & "&teacher=" & soTeacher & "&sort=2&ksdm=" & soKSDM & "&limit=" & tLimit
		Response.Write "	<div class=""pageTips"">共" & TotalRecord & "条记录　页数：" & tPage & "/" & TotalPage & "页</div>" & vbCrlf
		Response.Write "</div>" & vbCrlf


		'----------更新业绩报表测试
		'通过项目ID更新
		Dim tTemplate, tTableName, tStuType, arrStuType, iPoint, iCredit
		Set rs = Conn.Execute("Select * From HR_Class Where ClassID=1000")
			If Not(rs.BOF And rs.EOF) Then
				tTableName = "HR_Sheet_" & rs("ClassID")
				tStuType = FilterArrNull(Trim(rs("StudentType")), ",")
				If ChkTable(tTableName) Then			'检查数据表是否存在
					If HR_IsNull(tStuType) = False Then		'有学生类别时
						arrStuType = Split(tStuType, ",")
					Else
						Set rsTmp = Conn.Execute("Select Top 10 a.* From " & tTableName & " a Where a.VA1>0")
							If Not(rsTmp.BOF And rsTmp.EOF) Then
								Do While Not rsTmp.EOF
									Response.Write "<li>" & rsTmp("VA1") & "</li>"
									rsTmp.MoveNext
								Loop
							Else
								Response.Write "<li>无数据</li>"
							End If
						Set rsTmp = Nothing
					End If
				End If
			End If
		Set rs = Nothing
		If Action="Mine" Then Response.Write "<li>" & ChkTeacherKPI(UserYGDM) & "</li>"
		Dim runTime : runTime = Timer - BeginTime
		'Response.Write "<hr>" & sqlTmp
		'Response.Write "<hr>" & runTime
		Response.Flush

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
		tmpHtml = tmpHtml & "				location.href=""" & Action & ".html?total=""+ sototal +""&teacher=""+ soteacher +""&sort="" + soSort + ""&ksdm="" + soksdm + ""&limit="" + solimit + ""&soyear="" + soyear ;" & vbCrlf
		tmpHtml = tmpHtml & "			} else if (btnEvent == ""prev"") {" & vbCrlf
		tmpHtml = tmpHtml & "				location.href=""" & pageUrl & "&page=" & PrevPage & """;" & vbCrlf
		tmpHtml = tmpHtml & "			} else if (btnEvent == ""next"") {" & vbCrlf
		tmpHtml = tmpHtml & "				location.href=""" & pageUrl & "&page=" & NextPage & """;" & vbCrlf
		tmpHtml = tmpHtml & "			} else if (btnEvent == ""export"") {" & vbCrlf
		tmpHtml = tmpHtml & "				layer.confirm('本操作将导出所有员工的业绩报表！<br>导出的文件会自动保存到“我的电脑”的“下载”目录中<br>文件名为：" & tExcelFileName & "', {icon:3,title: ""导出报表""}, function(index){" & vbCrlf
		tmpHtml = tmpHtml & "					location.href=""" & Action & ".html?excel=true&total=" & tShowTotal & "&teacher=" & soTeacher & "&sort=" & tSort & "&ksdm=" & soKSDM & "&soyear=" & soYear & "&limit=10000"";" & vbCrlf
		tmpHtml = tmpHtml & "					layer.close(layer.index);" & vbCrlf
		tmpHtml = tmpHtml & "				});" & vbCrlf
		tmpHtml = tmpHtml & "			}" & vbCrlf
		tmpHtml = tmpHtml & "		});" & vbCrlf
		tmpHtml = tmpHtml & "		$("".navBtn"").on(""click"",function(){" & vbCrlf
		tmpHtml = tmpHtml & "			location.href = """ & ParmPath & "Help.html?file=helpAchieve.pdf"";" & vbCrlf
		tmpHtml = tmpHtml & "		});" & vbCrlf
		tmpHtml = tmpHtml & "	});" & vbCrlf
		tmpHtml = tmpHtml & "</script>" & vbCrlf

		strHtml = getPageFoot("Desktop", 1)
		strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
		Response.Write ReplaceCommonLabel(strHtml)
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
%>