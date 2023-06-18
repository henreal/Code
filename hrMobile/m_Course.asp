<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./m_Course_Inc.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "基础性教学"
Dim msgNum : msgNum = 0

Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "List" Call List()
	Case "ListItem" Call ListItem()
	Case "GetTeacher" Call GetTeacherData()
	
	Case "Edit" Call Edit()
	Case "SaveEdit" Call SaveEdit()
	Case "View" Call View()
	Case "ShowCourse" Call ShowCourse()
	Case "ViewAttach" Call ViewAttach()
	Case "SaveAttach" Call SaveAttach()

	Case "ApplyModi" Call ApplyModi()
	Case "ListTeacher" Call ListTeacher()
	Case "getFieldArr" Call getFieldArr()
	Case "getItemGradeArr" Call getItemGradeArr()	'异步取等级
	Case "Affirm" Call Affirm()		'确认提交
	Case "Swap" Call Swap()			'调换课程

	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim TypeID : TypeID = HR_Clng(Request("TypeID"))
	If TypeID = 2 Then SiteTitle = "激励性教学"
End Sub

Sub ListItem()		'项目列表
	Dim TypeID : TypeID = HR_Clng(Request("TypeID"))
	If TypeID = 2 Then SiteTitle = "激励性教学"

	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	'strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	'strHtml = strHtml & "		.weui-grid {padding:10px;}" & vbCrlf
	'strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Site_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))

	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-panel-item hr-fix"">" & vbCrlf
	Response.Write ShowNavMenu(TypeID)
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".hr-navmenu-main"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var tnext = $(this).next("".nav-child"");tnext.toggle();" & vbCrlf
	strHtml = strHtml & "		var dis = tnext.css(""display"");" & vbCrlf
	strHtml = strHtml & "		if(dis == ""block""){ $(this).find("".more i"").html(""&#xea45;"");}else{$(this).find("".more i"").html(""&#xea44;"");}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".addNew a"").css(""display"",""none"");" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub List()		'课程业绩列表
	Dim tTypeID, tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tEduYear : tEduYear = HR_CLng(Request("EduYear"))

	Dim IsAdd : IsAdd = HR_CBool(XmlText("Common", "AddSwitch", "0"))		'录入开关
	If tEduYear = 0 Then tEduYear = DefYear
	ErrMsg = ""

	Dim tItemName, tTemplate, tSheetName, lenField, tFieldHead, tArrHead, tUnit, tEduGrade
	Set rsTmp = Conn.Execute("Select Top 1 ClassName,Unit,SheetName,FieldLen, FieldHead, Template, ClassType From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tTypeID = HR_Clng(rsTmp("ClassType"))
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			lenField = HR_Clng(rsTmp("FieldLen"))
			tFieldHead = Trim(rsTmp("FieldHead"))
			tUnit = Trim(rsTmp("Unit"))
			tSheetName = "HR_Sheet_" & tItemID
			If Not(ChkTable(tSheetName)) Then ErrMsg = "未找到数据表 " & tSheetName & "！<br>"
		Else
			ErrMsg = "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing
	If tFieldHead <> "" Then
		tFieldHead = FilterArrNull(tFieldHead, ",")
		tArrHead = Split(tFieldHead, ",")
		If Ubound(tArrHead) <> lenField Then Redim Preserve tArrHead(lenField)
	Else
		Redim tArrHead(lenField)
	End If

	'----- 等级
	Set rsTmp = Conn.Execute("Select Top 1 Grade From HR_KPI_SUM Where YGDM>0 And YGDM=" & HR_Clng(UserYGDM) & " And scYear=" & DefYear)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tEduGrade = Trim(rsTmp("Grade"))
		End If
	Set rsTmp = Nothing
	If ErrMsg<>"" Then Response.Write GetErrBody(0) : Exit Sub
	SiteTitle = tItemName

	Dim rsList, sqlList, strList, tKSMC
	Dim tVA4, tVA7, tmpTime
	Dim tmpYGDM : tmpYGDM = HR_Clng(Request("SearchWord"))
	If HR_Clng(UserYGDM) > 0 Then tmpYGDM = HR_Clng(UserYGDM)

	sqlList = "Select Top 300 * From " & tSheetName & " Where VA1>0"
	If tmpYGDM > 0 Then sqlList = sqlList & " And VA1=" & tmpYGDM
	If DefYear > 2000 Then sqlList = sqlList & " And scYear=" & tEduYear
	sqlList = sqlList & " Order By AppendTime DESC"
	Set rsList = Server.CreateObject("ADODB.RecordSet")
		rsList.Open(sqlList), Conn, 1, 1
		If Not(rsList.BOF And rsList.EOF) Then
			Do While Not rsList.EOF
				tVA4 = Trim(rsList("VA4"))
				If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Then
					If HR_Clng(tVA4) > 0 Then tVA4 = FormatDate(ConvertNumDate(tVA4), 2)
				End If

				If tTemplate = "TempTableA" Then
					tVA7 = Trim(rsList("VA7"))
					tmpTime = GetPeriodTime(Trim(rsList("VA11")), tVA7, 0)		'计算节次时间
				End If
				tKSMC = Trim(strGetTypeName("HR_Teacher", "KSMC", "YGDM", rsList("VA1")))
				strList = strList & "<ul class=""listPanel"">" & vbCrlf
				If tTemplate = "TempTableA" Then
					strList = strList & "	<li class=""title"">" & rsList("VA9") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(8) & "：</b>" & rsList("VA8") & "　<b>校区：</b>" & rsList("VA11") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>姓名：</b>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(10) & "：</b>" & rsList("VA10") & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>日期：</b>" & tVA4 & " 星期" & rsList("VA6") & "　<b>周次：</b>" & rsList("VA5") & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>节次：</b>" & rsList("VA7") & "　<b>时间：</b>" & tmpTime & "</li>" & vbCrlf
				ElseIf tTemplate = "TempTableB" Then
					strList = strList & "	<li class=""title"">" & rsList("VA5") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>姓名：</b>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>" & tArrHead(4) & "：</b>" & tVA4 & "　<b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>" & tArrHead(5) & "：</b>" & rsList("VA5") & "</li>" & vbCrlf
				ElseIf tTemplate = "TempTableC" Then
					strList = strList & "	<li class=""title"">" & rsList("VA6") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info"">" & rsList("VA7") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>姓名：</b>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>" & tArrHead(4) & "：</b>" & tVA4 & "　<b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>" & tArrHead(5) & "：</b>" & rsList("VA5") & "</li>" & vbCrlf
				ElseIf tTemplate = "TempTableD" Then
					strList = strList & "	<li class=""title"">" & rsList("VA6") & "</li>" & vbCrlf
					If Trim(rsList("VA8")) <> "" Then strList = strList & "	<li class=""info"">" & rsList("VA8") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(7) & "：</b>" & rsList("VA7") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>姓名：</b>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>" & tArrHead(4) & "：</b>" & tVA4 & "　<b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>" & tArrHead(5) & "：</b>" & rsList("VA5") & "</li>" & vbCrlf
				ElseIf tTemplate = "TempTableE" Then
					strList = strList & "	<li class=""title"">" & rsList("VA6") & "</li>" & vbCrlf
					If Trim(rsList("VA9")) <> "" Then strList = strList & "	<li class=""info"">" & rsList("VA9") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(7) & "：</b>" & rsList("VA7") & "　<b>" & tArrHead(8) & "：</b>" & rsList("VA8") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>姓名：</b>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>" & tArrHead(4) & "：</b>" & tVA4 & "　<b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>" & tArrHead(5) & "：</b>" & rsList("VA5") & "</li>" & vbCrlf
				ElseIf tTemplate = "TempTableF" Then
					strList = strList & "	<li class=""title"">" & rsList("VA5") & "</li>" & vbCrlf
					If Trim(rsList("VA8")) <> "" Then strList = strList & "	<li class=""info"">" & rsList("VA8") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(6) & "：</b>" & rsList("VA6")
					If Trim(rsList("VA7")) <> "" Then strList = strList & "　<b>" & tArrHead(7) & "：</b>" & rsList("VA7")
					strList = strList & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>姓名：</b>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>" & tArrHead(4) & "：</b>" & tVA4 & "　<b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</li>" & vbCrlf
				ElseIf tTemplate = "TempTableG" Then
					strList = strList & "	<li class=""title"">" & rsList("VA5") & "</li>" & vbCrlf
					If Trim(rsList("VA7")) <> "" Then strList = strList & "	<li class=""info"">" & rsList("VA7") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(6) & "：</b>" & rsList("VA6") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>姓名：</b>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>" & tArrHead(4) & "：</b>" & tVA4 & "　<b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</li>" & vbCrlf
				End If
				strList = strList & "	<li class=""time""><b>状态：</b>"
				If HR_Clng(rsList("Retreat")) = 1 Then strList = strList & "[退回]"
				If HR_Clng(rsList("State")) = 1 Then strList = strList & "[已确认]" Else strList = strList & "[未确认]"
				If HR_CBool(rsList("Passed")) Then strList = strList & "[已审]" Else strList = strList & "[未审]"
				strList = strList & "</li>" & vbCrlf

				strList = strList & "	<li class=""more"">" & vbCrlf
				'strList = strList & "		<a href=""" & ParmPath & "Course/Swap.html?ItemID=" & tItemID & "&ID=" & rsList("ID") & """ class=""weui-btn weui-btn_mini weui-btn_plain-primary"">调换课</a>" & vbCrlf
				strList = strList & "		<a href=""" & ParmPath & "Course/View.html?ItemID=" & tItemID & "&ID=" & rsList("ID") & """ class=""weui-btn weui-btn_mini weui-btn_plain-primary"">详情</a>" & vbCrlf
				strList = strList & "	</li>" & vbCrlf
				strList = strList & "</ul>" & vbCrlf
				rsList.MoveNext
			Loop
		Else
			strList = "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>" & tEduYear-1 & "-" & tEduYear & "学年暂时没有课程数据！</h3></div>" & vbCrlf
		End If
	Set rsList = Nothing

	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		body {background-color:#fff;}" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn {width:55px;height:55px;text-align:center;font-size:2.6rem;position: fixed;right: 1rem;bottom: 3rem;z-index: 99;}" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn i {color:#2196f3;}" & vbCrlf
	strHtml = strHtml & "		.ShowCourse {display:none;}" & vbCrlf
	strHtml = strHtml & "		.hr-inline {border-bottom:1px solid #e3e3e3;padding:8px;background:#eee;}" & vbCrlf
	strHtml = strHtml & "		.hr-inline em:first-child {flex-shrink:0;color:#999;}" & vbCrlf
	strHtml = strHtml & "		.grade {color:#999;}" & vbCrlf
	strHtml = strHtml & "		.grade b {color:#f30;}" & vbCrlf
	strHtml = strHtml & "		.hr-inline .yearbar {width:140px; border-radius:30px;background:#fff; position:relative; padding-right:10px}" & vbCrlf
	strHtml = strHtml & "		.hr-inline .yearbar .yearinput {border:0; width:100%; text-align:center; font-size:18px;}" & vbCrlf
	strHtml = strHtml & "		.hr-inline .yearbar span {color:#f30; position:absolute; top:-2px; right:7px;font-size:18px;}" & vbCrlf
	strHtml = strHtml & "		.hr-inline em, .hr-inline tt {font-size:0.8rem;}" & vbCrlf
	strHtml = strHtml & "		.viewPanel li b {width:auto;}" & vbCrlf
	strHtml = strHtml & "		.listPanel .more {position:relative;top:3px;text-align:left;padding-top:5px;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Site_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))

	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-fix"">" & vbCrlf
	Response.Write "	<div class=""hr-rows hr-inline"">" & vbCrlf
	Response.Write "		<em class=""hr-item hr-fixed"">学时数</em><em class=""hr-item"">"
	sqlTmp = "Select Sum(VA3) From " & tSheetName & " Where VA1>0 And VA1=" & HR_Clng(UserYGDM) & " And Passed=" & HR_True & " And scYear=" & tEduYear
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Response.Write "" & HR_CDbl(rsTmp(0)) & ""
		Else
			Response.Write 0
		End If
	Set rsTmp = Nothing
	Response.Write "</em><em class=""hr-item hr-grow grade"">等级:<b>" & tEduGrade & "</b></em>" & vbCrlf
	Response.Write "		<tt class=""hr-item hr-fixed"">学年</tt>" & vbCrlf
	Response.Write "		<tt class=""hr-item hr-fixed yearbar""><input name=""EduYear"" id=""EduYear"" class=""yearinput"" value=""" & tEduYear-1 & "-" & tEduYear & """ readonly /><span><i class=""hr-icon"">&#xf0d7;</i></span></tt>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-float-btn"">" & vbCrlf
	If IsAdd Then
		Response.Write "	<a href=""" & ParmPath & "Course/Edit.html?AddNew=True&ItemID=" & tItemID & """ class=""addBtn""><i class=""hr-icon"">&#xf3c0;</i></a>" & vbCrlf
	Else
		Response.Write "	<a href=""javascript:;"" class=""addBtn""><i class=""hr-icon"" style=""color:#aaa"">&#xf3c0;</i></a>" & vbCrlf
	End If
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-sides-x10 hr-fix"">" & vbCrlf
	Response.Write " " & strList
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".hr-navmenu-main"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var tnext = $(this).next("".nav-child"");tnext.toggle();" & vbCrlf
	strHtml = strHtml & "		var dis = tnext.css(""display"");" & vbCrlf
	strHtml = strHtml & "		if(dis == ""block""){ $(this).find("".more i"").html(""&#xea45;"");}else{$(this).find("".more i"").html(""&#xea44;"");}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Course/ListItem.html?TypeID=" & tTypeID & """; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".addNew a"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		location.href=""" & ParmPath & "Course/Edit.html?AddNew=True&ItemID=" & tItemID & """;" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	Dim tmpYearJson, cYear : cYear = Year(Date()) + 1
	tmpYearJson = ""
	For i = cYear To cYear-7 Step -1
		If i<cYear Then tmpYearJson = tmpYearJson & ","
		tmpYearJson = tmpYearJson & "{title:""" & i-1 & "-" & i & """, value:""" & i & """}"
	Next
	strHtml = strHtml & "	$(""#EduYear"").select({title:""选择学年""," & vbCrlf
	strHtml = strHtml & "		items:[" & tmpYearJson & "]," & vbCrlf
	strHtml = strHtml & "		onClose:function(res){" & vbCrlf
	strHtml = strHtml & "			location.href=""" & ParmPath & "Course/List.html?ItemID=" & tItemID & "&EduYear="" + res.data.values;" & vbCrlf
	'strHtml = strHtml & "			console.log(res.data.values);" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub View()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	ErrMsg = ""

	Dim tItemName, tTemplate, tSheetName, lenField, tFieldHead, tArrHead, tUnit
	Set rsTmp = Conn.Execute("Select Top 1 ClassName,Unit,SheetName,FieldLen,FieldHead,Template From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			lenField = HR_Clng(rsTmp("FieldLen"))
			tFieldHead = Trim(rsTmp("FieldHead"))
			tUnit = Trim(rsTmp("Unit"))
			tSheetName = "HR_Sheet_" & tItemID
			If Not(ChkTable(tSheetName)) Then ErrMsg = "未找到数据表 " & tSheetName & "！<br>"
		Else
			ErrMsg = "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing
	If tFieldHead <> "" Then
		tFieldHead = FilterArrNull(tFieldHead, ",")
		tArrHead = Split(tFieldHead, ",")
		If Ubound(tArrHead) <> lenField Then Redim Preserve tArrHead(lenField)
	Else
		Redim tArrHead(lenField)
	End If

	If ErrMsg<>"" Then Response.Write GetErrBody(0) : Exit Sub
	SiteTitle = tItemName & " 详情"

	Dim rsList, sqlList, strList, tKSMC, tPRZC
	Dim tVA4, tVA6, tVA7, tmpTime, tPassed, tState
	sqlList = "Select * From " & tSheetName & " Where ID=" & tmpID
	Set rsList = Server.CreateObject("ADODB.RecordSet")
		rsList.Open(sqlList), Conn, 1, 1
		If Not(rsList.BOF And rsList.EOF) Then
			Do While Not rsList.EOF
				tVA4 = Trim(rsList("VA4"))
				tPassed = HR_CBool(rsList("Passed"))
				tState = HR_Clng(rsList("State"))

				If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Then
					If HR_Clng(tVA4) > 0 Then tVA4 = FormatDate(ConvertNumDate(tVA4), 2)
				End If

				If tTemplate = "TempTableA" Then
					tVA7 = Trim(rsList("VA7"))
					tVA6 = " 星期" & rsList("VA6")
					tmpTime = GetPeriodTime(Trim(rsList("VA11")), tVA7, 0)		'计算节次时间
				End If
				tKSMC = Trim(strGetTypeName("HR_Teacher", "KSMC", "YGDM", rsList("VA1")))
				tPRZC = Trim(strGetTypeName("HR_Teacher", "PRZC", "YGDM", rsList("VA1")))

				strList = strList & "<ul class=""viewPanel"">" & vbCrlf
				strList = strList & "	<li class=""info""><b>" & tArrHead(2) & "：</b>" & rsList("VA2") & " [" & rsList("VA1") & "]</li>" & vbCrlf
				strList = strList & "	<li class=""info""><b>科室：</b>" & tKSMC & "</li>" & vbCrlf
				strList = strList & "	<li class=""info""><b>职称：</b>" & tPRZC & "</li>" & vbCrlf
				strList = strList & "	<li class=""hr-gap-20""></li>" & vbCrlf
				strList = strList & "	<li class=""time""><b>" & tArrHead(4) & "：</b>" & tVA4 & tVA6 & "</li>" & vbCrlf
				strList = strList & "	<li class=""info""><b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</li>" & vbCrlf

				If tTemplate = "TempTableA" Then
					strList = strList & "	<li class=""info""><b>" & tArrHead(8) & "：</b>" & rsList("VA8") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(9) & "：</b>" & rsList("VA9") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(10) & "：</b>" & rsList("VA10") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(11) & "：</b>" & rsList("VA11") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(12) & "：</b>" & rsList("VA12") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(5) & "：</b>" & rsList("VA5") & "</li>" & vbCrlf
					strList = strList & "	<li class=""time""><b>节次：</b>" & rsList("VA7") & "<b>时间：</b>" & tmpTime & "</li>" & vbCrlf
				ElseIf tTemplate = "TempTableB" Then

					strList = strList & "	<li class=""info""><b>" & tArrHead(5) & "：</b>" & rsList("VA5") & "</li>" & vbCrlf
					If Trim(rsList("VA6")) <> "" Then strList = strList & "	<li class=""info""><b>" & tArrHead(6) & "：</b>" & rsList("VA6") & "</li>" & vbCrlf		'备注
					
				ElseIf tTemplate = "TempTableC" Then
					strList = strList & "	<li class=""info""><b>" & tArrHead(6) & "：</b>" & rsList("VA6") & "</li>" & vbCrlf
					If Trim(rsList("VA7")) <> "" Then strList = strList & "	<li class=""info""><b>" & tArrHead(7) & "：</b>" & rsList("VA7") & "</li>" & vbCrlf		'备注
					strList = strList & "	<li class=""info""><b>" & tArrHead(5) & "：</b>" & rsList("VA5") & "</li>" & vbCrlf
				ElseIf tTemplate = "TempTableD" Then
					strList = strList & "	<li class=""info""><b>" & tArrHead(6) & "：</b>" & rsList("VA6") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(7) & "：</b>" & rsList("VA7") & "</li>" & vbCrlf
					If Trim(rsList("VA8")) <> "" Then strList = strList & "	<li class=""info""><b>" & tArrHead(8) & "：</b>" & rsList("VA8") & "</li>" & vbCrlf		'备注
				ElseIf tTemplate = "TempTableE" Then
					strList = strList & "	<li class=""info""><b>" & tArrHead(6) & "：</b>" & rsList("VA6") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(7) & "：</b>" & rsList("VA7") & "</li>" & vbCrlf
					If Trim(rsList("VA8")) <> "" Then strList = strList & "	<li class=""info""><b>" & tArrHead(8) & "：</b>" & rsList("VA8") & "</li>" & vbCrlf		'等级
					strList = strList & "	<li class=""info""><b>" & tArrHead(5) & "：</b>" & rsList("VA5") & "</li>" & vbCrlf
				ElseIf tTemplate = "TempTableF" Then
					strList = strList & "	<li class=""info""><b>" & tArrHead(5) & "：</b>" & rsList("VA5") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(6) & "：</b>" & rsList("VA6") & "</li>" & vbCrlf
					If Trim(rsList("VA7")) <> "" Then strList = strList & "	<li class=""info""><b>" & tArrHead(7) & "：</b>" & rsList("VA7") & "</li>" & vbCrlf		'等级
					If Trim(rsList("VA8")) <> "" Then strList = strList & "	<li class=""info""><b>" & tArrHead(8) & "：</b>" & rsList("VA8") & "</li>" & vbCrlf		'备注
				ElseIf tTemplate = "TempTableG" Then
					strList = strList & "	<li class=""info""><b>" & tArrHead(5) & "：</b>" & rsList("VA5") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><b>" & tArrHead(6) & "：</b>" & rsList("VA6") & "</li>" & vbCrlf
					If Trim(rsList("VA7")) <> "" Then strList = strList & "	<li class=""info""><b>" & tArrHead(7) & "：</b>" & rsList("VA7") & "</li>" & vbCrlf
				End If
				strList = strList & "	<li class=""info""><b>状态：</b>"
				If tPassed Then
					strList = strList & "【已审】" & vbCrlf
				Else
					strList = strList & "【未审】" & vbCrlf
				End If
				If tState=0 Then
					strList = strList & "【未确认】" & vbCrlf
				Else
					strList = strList & "【已确认】" & vbCrlf
				End If
				strList = strList & "</li>" & vbCrlf
				strList = strList & "</ul>" & vbCrlf
				strList = strList & "<div class=""btnBar""><button data-id=" & rsList("ID") & " data-item=" & tItemID & " class=""weui-btn weui-btn_primary btnManage"">操作</button></div>" & vbCrlf
				rsList.MoveNext
			Loop
		Else
			strList = "没有数据"
		End If
	Set rsList = Nothing

	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.viewPanel li.hr-gap-20 {background-color:#eee;}" & vbCrlf
	strHtml = strHtml & "		.btnBar {box-sizing:border-box;padding:10px;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-fix"">" & vbCrlf
	Response.Write " " & strList
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	FastClick.attach(document.body);" & vbCrlf
	strHtml = strHtml & "	$("".btnManage"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		$.actions({" & vbCrlf
	strHtml = strHtml & "			title: ""选择操作"",onClose: function() { console.log(""close"") },actions:[" & vbCrlf
	strHtml = strHtml & "				{text: ""查看附件"",className: ""color-danger"",onClick: function(){ location.href=""" & ParmPath & "Course/ViewAttach.html?ItemID=" & tItemID & "&ID=" & tmpID & """ }}" & vbCrlf
	If Not(tPassed) Then
		strHtml = strHtml & "				,{text: ""修改"",className:""color-primary"",onClick: function(){location.href=""" & ParmPath & "Course/Edit.html?ItemID=" & tItemID & "&ID=" & tmpID & """}}" & vbCrlf
		strHtml = strHtml & "				,{text: ""删除"",className: ""color-warning"",onClick: function() { $.alert(""删除暂时关闭，正在核对业绩生成！""); }}" & vbCrlf
	Else
		strHtml = strHtml & "				,{text: ""申请修改"",className: ""color-warning"",onClick: function(){" & vbCrlf
		strHtml = strHtml & "					location.href=""" & ParmPath & "Course/ApplyModi.html?ItemID=" & tItemID & "&ID=" & tmpID & """" & vbCrlf
		strHtml = strHtml & "				}}" & vbCrlf
	End If
	If tState = 0 Then
		strHtml = strHtml & "				,{text: ""确认提交"",className: ""color-success"",onClick: function(){" & vbCrlf
		strHtml = strHtml & "					$.getJSON(""" & ParmPath & "Course/Affirm.html"",{ItemID:" & tItemID & ", ID:" & tmpID & "}, function(reData){" & vbCrlf
		strHtml = strHtml & "						$.alert(reData.reMessge, reData.reTitle);" & vbCrlf
		strHtml = strHtml & "					});" & vbCrlf
		strHtml = strHtml & "				}}" & vbCrlf
	End If
	
	strHtml = strHtml & "			]" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	strHtml = strHtml & "	$("".addNew a"").css(""display"",""none"");" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

Sub Edit()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim SubButTxt : ErrMsg = "" : SubButTxt = "添加"
	Dim IsModify : IsModify = False

	Dim tItemName, tTemplate, strStuType, tSheetName, tFieldLen, tFieldHead, tArrHead, tUnit
	Set rsTmp = Conn.Execute("Select Top 1 ClassName,Unit,SheetName,FieldLen,FieldHead,Template,StudentType From HR_Class Where ClassID=" & tItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			strStuType = Trim(rsTmp("StudentType"))
			tFieldLen = HR_Clng(rsTmp("FieldLen"))
			tFieldHead = Trim(rsTmp("FieldHead"))
			tUnit = Trim(rsTmp("Unit"))
			tSheetName = "HR_Sheet_" & tItemID
			If Not(ChkTable(tSheetName)) Then ErrMsg = "未找到数据表 " & tSheetName & "！<br>"
		Else
			ErrMsg = "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing
	If tFieldHead <> "" Then
		tFieldHead = FilterArrNull(tFieldHead, ",")
		tArrHead = Split(tFieldHead, ",")
		If Ubound(tArrHead) <> tFieldLen Then Redim Preserve tArrHead(tFieldLen)
	Else
		Redim tArrHead(tFieldLen)
	End If

	If ErrMsg<>"" Then Response.Write GetErrBody(0) : Exit Sub

	Dim sqlEdit, rsEdit, strList, tStuType, tAttachFile, tPassed
	Dim strField, arrField : Redim arrField(tFieldLen)

	sqlEdit = "Select * From " & tSheetName & " Where ID=" & tmpID
	Set rsEdit = Server.CreateObject("ADODB.RecordSet")
		rsEdit.Open(sqlEdit), Conn, 1, 1
		If Not(rsEdit.BOF And rsEdit.EOF) Then
			IsModify = True
			SubButTxt = "修改"
			tStuType = Trim(rsEdit("StudentType"))
			tAttachFile = Trim(rsEdit("Explain"))
			tAttachFile = FilterArrNull(tAttachFile, "|")
			tPassed = HR_CBool(rsEdit("Passed"))
			For i = 0 To tFieldLen-1
				arrField(i) = rsEdit("VA" & i)
			Next
		Else
			If UserYGDM <> "" Then arrField(1) = UserYGDM
			If UserYGXM <> "" Then arrField(2) = UserYGXM
		End If
	Set rsEdit = Nothing


	Dim  strVA4, IsDate, tVA4 : tVA4 = Trim(arrField(4))
	If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Then
		tVA4 = FormatDate(ConvertNumDate(tVA4), 2)			'转为日期
		strVA4 = "<input name=""VA4"" class=""weui-input"" id=""VA4"" type=""text"" value=""" & tVA4 & """ placeholder=""" & tArrHead(4) & """>"
		IsDate = True
	Else
		strVA4 = "<input name=""VA4"" class=""weui-input"" id=""VA4"" type=""text"" value=""" & tVA4 & """ placeholder=""点击选择" & tArrHead(4) & """>"
		IsDate = False
	End If

	SiteTitle = SubButTxt & "课程内容"

			strList = strList & "<div class=""weui-cells weui-cells_form"">" & vbCrlf
			strList = strList & "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">姓　名：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""ygxm"" class=""weui-input"" id=""ygxm"" type=""text"" value=""" & arrField(2) & """ data-key=""ygxm"" data-value=""ygdm"" placeholder="""" readonly>" & vbCrlf
			strList = strList & "			<input name=""ygdm"" class=""weui-input"" id=""ygdm"" type=""hidden"" value=""" & arrField(1) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			If HR_IsNull(strStuType) = False Then		'有学生类别
				strList = strList & "	<div class=""weui-cell"">" & vbCrlf
				strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">学生类别：</label></div>" & vbCrlf
				strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
				strList = strList & "			<input name=""StudentType"" class=""weui-input"" id=""StuType"" type=""text"" value=""" & tStuType & """ placeholder=""学生类别"""
				If IsModify Then strList = strList & " disabled"
				strList = strList & ">" & vbCrlf
				strList = strList & "		</div>" & vbCrlf
				strList = strList & "	</div>" & vbCrlf
			End If

			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(3) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA3"" class=""weui-input"" id=""VA3"" type=""number"" value=""" & arrField(3) & """ placeholder="""">" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__ft Unit"">" & tUnit & "</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf

			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(4) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			" & strVA4 & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf

		If tTemplate = "TempTableA" Then
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(6) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA6"" class=""weui-input"" id=""VA6"" type=""text"" value=""" & arrField(6) & """ placeholder=""星期"">" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf

			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(5) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__ft"">" & vbCrlf
			strList = strList & "			<div class=""weui-count"">" & vbCrlf
			strList = strList & "				<a class=""weui-count__btn weui-count__decrease""></a>" & vbCrlf
			strList = strList & "				<input name=""VA5"" class=""weui-count__number"" id=""VA5"" type=""number"" value=""" & HR_Clng(arrField(5)) & """>" & vbCrlf
			strList = strList & "				<a class=""weui-count__btn weui-count__increase""></a>" & vbCrlf
			strList = strList & "			</div>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(7) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA7"" class=""weui-input opt1"" id=""VA7"" type=""text"" value=""" & arrField(7) & """ placeholder=""格式:3-4"">" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(8) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA8"" class=""weui-input opt1"" id=""VA8"" type=""text"" value=""" & arrField(8) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(9) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA9"" class=""weui-input opt1"" id=""VA9"" type=""text"" value=""" & arrField(9) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(10) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA10"" class=""weui-input opt1"" id=""VA10"" type=""text"" value=""" & arrField(10) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(11) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA11"" class=""weui-input opt1"" id=""VA11"" type=""text"" value=""" & arrField(11) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(12) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA12"" class=""weui-input opt1"" id=""VA12"" type=""text"" value=""" & arrField(12) & """ placeholder="" "">" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
		ElseIf tTemplate = "TempTableB" Then
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(5) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA5"" class=""weui-input"" id=""VA5"" type=""text"" value=""" & arrField(5) & """ placeholder=""" & tArrHead(5) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(6) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA6"" class=""weui-input"" id=""VA6"" type=""text"" value=""" & arrField(6) & """ placeholder=""" & tArrHead(6) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
		ElseIf tTemplate = "TempTableC" Then
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(5) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA5"" class=""weui-input"" id=""VA5"" type=""text"" value=""" & arrField(5) & """ placeholder=""" & tArrHead(5) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(6) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA6"" class=""weui-input"" id=""VA6"" type=""text"" value=""" & arrField(6) & """ placeholder=""" & tArrHead(6) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(7) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA7"" class=""weui-input"" id=""VA7"" type=""text"" value=""" & arrField(7) & """ placeholder=""" & tArrHead(7) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
		ElseIf tTemplate = "TempTableD" Then
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(5) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA5"" class=""weui-input"" id=""VA5"" type=""text"" value=""" & arrField(5) & """ placeholder=""" & tArrHead(5) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(6) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA6"" class=""weui-input"" id=""VA6"" type=""text"" value=""" & arrField(6) & """ placeholder=""" & tArrHead(6) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(7) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA7"" class=""weui-input"" id=""VA7"" type=""text"" value=""" & arrField(7) & """ placeholder=""" & tArrHead(7) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(8) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA8"" class=""weui-input"" id=""VA8"" type=""text"" value=""" & arrField(8) & """ placeholder=""" & tArrHead(8) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
		ElseIf tTemplate = "TempTableE" Then
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(5) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA5"" class=""weui-input"" id=""VA5"" type=""text"" value=""" & arrField(5) & """ placeholder=""" & tArrHead(5) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(6) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA6"" class=""weui-input"" id=""VA6"" type=""text"" value=""" & arrField(6) & """ placeholder=""" & tArrHead(6) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(7) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA7"" class=""weui-input"" id=""VA7"" type=""text"" value=""" & arrField(7) & """ data-values="""" placeholder=""" & tArrHead(7) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(8) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA8"" class=""weui-input"" id=""VA8"" type=""text"" value=""" & arrField(8) & """ placeholder=""" & tArrHead(8) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(9) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA9"" class=""weui-input"" id=""VA9"" type=""text"" value=""" & arrField(9) & """ placeholder=""" & tArrHead(9) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
		ElseIf tTemplate = "TempTableF" Then
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(5) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA5"" class=""weui-input"" id=""VA5"" type=""text"" value=""" & arrField(5) & """ placeholder=""" & tArrHead(5) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(6) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA6"" class=""weui-input"" id=""VA6"" type=""text"" value=""" & arrField(6) & """ placeholder=""" & tArrHead(6) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(7) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA7"" class=""weui-input"" id=""VA7"" type=""text"" value=""" & arrField(7) & """ placeholder=""" & tArrHead(7) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
		ElseIf tTemplate = "TempTableG" Then
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(5) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA5"" class=""weui-input"" id=""VA5"" type=""text"" value=""" & arrField(5) & """ placeholder=""" & tArrHead(5) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell"">" & vbCrlf
			strList = strList & "		<div class=""weui-cell__hd""><label class=""weui-label"">" & tArrHead(6) & "：</label></div>" & vbCrlf
			strList = strList & "		<div class=""weui-cell__bd"">" & vbCrlf
			strList = strList & "			<input name=""VA6"" class=""weui-input"" id=""VA6"" type=""text"" value=""" & arrField(6) & """ placeholder=""" & tArrHead(6) & """>" & vbCrlf
			strList = strList & "		</div>" & vbCrlf
			strList = strList & "	</div>" & vbCrlf
		End If

		Dim picArr, strAttach, tArrAttach, AttachNum : AttachNum = 0
		If HR_IsNull(tAttachFile) = False Then
			tArrAttach = Split(tAttachFile, "|")
			AttachNum = Ubound(tArrAttach) + 1
			For i = 0 To Ubound(tArrAttach)
				strAttach = strAttach & "<span class='pic_look' data-img='" & tArrAttach(i) & "' style='background-image: url(" & tArrAttach(i) & ")'><em id='delete_pic'>-</em></span>"
				If i> 0 Then picArr = picArr & ","
				picArr = picArr & """" & tArrAttach(i) & """"
			Next
		End If

		strList = strList & "	<div class=""hr-gap-20""></div>" & vbCrlf
		strList = strList & "	<div class=""celltit""><b>上传附件：</b></div>" & vbCrlf
		strList = strList & "	<div class=""release_up_pic"">" & vbCrlf
		strList = strList & "		<div class=""up_pic"">" & vbCrlf
		strList = strList & "			" & strAttach & vbCrlf
		strList = strList & "			<span id=""chose_pic_btn"" style=""""><input type=""file"" accept=""image/*""></span>" & vbCrlf
		strList = strList & "		</div>" & vbCrlf
		strList = strList & "	</div>" & vbCrlf
		strList = strList & "	<div id=""show1""></div>" & vbCrlf
		strList = strList & "	<input name=""UploadAttach"" id=""AttachFile"" type=""hidden"" value=""" & tAttachFile & """ />" & vbCrlf

			strList = strList & "	<input name=""ItemID"" id=""ItemID"" type=""hidden"" value=""" & tItemID & """><input name=""ID"" id=""ID"" type=""hidden"" value=""" & tmpID & """>" & vbCrlf
			If IsModify Then strList = strList & "	<input name=""Modify"" id=""Modify"" type=""hidden"" value=""True"">" & vbCrlf
			strList = strList & "	<div class=""weui-cell""><button type=""button"" name=""editPost"" class=""weui-btn weui-btn_primary"" id=""editPost"">保存</button>" & vbCrlf
			strList = strList & "</form>" & vbCrlf
			strList = strList & "</div>" & vbCrlf
			strList = strList & "" & vbCrlf
			strList = strList & "" & vbCrlf


	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.itemTit{text-align:center;font-size:1.1rem;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells {font-size:0.85rem;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .popWin, .weui-cells_form .Unit {font-size:0.85rem;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .popWin i {font-size:1.2rem;position:relative;top:2px;color:#4ce}" & vbCrlf
	strHtml = strHtml & "		.weui-label {text-align:right;}" & vbCrlf
	strHtml = strHtml & "		.btnBar {box-sizing:border-box;padding:10px 0;}" & vbCrlf
	strHtml = strHtml & "		#EditForm{width:100%;height:100%;margin:0;padding:0;border:0;}" & vbCrlf
	strHtml = strHtml & "		#listFrame {border:0;}" & vbCrlf
	strHtml = strHtml & "		.weui-count .weui-count__number {width:2rem;}" & vbCrlf
	strHtml = strHtml & "		.hr-gap-20 {box-sizing:border-box;height:10px}" & vbCrlf

	strHtml = strHtml & "	.release_up_pic .tit{padding:12px;font-size:1.4rem;color:#999}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .tit h4{font-weight:400}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .tit h4 em{font-size:1.1rem}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .up_pic{background-color:#fff;padding:15px 12px;font-size:0;margin-left:-3.33333%;padding-bottom:3px;border-bottom:1px solid #e7e7e7;border-top:1px solid #e7e7e7}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .up_pic .pic_look{width:30%;height:80px;display:inline-block;background-size:cover;background-position:center center;background-repeat:no-repeat;box-sizing:border-box;margin-left:3.3333%;margin-bottom:12px;position:relative}" & vbCrlf
	strHtml = strHtml & "	.release_up_pic .up_pic .pic_look em{position:absolute;display:inline-block;width:25px;height:25px;background-color:red;color:#fff;font-size:18px;right:5px;top:5px;text-align:center;line-height:22px;border-radius:50%;font-weight:700}" & vbCrlf
	strHtml = strHtml & "	#chose_pic_btn {width:30%;height:80px;position:relative;display:inline-block;background:#eee url(" & InstallDir & "Static/images/upload.png) center no-repeat;box-sizing:border-box;background-size:30px 30px;border:1px solid #dbdbdb;margin-left:3.3333%;margin-bottom:12px}" & vbCrlf
	strHtml = strHtml & "	#chose_pic_btn input{position:absolute;left:0;top:0;opacity:0;width:100%;height:100%}" & vbCrlf
	strHtml = strHtml & "	.release_btn{padding:0 24px;margin-top:70px}" & vbCrlf
	strHtml = strHtml & "	.release_btn button{width:100%;background-color:#2c87af;font-size:1.4rem;color:#fff;border:0;border-radius:3px;height:45px;outline:0}" & vbCrlf
	strHtml = strHtml & "	.release_btn button.none_btn{background-color:#f2f2f2;color:#2c87af;border:1px solid #2c87af;margin-top:15px}" & vbCrlf
	strHtml = strHtml & "	.upbtn {box-sizing:border-box;padding:10px;} .upbtn em{width:50%;text-align:center;box-sizing:border-box;padding:0 10px;}" & vbCrlf
	strHtml = strHtml & "	#show1 {word-break: break-all;word-wrap: break-word;white-space: pre-wrap;}" & vbCrlf
	strHtml = strHtml & "	#loading1 {display:none;position:absolute;left:0;top:0;background:rgba(0,0,0,0.5) url(" & InstallDir & "Static/layui/css/modules/layer/default/loading-1.gif) center no-repeat;width:100%;height:100%;z-index:1000}" & vbCrlf
	strHtml = strHtml & "	.weui-photo-browser-modal {z-index:1000}" & vbCrlf
	strHtml = strHtml & "	.celltit {padding:5px;}" & vbCrlf
	strHtml = strHtml & "	.celltit b {font-weight:normal;font-size:1.1rem;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf

	'strHtml = strHtml & "		$("".pList"").myList();" & vbCrlf

	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-fix"">" & vbCrlf
	Response.Write "	<div class=""itemTit"">" & tItemName & "</div>" & vbCrlf
	Response.Write "	" & strList
	Response.Write "</div>" & vbCrlf

	strHtml = "	<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/js/swiper.min.js?v=3.3.1""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Upload/localResizeIMG.js""></script>" & vbCrlf
	strHtml = strHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Upload/mobileBUGFix.mini.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf


	strHtml = strHtml & "	var stuType = (""" & XmlText("Common", "StudentType", "") & """).split(""|"");" & vbCrlf
	strHtml = strHtml & "	$(""#StuType"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择学生类别"",items:stuType" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	If tTemplate = "TempTableA" Then
		strHtml = strHtml & "	var sPeriod=[], ePeriod=[];" & vbCrlf
		strHtml = strHtml & "	for(var k=1; k<20; k++){ sPeriod.push(k); ePeriod.push(k+1); }" & vbCrlf
		strHtml = strHtml & "	$(""#VA7"").picker({title: ""请选择节次""," & vbCrlf			'节次
		strHtml = strHtml & "		title: ""请选择节次"",cols:[" & vbCrlf
		strHtml = strHtml & "			{textAlign:'center',values:sPeriod}," & vbCrlf
		strHtml = strHtml & "			{textAlign:'center',values:ePeriod}," & vbCrlf
		strHtml = strHtml & "		]," & vbCrlf
		strHtml = strHtml & "		onClose:function(e){" & vbCrlf
		strHtml = strHtml & "			if(parseInt(e.value[0])>parseInt(e.value[1])){ $.toast('开始节次不能大于结束节次', 'forbidden'); $(""#VA7"").val(""""); return false; }" & vbCrlf
		strHtml = strHtml & "			$(""#VA7"").val(e.value[0] +""-""+ e.value[1]);" & vbCrlf
		strHtml = strHtml & "		}" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
		strHtml = strHtml & "	$(""#VA8"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择"",items:[" & GetCourseSelect("VA8", "") & "]" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
		strHtml = strHtml & "	$(""#VA10"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择"",items:[" & getFieldSelect(tItemID, "VA10", "") & "]" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
		strHtml = strHtml & "	$(""#VA11"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择"",items:[" & GetCampusSelect("VA11", "") & "]" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
		strHtml = strHtml & "	$(""#VA4"").calendar({dateFormat:""yyyy-mm-dd""});" & vbCrlf
		strHtml = strHtml & "	$(""#VA4"").on(""change"",function(){" & vbCrlf
		strHtml = strHtml & "		var today = new Array('日','一','二','三','四','五','六'), day = new Date($(this).val());" & vbCrlf
		strHtml = strHtml & "		var week = today[day.getDay()];$(""#VA6"").val(week);" & vbCrlf
		strHtml = strHtml & "		console.log($(this).val());" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
		

	ElseIf tTemplate = "TempTableD" Then
		strHtml = strHtml & "	var stuSemester = (""" & GetSemesterArr(0,0) & """).split(""|"");" & vbCrlf
		strHtml = strHtml & "	$(""#VA5"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择学年"",items:stuSemester" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf

		strHtml = strHtml & "	$(""#VA4"").calendar({dateFormat:""yyyy-mm-dd""});" & vbCrlf
		strHtml = strHtml & "	$(""#VA7"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择级别"",items:[" & GetSubmoduleSelect(tItemID, "") & "]" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
	ElseIf tTemplate = "TempTableE" Then
		strHtml = strHtml & "	$(""#VA7"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择级别"",items:[" & GetSubmoduleSelect(tItemID, "") & "]" & vbCrlf
		strHtml = strHtml & "		,onClose:function(){" & vbCrlf
		strHtml = strHtml & "			var level = $(""#VA7"").val();" & vbCrlf
		strHtml = strHtml & "			$.getJSON(""" & ParmPath & "Course/getItemGradeArr.html"",{Item:" & tItemID & ",Field:level, value:""""}, function(strForm){" & vbCrlf
		strHtml = strHtml & "				var arrup = strForm.reData;" & vbCrlf
		strHtml = strHtml & "				$(""#VA8"").select(" & vbCrlf
		strHtml = strHtml & "					""update"",{items:arrup}" & vbCrlf
		strHtml = strHtml & "				);" & vbCrlf
		strHtml = strHtml & "			});" & vbCrlf
		strHtml = strHtml & "		}" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
		strHtml = strHtml & "	$(""#VA8"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择等级"",items:[" & GetItemGradeOption(tItemID, arrField(7), arrField(8)) & "]" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
	ElseIf tTemplate = "TempTableF" Or tTemplate = "TempTableG" Then
		strHtml = strHtml & "	$(""#VA6"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择级别"",items:[" & GetSubmoduleSelect(tItemID, "") & "]" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
	End If
	If IsDate=False Then
		strHtml = strHtml & "	var stuSemester = (""" & GetSemesterArr(0,0) & """).split(""|"");" & vbCrlf
		strHtml = strHtml & "	$(""#VA4"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择学年"",items:stuSemester" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
	End If
	strHtml = strHtml & "	var maxNum = 99, minNum = 1;" & vbCrlf
	strHtml = strHtml & "	$('.weui-count__decrease').click(function (e) {" & vbCrlf
	strHtml = strHtml & "		var $input = $(e.currentTarget).parent().find('.weui-count__number');" & vbCrlf
	strHtml = strHtml & "		var number = parseInt($input.val() || ""0"") - 1" & vbCrlf
	strHtml = strHtml & "		if (number < minNum) number = minNum;" & vbCrlf
	strHtml = strHtml & "		$input.val(number)" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf
	strHtml = strHtml & "	$('.weui-count__increase').click(function (e) {" & vbCrlf
	strHtml = strHtml & "		var $input = $(e.currentTarget).parent().find('.weui-count__number');" & vbCrlf
	strHtml = strHtml & "		var number = parseInt($input.val() || ""0"") + 1" & vbCrlf
	strHtml = strHtml & "		if (number > maxNum) number = maxNum;" & vbCrlf
	strHtml = strHtml & "		$input.val(number)" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf

	'strHtml = strHtml & "	FastClick.attach(document.body);" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Course/List.html?ItemID=" & tItemID & """; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$(""#editPost"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "		$.getJSON(""" & ParmPath & "Course/SaveEdit.html"",$(""#EditForm"").serialize(), function(strForm){" & vbCrlf
	strHtml = strHtml & "			$.alert(strForm.reMessge,function(){" & vbCrlf
	strHtml = strHtml & "				if(strForm.Return){location.href=""" & ParmPath & "Course/List.html?ItemID=" & tItemID & "&ID=""+ strForm.id}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	var picArr = new Array(" & picArr & ");" & vbCrlf			'存储图片
	strHtml = strHtml & "	$(""input:file"").localResizeIMG({" & vbCrlf
	strHtml = strHtml & "		width:1920," & vbCrlf			'宽度
	strHtml = strHtml & "		quality: 0.6," & vbCrlf			'压缩参数 1 不压缩 越小清晰度越低
	strHtml = strHtml & "		success: function (result) {" & vbCrlf
	strHtml = strHtml & "			var img = new Image();" & vbCrlf
	strHtml = strHtml & "			img.src = result.base64;" & vbCrlf
	strHtml = strHtml & "			$.showLoading();" & vbCrlf			'上传提示
	strHtml = strHtml & "			$.ajax({" & vbCrlf
	strHtml = strHtml & "				url:""" & InstallDir & "API/UploadBase.htm"",type: ""POST"",data:{formFile:img.src,UploadDir:""Attach""}," & vbCrlf
	strHtml = strHtml & "				dataType: ""HTML"",timeout: 20000,error: function(){alert(""上传超时"");},success: function(reUrl){" & vbCrlf
	strHtml = strHtml & "					var _str = ""<span class='pic_look' data-img='""+ reUrl + ""' style='background-image: url(""+ reUrl + "")'><em id='delete_pic'>-</em></span>""" & vbCrlf
	strHtml = strHtml & "					$('#chose_pic_btn').before(_str);" & vbCrlf
	strHtml = strHtml & "					$.hideLoading();" & vbCrlf				'关闭提示
	strHtml = strHtml & "					var _i =  picArr.length;" & vbCrlf
	strHtml = strHtml & "					picArr[_i] = reUrl;" & vbCrlf
	strHtml = strHtml & "					$('#AttachFile').val(picArr.join(""|""));" & vbCrlf
	strHtml = strHtml & "				}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	// 删除" & vbCrlf
	strHtml = strHtml & "	$(document).on('click', '#delete_pic', function(event){" & vbCrlf
	strHtml = strHtml & "		var aa = $(this).parents("".pic_look"").index();" & vbCrlf
	strHtml = strHtml & "		picArr.splice(aa,1);" & vbCrlf
	strHtml = strHtml & "		$(this).parents("".pic_look"").remove();" & vbCrlf
	strHtml = strHtml & "		console.log(picArr);" & vbCrlf
	strHtml = strHtml & "	});	" & vbCrlf
	strHtml = strHtml & "	$(document).on('click', '.save', function(event){" & vbCrlf
	strHtml = strHtml & "		console.log(picArr);" & vbCrlf
	strHtml = strHtml & "		$.post(""" & ParmPath & "Course/SaveAttach.html"",{pic:picArr.join(""|""), ItemID:" & tItemID & ", ID:" & tmpID & "},function(reStr){" & vbCrlf
	strHtml = strHtml & "			$.toast(reStr.errmsg,function(){" & vbCrlf
	strHtml = strHtml & "				if(!reStr.err){ location.href=""" & ParmPath & "Course/View.html?ItemID=" & tItemID & "&ID=" & tmpID & """; }" & vbCrlf
	strHtml = strHtml & "			});	" & vbCrlf
	strHtml = strHtml & "		},""json"");" & vbCrlf
	strHtml = strHtml & "	});	" & vbCrlf
	strHtml = strHtml & "	$(document).on('click', '.preview', function(event){" & vbCrlf
	strHtml = strHtml & "		var pb1 = $.photoBrowser({ items:[" & picArr & "]});pb1.open(2);" & vbCrlf
	strHtml = strHtml & "		console.log(picArr);" & vbCrlf
	strHtml = strHtml & "	});	" & vbCrlf

	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Function ShowNavMenu(fType)
	Dim strMenu, rsMenu, sqlMenu, fParentID, fItemName, fOrder, fIcon
	fType = HR_Clng(fType) : If fType = 0 Then fType = 1
	sqlMenu = "Select * From HR_Class Where ClassType=" & fType & " Order By RootID ASC, OrderID ASC"
	Set rsMenu = Conn.Execute(sqlMenu)
		If Not(rsMenu.BOF And rsMenu.EOF) Then
			strMenu = strMenu & vbCrlf
			Do While Not rsMenu.EOF
				fParentID = HR_Clng(rsMenu("ParentID"))
				fItemName = Trim(rsMenu("ClassName"))
				If fParentID > 0 Then
					Set rsTmp = Conn.Execute("Select Max(OrderID) From HR_Class Where ParentID=" & fParentID)
						fOrder = HR_Clng(rsTmp(0))
					Set rsTmp = Nothing
				End If
				fIcon = "<i class=""hr-icon"">&#xf34f;</i>"
				If fParentID > 0 And HR_Clng(rsMenu("OrderID")) <> fOrder Then strMenu = strMenu & "		"
				If fParentID > 0 And HR_Clng(rsMenu("OrderID")) <= fOrder Then fIcon = "<i class=""hr-icon"">&#xf328;</i>"
				If HR_Clng(rsMenu("OrderID")) = fOrder And fParentID > 0 Then strMenu = strMenu & "		"

				strMenu = strMenu & "<div class=""hr-flex_item"" data-id=""" & rsMenu("ClassID") & """>"
				If HR_Clng(rsMenu("Child")) > 0 Then strMenu = strMenu & vbCrlf & "	"
				strMenu = strMenu & "<a "
				If HR_Clng(rsMenu("Child")) > 0 Then
					strMenu = strMenu & " class=""hr-navmenu-main""><em class=""title"">" & fIcon & fItemName & "</em><em class=""more""><i class=""hr-icon"">&#xea44;</i></em></a>"
				Else
					strMenu = strMenu & " class=""hr-navmenu"" href=""" & ParmPath & "Course/List.html?ItemID=" & rsMenu("ClassID") & """><em class=""title"">" & fIcon & fItemName & "</em><em class=""more""><i class=""hr-icon"">&#xef91;</i></em></a>"
				End If
				If HR_Clng(rsMenu("Child")) > 0 Then strMenu = strMenu & vbCrlf & "	<div class=""nav-child"">" & vbCrlf
				If HR_Clng(rsMenu("Child")) = 0 Then strMenu = strMenu & "</div>" & vbCrlf
				If HR_Clng(rsMenu("OrderID")) = fOrder And fParentID > 0 Then
					strMenu = strMenu & "	</div>" & vbCrlf
					strMenu = strMenu & "</div>" & vbCrlf
				End If
				rsMenu.MoveNext
			Loop
		End If
	Set rsMenu = Nothing
	ShowNavMenu = strMenu
End Function

'=====================================================================
'函数名：GetSemesterArr		【取学期/学年数组】
'=====================================================================
Function GetSemesterArr(fType, fVal)
	Dim strFun, iFun, sYear, eYear
	eYear = Year(Date())
	If Month(Date()) > 6 Then eYear = eYear + 1
	sYear = eYear-3
	For iFun = eYear To sYear Step -1
		If iFun > sYear Then
			If HR_Clng(fType) <> 2 Then
				strFun = strFun & "" & iFun-1 & "-" & iFun & "|"
			End If
			strFun = strFun & "" & iFun-1 & "-" & iFun & "-1|"

			strFun = strFun & "" & iFun-1 & "-" & iFun & "-2|"
		End If
	Next
	GetSemesterArr = strFun
End Function
%>