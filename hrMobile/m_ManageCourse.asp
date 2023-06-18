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
Dim scriptCtrl : SiteTitle = "业绩管理"

Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "ListItem" Call ListItem()
	Case "List" Call List()
	Case "Edit" Call Edit()
	Case "SaveEdit" Call SaveEdit()
	Case "Filter" Call FilterForm()
	Case "Details" Call Details()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim Count1, rsNum, numItem		'汇总业绩数
	sql = "" : numItem = 0
	Set rsNum = Conn.Execute("Select ClassID From HR_Class Where ModuleID=1001 And Child=0")
		If Not(rsNum.BOF And rsNum.EOF) Then
			i = 0
			Do While Not rsNum.EOF
				If i > 0 Then sql = sql & " union all "
				sql = sql & "select count(1) as CNT From HR_Sheet_" & rsNum("ClassID") & " Where scYear=" & DefYear
				rsNum.MoveNext
				i = i + 1
			Loop
			numItem = i
		End If
	Set rsNum = Nothing

	sql="select sum(CNT) from (" & sql & ") as nTab"
	Set rsNum = Conn.Execute(sql)
		Count1 = HR_Clng(rsNum(0))
	Set rsNum = Nothing

	Dim arrClassNum(1)
	Set rsNum = Conn.Execute("Select Count(0) From HR_Class Where ModuleID=1001 And ClassType=1")
		arrClassNum(0) = HR_Clng(rsNum(0))
	Set rsNum = Nothing
	Set rsNum = Conn.Execute("Select Count(0) From HR_Class Where ModuleID=1001 And ClassType=2")
		arrClassNum(1) = HR_Clng(rsNum(0))
	Set rsNum = Nothing

	tmpHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	tmpHtml = tmpHtml & "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", tmpHtml)
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-cell"">" & vbCrlf
	Response.Write "	<div class=""weui-cell__bd"">业绩总数</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell__ft"">" & Count1 & " 条</div>"  & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
	Response.Write "<div class=""hr-panel-item hr-fix"">" & vbCrlf
	Response.Write "	<div class=""hr-flex_item"" data-type=""1"">" & vbCrlf
	Response.Write "		<a class=""hr-navmenu"" href=""" & ParmPath & "ManageCourse/ListItem.html?TypeID=1"">" & vbCrlf
	Response.Write "			<em class=""title""><i class=""hr-icon"">&#xe1b2;</i>基础性教学</em>" & vbCrlf
	Response.Write "			<em class=""tips"">共<b>" & arrClassNum(0) & "</b>类</em>" & vbCrlf
	Response.Write "			<em class=""more""><i class=""hr-icon"">&#xf054;</i></em>" & vbCrlf
	Response.Write "		</a>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-flex_item"" data-type=""2"">" & vbCrlf
	Response.Write "		<a class=""hr-navmenu"" href=""" & ParmPath & "ManageCourse/ListItem.html?TypeID=2"">" & vbCrlf
	Response.Write "			<em class=""title""><i class=""hr-icon"">&#xe8a3;</i>激励性教学</em>" & vbCrlf
	Response.Write "			<em class=""tips"">共<b>" & arrClassNum(1) & "</b>类</em>" & vbCrlf
	Response.Write "			<em class=""more""><i class=""hr-icon"">&#xf054;</i></em>" & vbCrlf
	Response.Write "		</a>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Manage/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)

End Sub

Sub ListItem()
	SiteTitle = "基础性教学"
	Dim TypeID : TypeID = HR_Clng(Request("TypeID"))
	If TypeID = 2 Then SiteTitle = "激励性教学"
	Dim rsNum, numRecord, numItem
	Set rsNum = Conn.Execute("Select ClassID From HR_Class Where ModuleID=1001 And Child=0 And ClassType=" & TypeID)
		If Not(rsNum.BOF And rsNum.EOF) Then
			i = 0
			Do While Not rsNum.EOF
				If i > 0 Then sql = sql & " union all "
				sql = sql & "select count(1) as CNT From HR_Sheet_" & rsNum("ClassID") & " Where scYear=" & DefYear
				rsNum.MoveNext
				i = i + 1
			Loop
			numItem = i
		End If
	Set rsNum = Nothing

	sql="select sum(CNT) from (" & sql & ") as nTab"
	Set rsNum = Conn.Execute(sql)
		numRecord = HR_Clng(rsNum(0))
	Set rsNum = Nothing


	tmpHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	tmpHtml = tmpHtml & "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", tmpHtml)
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-cell"">" & vbCrlf
	Response.Write "	<div class=""weui-cell__bd"">" & SiteTitle & " 业绩数：</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell__ft"">" & numRecord & " 条</div>"  & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "<div class=""hr-panel-item hr-fix"">" & vbCrlf
	Response.Write ShowNavMenu(TypeID)
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "ManageCourse/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".hr-navmenu-main"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var tnext = $(this).next("".nav-child"");tnext.toggle();" & vbCrlf
	tmpHtml = tmpHtml & "		var dis = tnext.css(""display"");" & vbCrlf
	tmpHtml = tmpHtml & "		if(dis == ""block""){ $(this).find("".more i"").html(""&#xea45;"");}else{$(this).find("".more i"").html(""&#xea44;"");}" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)

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
					strMenu = strMenu & " class=""hr-navmenu"" href=""" & ParmPath & "ManageCourse/List.html?ItemID=" & rsMenu("ClassID") & """><em class=""title"">" & fIcon & fItemName & "</em><em class=""more""><i class=""hr-icon"">&#xef91;</i></em></a>"
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


Sub List()		'课程业绩列表
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim scYear : scYear = HR_Clng(Request("Year")) : If scYear = 0 Then scYear = DefYear
	ErrMsg = ""

	Dim tItemName, tTemplate, tSheetName, lenField, tFieldHead, tArrHead, tUnit
	Set rsTmp = Conn.Execute("Select Top 1 ClassName,Unit,SheetName,FieldLen, FieldHead, Template From HR_Class Where ClassID=" & tItemID)
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
	SiteTitle = tItemName

	Dim rsList, sqlList, strList, tKSMC
	Dim tVA4, tVA7, tmpTime
	Dim tmpYGDM : tmpYGDM = HR_Clng(Request("SearchWord"))

	sqlList = "Select Top 200 * From " & tSheetName & " Where VA1>0"
	If tmpYGDM > 0 Then sqlList = sqlList & " And VA1=" & tmpYGDM
	If scYear > 2000 Then sqlList = sqlList & " And scYear=" & scYear
	sqlList = sqlList & " Order By AppendTime DESC"
	Set rsList = Server.CreateObject("ADODB.RecordSet")
		rsList.Open(sqlList), Conn, 1, 1
		If Not(rsList.BOF And rsList.EOF) Then
			strList = strList & "<div class=""weui-cell"">" & vbCrlf
			strList = strList & "	<div class=""weui-cell__bd"">" & tItemName & " 业绩数：</div>" & vbCrlf
			strList = strList & "	<div class=""weui-cell__ft"">" & rsList.RecordCount & " 条</div>"  & vbCrlf
			strList = strList & "</div>" & vbCrlf
			'strList = strList & "<div class=""hr-gap-20""></div>" & vbCrlf

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
				strList = strList & "<ul class=""hr-fix listPanel"">" & vbCrlf
				If tTemplate = "TempTableA" Then
					strList = strList & "	<li class=""title"">" & rsList("VA8") & "</li>" & vbCrlf
					strList = strList & "	<li class=""time"">" & tVA4 & "　星期" & rsList("VA6") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>教师姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>科室：</lable><em>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(8) & "：</lable><em>" & rsList("VA8") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(9) & "：</lable><em>" & rsList("VA9") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>校区：</lable><em>" & rsList("VA11") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(3) & "：</lable><em>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
					
					strList = strList & "	<li class=""info""><lable>" & tArrHead(10) & "：</lable><em>" & rsList("VA10") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>周次：</lable><em>" & rsList("VA5") & "周</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>节次：</lable><em>" & rsList("VA7") & "节　" & tmpTime & "</em></li>" & vbCrlf
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
					strList = strList & "	<li class=""time"">" & tVA4 & "</li>" & vbCrlf
					If Trim(rsList("VA8")) <> "" Then strList = strList & "	<li class=""info""><lable>" & tArrHead(8) & "：</lable><em>" & rsList("VA8") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(7) & "：</lable><em>" & rsList("VA7") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(4) & "：</lable><em>" & tVA4 & "　<b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(5) & "：</lable><em>" & rsList("VA5") & "</em></li>" & vbCrlf
				ElseIf tTemplate = "TempTableE" Then
					strList = strList & "	<li class=""title"">" & rsList("VA6") & "</li>" & vbCrlf
					If Trim(rsList("VA9")) <> "" Then strList = strList & "	<li class=""info"">" & rsList("VA9") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(7) & "：</lable><em>" & rsList("VA7") & "　<b>" & tArrHead(8) & "：</b>" & rsList("VA8") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""time""><lable>" & tArrHead(4) & "：</lable><em>" & tVA4 & "　<b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""time""><lable>" & tArrHead(5) & "：</lable><em>" & rsList("VA5") & "</em></li>" & vbCrlf
				ElseIf tTemplate = "TempTableF" Then
					strList = strList & "	<li class=""title"">" & rsList("VA5") & "</li>" & vbCrlf
					If Trim(rsList("VA8")) <> "" Then strList = strList & "	<li class=""info"">" & rsList("VA8") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(6) & "：</lable><em>" & rsList("VA6")
					If Trim(rsList("VA7")) <> "" Then strList = strList & "　<b>" & tArrHead(7) & "：</b>" & rsList("VA7")
					strList = strList & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""time""><lable>" & tArrHead(4) & "：</lable><em>" & tVA4 & "　<b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
				ElseIf tTemplate = "TempTableG" Then
					strList = strList & "	<li class=""title"">" & rsList("VA5") & "</li>" & vbCrlf
					If Trim(rsList("VA7")) <> "" Then strList = strList & "	<li class=""info"">" & rsList("VA7") & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(6) & "：</lable><em>" & rsList("VA6") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]　<b>科室：</b>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""time""><lable>" & tArrHead(4) & "：</lable><em>" & tVA4 & "　<b>" & tArrHead(3) & "：</b>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
				End If
				strList = strList & "	<li class=""info status""><lable>状态：</lable><em>"
				If HR_Clng(rsList("Retreat")) = 1 Then strList = strList & "<span class=""txt_war"">[退回]</span>"
				If HR_Clng(rsList("State")) = 1 Then strList = strList & "<span class=""txt_ok"">[已确认]</span>" Else strList = strList & "<span class=""txt_war"">[未确认]</span>"
				If HR_CBool(rsList("Passed")) Then strList = strList & "<span class=""txt_ok"">[已审]</span>" Else strList = strList & "<span class=""txt_war"">[未审]</span>"
				strList = strList & "</em></li>" & vbCrlf

				strList = strList & "	<li class=""morehref"">" & vbCrlf
				strList = strList & "		<lable><a href=""" & ParmPath & "ManageCourse/Details.html?ItemID=" & tItemID & "&ID=" & rsList("ID") & """ class=""href"">课程详情<i class=""hr-icon hr-icon-top"">&#xf321;</i></a></lable>" & vbCrlf
				strList = strList & "	</li>" & vbCrlf
				strList = strList & "</ul>" & vbCrlf
				rsList.MoveNext
			Loop
		Else
			strList = "<div class=""nolist""><em><i class=""hr-icon"">&#xef61;</i></em><span>没有课程记录</span></div>"
		End If
	Set rsList = Nothing

	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		body {background-color:#f1f1f1;}" & vbCrlf
	strHtml = strHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn {width:55px;height:55px;text-align:center;font-size:2.6rem;position: fixed;right: 1rem;bottom: 3rem;z-index: 99;}" & vbCrlf
	strHtml = strHtml & "		.hr-float-btn i {color:#2196f3;}" & vbCrlf
	strHtml = strHtml & "		.listPanel {border-color:#e5e5e5;padding:10px;border-radius:0;margin:10px 0;}" & vbCrlf
	strHtml = strHtml & "		.listPanel .time {border-bottom:1px solid #f2f2f2;margin-bottom:5px;color:#39b2e4}" & vbCrlf
	strHtml = strHtml & "		.listPanel .info {display:flex;justify-content:space-between;align-items:stretch;}" & vbCrlf
	strHtml = strHtml & "		.listPanel .info lable {width:32%;text-align:right;color:#999} .listPanel .info em {width:70%;flex-grow:1;font-size:0.9rem;}" & vbCrlf
	strHtml = strHtml & "		.listPanel .morehref {border-top:1px solid #eee;margin-top:5px;text-align:right;} .morehref a {color:#39b2e4}" & vbCrlf
	strHtml = strHtml & "		.txt_war {color:#f30;} .txt_ok {color:#3a3;}" & vbCrlf
	strHtml = strHtml & "		.nolist {text-align:center;font-size:1.2rem;} .nolist em {color:#f30;font-size:4rem;}" & vbCrlf
	
	strHtml = strHtml & "		.hr-float-btn {border-radius:5px;width:45px;height:45px;line-height:45px;background-color:#2196f3;font-size:1.5rem;text-align:center} .hr-float-btn i{color:#fff;}" & vbCrlf
	strHtml = strHtml & "		.filter {bottom:6rem;}" & vbCrlf
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
	Response.Write "<div class=""hr-float-btn""><a href=""" & ParmPath & "ManageCourse/Edit.html?AddNew=True&ItemID=" & tItemID & """ class=""addBtn""><i class=""hr-icon"">&#xf067;</i></a></div>" & vbCrlf
	Response.Write "<div class=""hr-float-btn filter""><a href=""" & ParmPath & "ManageCourse/Filter.html?AddNew=True&ItemID=" & tItemID & """ class=""addBtn""><i class=""hr-icon"">&#xefc3;</i></a></div>" & vbCrlf
	Response.Write "<div class=""hr-sides-x10 hr-fix"">筛选：</div>" & vbCrlf
	Response.Write "<div class=""hr-sides-x10 hr-fix"">" & vbCrlf
	Response.Write " " & strList
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".hr-navmenu-main"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var tnext = $(this).next("".nav-child"");tnext.toggle();" & vbCrlf
	strHtml = strHtml & "		var dis = tnext.css(""display"");" & vbCrlf
	strHtml = strHtml & "		if(dis == ""block""){ $(this).find("".more i"").html(""&#xea45;"");}else{$(this).find("".more i"").html(""&#xea44;"");}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".addNew a"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		location.href=""" & ParmPath & "Course/Edit.html?AddNew=True&ItemID=" & tItemID & """;" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

Sub FilterForm()
	SiteTitle = "筛选条件"
	Dim ItemType : ItemType = "基础性教学"
	Dim TypeID, ItemID : ItemID = HR_Clng(Request("ItemID"))
	Dim tItemName, tTemplate, tSheetName
	Set rsTmp = Conn.Execute("Select Top 1 ClassName,ClassType,Unit,SheetName,FieldLen,FieldHead,Template From HR_Class Where ClassID=" & ItemID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			tItemName = Trim(rsTmp("ClassName"))
			tTemplate = Trim(rsTmp("Template"))
			TypeID = Trim(rsTmp("ClassType"))
			tSheetName = "HR_Sheet_" & ItemID
			If Not(ChkTable(tSheetName)) Then ErrMsg = "未找到数据表 " & tSheetName & "！<br>"
		Else
			ErrMsg = "业绩考核项目不存在！<br>"
		End If
	Set rsTmp = Nothing
	If TypeID = 2 Then ItemType = "激励性教学"

	If ErrMsg<>"" Then Response.Write GetErrBody(0) : Exit Sub

	tmpHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	tmpHtml = tmpHtml & "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cells {font-size:1rem;} .weui-cells_form .weui-cell__ft {font-size:1.1rem}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-cell {padding:8px 10px;} .weui-cell .weui-label {text-align:right;color:#999;width:115px;}" & vbCrlf

	tmpHtml = tmpHtml & "		.select1 {color:#39b2e4;} " & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", tmpHtml)
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "		" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label for="""" class=""weui-label"">教师：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input class=""weui-input"" type=""text"" name=""YGXM"" id=""ygxm"" value="""" data-key=""ygxm"" data-value=""ygdm"" placeholder=""请选择教师"" readonly></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft select1"" data-id=""ygxm""><i class=""hr-icon"">&#xeeed;</i>选择</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label for="""" class=""weui-label"">教师工号：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input class=""weui-input"" type=""text"" name=""YGDM"" id=""ygdm"" value="""" readonly></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label for="""" class=""weui-label"">选择部门：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><input class=""weui-input"" type=""text"" name=""KSMC"" id=""ksmc"" value="""" data-key=""ksmc"" data-value=""ksdm"" placeholder=""点击选择"" readonly></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft select2"" data-id=""ksmc""><i class=""hr-icon"">&#xf0e8;</i>选择</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<input class=""weui-input"" type=""hidden"" name=""KSDM"" id=""ksdm"" value="""">" & vbCrlf
	Response.Write "	<div class=""weui-cell""><button type=""button"" name=""searchPost"" class=""weui-btn weui-btn_primary"" id=""searchPost"">搜索</button>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div id=""full"" class=""hr-popup"">" & vbCrlf
	Response.Write "	<iframe src=""about:bank"" name=""listFrame"" id=""listFrame"" title=""ListFrame"" width=""100%"" height=""100%"" frameborder=""0""></iframe>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "ManageCourse/Index.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".select1"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#full"").show(); var obj=$(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#listFrame"").attr(""src"",""" & ParmPath & "Directories/SelectTeacher.html?Type=3&reObjTxt="" + $(""#""+obj).data(""key"") + ""&reObjValue="" +  $(""#""+obj).data(""value""));" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#ygdm"").bind(""oninput"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			alert(""111"");" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	$("".select2"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#full"").show(); var obj=$(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#listFrame"").attr(""src"",""" & ParmPath & "ManageDepart/SelectDepart.html?reObjTxt="" + $(""#""+obj).data(""key"") + ""&reObjValue="" +  $(""#""+obj).data(""value""));" & vbCrlf
	tmpHtml = tmpHtml & "		$(""#ygdm"").bind(""oninput"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			alert(""111"");" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	
	tmpHtml = tmpHtml & "	function show(){" & vbCrlf
	tmpHtml = tmpHtml & "		alert(""CCCC"");" & vbCrlf
	tmpHtml = tmpHtml & "	}" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

Sub Details()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tSheetName : tSheetName = "HR_Sheet_" & tItemID
	Dim tmpView, tVA2, tVA4, tCourse, tItem, tCourseDate, tPeriod, tPeriodTime, tPlace, tStuClass, NotModi
	Dim tFieldHead, tArrHead
	NotModi = False : ErrMsg = "" : SiteTitle = "课程详情"
	If Not(ChkTable(tSheetName)) Then ErrMsg = "未找到数据表 " & tSheetName & "！<br>"

	If ChkTable(tSheetName) Then
		sqlTmp = "Select a.*,(Select ClassName From HR_Class Where ClassID=a.ItemID) As ItemName, (Select Template From HR_Class Where ClassID=a.ItemID) As Template"
		sqlTmp = sqlTmp & ",(Select FieldLen From HR_Class Where ClassID=a.ItemID) As FieldLen,(Select FieldHead From HR_Class Where ClassID=a.ItemID) As FieldHead"
		sqlTmp = sqlTmp & " From " & tSheetName & " a Where a.ID=" & tmpID
		Set rsTmp = Conn.Execute(sqlTmp)
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				If HR_IsNull(rsTmp("FieldHead")) = False Then
					tFieldHead = FilterArrNull(rsTmp("FieldHead"), ",")
					tArrHead = Split(tFieldHead, ",")
					If Ubound(tArrHead) <> HR_Clng(rsTmp("FieldLen")) Then Redim Preserve tArrHead(HR_Clng(rsTmp("FieldLen")))
				Else
					Redim tArrHead(HR_Clng(rsTmp("FieldLen")))
				End If


				tmpView = tmpView & "		<dl class=""hr-rows""><dt>所属项目：</dt><dd>" & Trim(rsTmp("ItemName")) & "</dd></dl>" & vbCrlf
				tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(2) & "：</dt><dd>" & Trim(rsTmp("VA2")) & " [" & Trim(rsTmp("VA1")) & "]</dd></dl>" & vbCrlf
				tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(3) & "：</dt><dd>" & Trim(rsTmp("VA3")) & "学时</dd></dl>" & vbCrlf
				If rsTmp("Template") = "TempTableA" Or rsTmp("Template") = "TempTableC" Or rsTmp("Template") = "TempTableD" Or rsTmp("Template") = "TempTableE" Then
					tCourseDate = FormatDate(ConvertNumDate(rsTmp("VA4")), 4)
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(4) & "：</dt><dd>" & tCourseDate & "</dd></dl>" & vbCrlf
				Else
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(4) & "：</dt><dd>" & Trim(rsTmp("VA4")) & "</dd></dl>" & vbCrlf
				End If

				If rsTmp("Template") = "TempTableA" Then
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>　</dt><dd>周" & Trim(rsTmp("VA6")) & " 第" & Trim(rsTmp("VA7")) & "节 " & GetPeriodTime(rsTmp("VA11"), rsTmp("VA7"), 1) & "</dd></dl>" & vbCrlf
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>周　次：</dt><dd>" & Trim(rsTmp("VA5")) & "周</dd></dl>" & vbCrlf
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(8) & "：</dt><dd>" & Trim(rsTmp("VA8")) & "</dd></dl>" & vbCrlf
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>授课内容：</dt><dd>" & Trim(rsTmp("VA9")) & "</dd></dl>" & vbCrlf
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>授课对象：</dt><dd>" & Trim(rsTmp("VA10")) & "</dd></dl>" & vbCrlf
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(12) & "：</dt><dd>" & Trim(rsTmp("VA11")) & " " & Trim(rsTmp("VA12")) & "</dd></dl>" & vbCrlf
				ElseIf rsTmp("Template") = "TempTableB" Then
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(5) & "：</dt><dd>" & Trim(rsTmp("VA5")) & "</dd></dl>" & vbCrlf
					If HR_IsNull(rsTmp("VA6")) = False Then tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(6) & "：</dt><dd>" & Trim(rsTmp("VA6")) & "</dd></dl>" & vbCrlf
				ElseIf rsTmp("Template") = "TempTableC" Or rsTmp("Template") = "TempTableD" Or rsTmp("Template") = "TempTableG" Then
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(5) & "：</dt><dd>" & Trim(rsTmp("VA5")) & "</dd></dl>" & vbCrlf
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(6) & "：</dt><dd>" & Trim(rsTmp("VA6")) & "</dd></dl>" & vbCrlf
					If rsTmp("Template") = "TempTableC" Or rsTmp("Template") = "TempTableD" Or rsTmp("Template") = "TempTableG" Then
						If HR_IsNull(rsTmp("VA7")) = False Then tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(7) & "：</dt><dd>" & Trim(rsTmp("VA7")) & "</dd></dl>" & vbCrlf
						If rsTmp("Template") = "TempTableD" Then tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(8) & "：</dt><dd>" & Trim(rsTmp("VA8")) & "</dd></dl>" & vbCrlf
					End If
				ElseIf rsTmp("Template") = "TempTableE" Or rsTmp("Template") = "TempTableF" Then
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(5) & "：</dt><dd>" & Trim(rsTmp("VA5")) & "</dd></dl>" & vbCrlf
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(6) & "：</dt><dd>" & Trim(rsTmp("VA6")) & "</dd></dl>" & vbCrlf
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(7) & "：</dt><dd>" & Trim(rsTmp("VA7")) & "</dd></dl>" & vbCrlf
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(8) & "：</dt><dd>" & Trim(rsTmp("VA8")) & "</dd></dl>" & vbCrlf
					If rsTmp("Template") = "TempTableE" Then tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(9) & "：</dt><dd>" & Trim(rsTmp("VA9")) & "</dd></dl>" & vbCrlf
				Else
					tmpView = tmpView & "		<dl class=""hr-rows""><dt>" & tArrHead(5) & "：</dt><dd>" & Trim(rsTmp("VA5")) & "</dd></dl>" & vbCrlf
				End If

				tmpView = tmpView & "		<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
				tmpView = tmpView & "		<dl class=""hr-rows""><dt>状态：</dt><dd>"
				If HR_Clng(rsTmp("Retreat")) = 1 Then tmpView = tmpView & "<span class=""txt_war"">[退回]</span>"
				If HR_Clng(rsTmp("State")) = 1 Then tmpView = tmpView & "<span class=""txt_ok"">[已确认]</span>" Else tmpView = tmpView & "<span class=""txt_war"">[未确认]</span>"
				If HR_CBool(rsTmp("Passed")) Then tmpView = tmpView & "<span class=""txt_ok"">[已审]</span>" Else tmpView = tmpView & "<span class=""txt_war"">[未审]</span>"
				tmpView = tmpView & "</dd></dl>" & vbCrlf
			End If
		Set rsTmp = Nothing
	End If

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dl {align-items:stretch;border-bottom:1px solid #eee;padding:5px 0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dt {width:30%;text-align:right;color:#999}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dd {flex-grow:2;width:70%;box-sizing: border-box;padding-right:3px}" & vbCrlf
	tmpHtml = tmpHtml & "		.editbtn {width:50%;padding:10px 5px;} .editbtn em {width:50%;padding:10px 5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-dialog__hd {padding:2px 0 5px 0;border-bottom:1px solid #ccc;}" & vbCrlf
	tmpHtml = tmpHtml & "		.weui-dialog__bd {padding:0 0 15px 0;} .weui-prompt-input {width:95%;height:3rem;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-swap-box"">" & vbCrlf
	Response.Write "	<div class=""hr-swap-items"">" & vbCrlf
	Response.Write "		" & tmpView
	Response.Write "	</div>" & vbCrlf
	
	If NotModi = False Then
		Response.Write "	<div class=""hr-shrink-x20""></div>" & vbCrlf
		Response.Write "	<div class=""hr-rows hr-editbtn"">" & vbCrlf
		Response.Write "		<em><i class=""hr-icon"">&#xea3f;</i></em>" & vbCrlf
		Response.Write "		<em><button type=""button"" name=""pass"" class=""pass"" id=""Pass"" data-id=""" & tmpID & """>审核</button></em>" & vbCrlf
		Response.Write "		<em><button type=""button"" name=""edit"" class=""edit"" id=""Edit"" data-id=""" & tmpID & """>修改</button></em>" & vbCrlf
		Response.Write "		<em><button type=""button"" name=""delete"" class=""delete"" id=""Delete"" data-id=""" & tmpID & """>删除</button></em>" & vbCrlf
		Response.Write "		<em><button type=""button"" name=""retreat"" class=""retreat"" id=""Retreat"" data-id=""" & tmpID & """>退回</button></em>" & vbCrlf
		Response.Write "	</div>" & vbCrlf
	End If
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#Edit"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var cid=$(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		location.href=""" & ParmPath & "ManageCourse/Edit.html?Modify=True&ItemID=" & tItemID & "&ID=""+ cid;" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#Delete"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var cid=$(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		$.confirm(""您确定要删除该条课程？"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "			$.getJSON(""" & ParmPath & "ManageCourse/Delete.html"",{ItemID:" & tItemID & ",ID:cid},function(reData){" & vbCrlf
	tmpHtml = tmpHtml & "				$.toast(reData.reMessge, function(){location.href=""" & ParmPath & "ManageCourse//List.html?ItemID=" & tItemID & """; });" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#Pass"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$.toast(""审核课程发生未知错误！"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#Retreat"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		$.prompt({title: '退回课程',input: '请输入退回理由'});" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub


'编辑
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


	Dim  strVA4, tVA4 : tVA4 = Trim(arrField(4))
	If tTemplate = "TempTableA" Or tTemplate = "TempTableC" Or tTemplate = "TempTableD" Or tTemplate = "TempTableE" Then
		tVA4 = FormatDate(ConvertNumDate(tVA4), 2)			'转为日期
		strVA4 = "date"
	Else
		strVA4 = "text"
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
			strList = strList & "			<input name=""VA4"" class=""weui-input"" id=""VA4"" type=""" & strVA4 & """ value=""" & tVA4 & """ placeholder=""" & tArrHead(4) & """>" & vbCrlf
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
			strList = strList & "	<input name=""ItemID"" id=""ItemID"" type=""hidden"" value=""" & tItemID & """><input name=""ID"" id=""ID"" type=""hidden"" value=""" & tmpID & """>" & vbCrlf
			If IsModify Then strList = strList & "	<input name=""Modify"" id=""Modify"" type=""hidden"" value=""True"">" & vbCrlf
			strList = strList & "	<div class=""weui-cell""><button type=""button"" name=""editPost"" class=""weui-btn weui-btn_primary"" id=""editPost"">保存</button>" & vbCrlf
			strList = strList & "</form>" & vbCrlf
			strList = strList & "</div>" & vbCrlf
			strList = strList & "" & vbCrlf
			strList = strList & "" & vbCrlf


	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-header {background-color:#406;}" & vbCrlf
	strHtml = strHtml & "		.itemTit{text-align:center;font-size:1.2rem;border-bottom:1px solid #ccc;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells {font-size:0.9rem;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .popWin, .weui-cells_form .Unit {font-size:0.9rem;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .popWin i {font-size:1.1rem;position:relative;top:2px;color:#4ce}" & vbCrlf
	strHtml = strHtml & "		.weui-label {text-align:right;}" & vbCrlf
	strHtml = strHtml & "		.btnBar {box-sizing:border-box;padding:10px 0;}" & vbCrlf
	strHtml = strHtml & "		#EditForm{width:100%;height:100%;margin:0;padding:0;border:0;}" & vbCrlf
	strHtml = strHtml & "		#listFrame {border:0;}" & vbCrlf
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
	Response.Write "<div class=""hr-sides-x10 hr-fix"">" & vbCrlf
	Response.Write "	<div class=""itemTit"">" & tItemName & "</div>" & vbCrlf
	Response.Write "	" & strList
	Response.Write "</div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	var stuType = (""" & XmlText("Common", "StudentType", "") & """).split(""|"");" & vbCrlf
	strHtml = strHtml & "	$(""#StuType"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择学生类别"",items:stuType" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	If tTemplate = "TempTableA" Then
		strHtml = strHtml & "	$(""#VA8"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择"",items:[" & GetCourseSelect("VA8", "") & "]" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
		strHtml = strHtml & "	$(""#VA10"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择"",items:[" & getFieldSelect(tItemID, "VA10", "") & "]" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
		strHtml = strHtml & "	$(""#VA11"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择"",items:[" & GetCampusSelect("VA11", "") & "]" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
		strHtml = strHtml & "	$(""#VA4"").on(""change"",function(){" & vbCrlf
		strHtml = strHtml & "		var today = new Array('日','一','二','三','四','五','六'), day = new Date($(this).val());" & vbCrlf
		strHtml = strHtml & "		var week = today[day.getDay()];$(""#VA6"").val(week);" & vbCrlf
		strHtml = strHtml & "		console.log($(this).val());" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
	ElseIf tTemplate = "TempTableD" Then
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
		strHtml = strHtml & "		title: ""请选择等级"",items:[" & GetItemGradeSelect(tItemID, arrField(7), arrField(8)) & "]" & vbCrlf
		strHtml = strHtml & "	});" & vbCrlf
	ElseIf tTemplate = "TempTableF" Or tTemplate = "TempTableG" Then
		strHtml = strHtml & "	$(""#VA6"").select({" & vbCrlf
		strHtml = strHtml & "		title: ""请选择级别"",items:[" & GetSubmoduleSelect(tItemID, "") & "]" & vbCrlf
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
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "ManageCourse/List.html?ItemID=" & tItemID & """; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$(""#editPost"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "		$.getJSON(""" & ParmPath & "ManageCourse/SaveEdit.html"",$(""#EditForm"").serialize(), function(strForm){" & vbCrlf
	strHtml = strHtml & "			$.alert(strForm.reMessge,function(){" & vbCrlf
	strHtml = strHtml & "				if(strForm.Return){location.href=""" & ParmPath & "ManageCourse/List.html?ItemID=" & tItemID & "&ID=""+ strForm.id}" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub SaveEdit()
	Response.Write "{""Return"":false,""reMessge"":""未知错误，保存失败！""}"
End Sub
%>