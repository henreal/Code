<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./m_ManageCommon.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "已退回课程"

Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "List" Call List()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	strHtml = "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		body {background-color:#f2f2f2;} .hr-header {background-color:#061;}" & vbCrlf
	strHtml = strHtml & "		.hr-panel-item {background-color:#fff;}" & vbCrlf
	strHtml = strHtml & "		.hr-flex_item a .title {flex-grow:2;}" & vbCrlf
	strHtml = strHtml & "		.hr-panel-title {background-color:#fff;text-align:center;border-bottom:1px solid #eee;font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf

	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ FastClick.attach(document.body); });" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim TotalNum : sql = ""
	Set rs = Conn.Execute("Select ClassID From HR_Class Where ModuleID=1001 And Child=0")
		If Not(rs.BOF And rs.EOF) Then
			i = 0
			Do While Not rs.EOF
				If i > 0 Then sql = sql & " union all "
				sql = sql & "select count(1) as CNT From HR_Sheet_" & rs("ClassID") & " Where Retreat=1 And scYear=" & DefYear
				rs.MoveNext
				i = i + 1
			Loop
		End If
	Set rs = Nothing
	sql="select sum(CNT) from (" & sql & ") as nTab"
	Set rs = Conn.Execute(sql)
		TotalNum = HR_Clng(rs(0))
	Set rs = Nothing

	Response.Write "<div class=""weui-form-preview"">" & vbCrlf
	Response.Write "	<div class=""weui-form-preview__hd"">" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item""><label class=""weui-form-preview__label"">已退回课程数</label><em class=""weui-form-preview__value"">" & HR_Clng(TotalNum) & "</em></div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
	Response.Write "<div class=""hr-panel-title"">基础性教学</div>" & vbCrlf
	Response.Write "<div class=""hr-panel-item hr-fix"">" & vbCrlf
	Response.Write ListGobackCourseItem(1)
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
	Response.Write "<div class=""hr-panel-title"">激励性教学</div>" & vbCrlf
	Response.Write "<div class=""hr-panel-item hr-fix"">" & vbCrlf
	Response.Write ListGobackCourseItem(2)
	Response.Write "</div>" & vbCrlf


	Response.Write "<div class=""hr-shrink-x10""></div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Manage/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".hr-navmenu-main"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var tnext = $(this).next("".nav-child"");tnext.toggle();" & vbCrlf
	strHtml = strHtml & "		var dis = tnext.css(""display"");" & vbCrlf
	strHtml = strHtml & "		if(dis == ""block""){ $(this).find("".more i"").html(""&#xea45;"");}else{$(this).find("".more i"").html(""&#xea44;"");}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

Function ListGobackCourseItem(fType)		'列表项目
	Dim strMenu, rsMenu, sqlMenu, fParentID, fItemName, fOrder, fIcon, rsCount, iRetreat
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
					Set rsCount = Conn.Execute("Select count(1) as CNT From HR_Sheet_" & rsMenu("ClassID") & " Where Retreat=1 And scYear=" & DefYear)
						iRetreat = HR_Clng(rsCount(0))
					Set rsCount = Nothing
					strMenu = strMenu & " class=""hr-navmenu"" href=""" & ParmPath & "ManageGoback/List.html?ItemID=" & rsMenu("ClassID") & """><em class=""title"">" & fIcon & fItemName & "</em><em>" & iRetreat & "</em><em class=""more""><i class=""hr-icon"">&#xef91;</i></em></a>"
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
	ListGobackCourseItem = strMenu
End Function

Sub List()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
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
	SiteTitle = "退回课程"

	Dim TotalNum, sqlList, rsList, strList, tVA4, tVA7, tmpTime, tKSMC
	Set rs = Conn.Execute("Select Count(1) From " & tSheetName & " Where Retreat=1 And scYear=" & DefYear)
		TotalNum = HR_Clng(rs(0))
	Set rs = Nothing

	sqlList = "Select Top 300 a.*,(Select KSMC From HR_Teacher Where YGDM=a.VA1) As KSMC1 From " & tSheetName & " a Where Retreat=1 And scYear=" & DefYear
	sqlList = sqlList & " Order By VA4 DESC"
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
				tKSMC = Trim(rsList("KSMC1"))
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
					strList = strList & "	<li class=""time"">" & tArrHead(4) & " " & tVA4 & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>教师姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>科室：</lable><em>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(3) & "：</lable><em>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(5) & "：</lable><em>" & rsList("VA5") & "</em></li>" & vbCrlf
				ElseIf tTemplate = "TempTableC" Then
					strList = strList & "	<li class=""title"">" & rsList("VA6") & "</li>" & vbCrlf
					strList = strList & "	<li class=""time"">" & tArrHead(4) & " " & tVA4 & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>教师姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>科室：</lable><em>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(3) & "：</lable><em>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(5) & "：</lable><em>" & rsList("VA5") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(7) & "：</lable><em>" & rsList("VA7") & "</em></li>" & vbCrlf
				ElseIf tTemplate = "TempTableD" Then
					strList = strList & "	<li class=""title"">" & tItemName & "</li>" & vbCrlf
					strList = strList & "	<li class=""time"">" & tArrHead(4) & " " & tVA4 & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>教师姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>科室：</lable><em>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(3) & "：</lable><em>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(5) & "：</lable><em>" & rsList("VA5") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(6) & "：</lable><em>" & rsList("VA6") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(7) & "：</lable><em>" & rsList("VA7") & "</em></li>" & vbCrlf
					If Trim(rsList("VA8")) <> "" Then strList = strList & "	<li class=""info""><lable>" & tArrHead(8) & "：</lable><em>" & rsList("VA8") & "</em></li>" & vbCrlf
				ElseIf tTemplate = "TempTableE" Then
					strList = strList & "	<li class=""title"">" & tItemName & "</li>" & vbCrlf
					strList = strList & "	<li class=""time"">" & tArrHead(4) & " " & tVA4 & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>教师姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>科室：</lable><em>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(3) & "：</lable><em>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(5) & "：</lable><em>" & rsList("VA5") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(6) & "：</lable><em>" & rsList("VA6") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(7) & "：</lable><em>" & rsList("VA7") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(8) & "：</lable><em>" & rsList("VA8") & "</em></li>" & vbCrlf
					If Trim(rsList("VA9")) <> "" Then strList = strList & "	<li class=""info""><lable>" & tArrHead(9) & "：</lable><em>" & rsList("VA9") & "</em></li>" & vbCrlf
				ElseIf tTemplate = "TempTableF" Then
					strList = strList & "	<li class=""title"">" & tItemName & "</li>" & vbCrlf
					strList = strList & "	<li class=""time"">" & tArrHead(4) & " " & tVA4 & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>教师姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>科室：</lable><em>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(3) & "：</lable><em>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>教师姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>科室：</lable><em>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(5) & "：</lable><em>" & rsList("VA5") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(6) & "：</lable><em>" & rsList("VA6") & "</em></li>" & vbCrlf
					If Trim(rsList("VA7")) <> "" Then strList = strList & "	<li class=""info""><lable>" & tArrHead(7) & "：</lable><em>" & rsList("VA7") & "</em></li>" & vbCrlf
					If Trim(rsList("VA8")) <> "" Then strList = strList & "	<li class=""info""><lable>" & tArrHead(8) & "：</lable><em>" & rsList("VA8") & "</em></li>" & vbCrlf
				ElseIf tTemplate = "TempTableG" Then
					strList = strList & "	<li class=""title"">" & tItemName & "</li>" & vbCrlf
					strList = strList & "	<li class=""time"">" & tArrHead(4) & " " & tVA4 & "</li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>教师姓名：</lable><em>" & rsList("VA2") & " [" & rsList("VA1") & "]</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>科室：</lable><em>" & tKSMC & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(3) & "：</lable><em>" & rsList("VA3") & " " & tUnit & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(5) & "：</lable><em>" & rsList("VA5") & "</em></li>" & vbCrlf
					strList = strList & "	<li class=""info""><lable>" & tArrHead(6) & "：</lable><em>" & rsList("VA6") & "</em></li>" & vbCrlf
					If Trim(rsList("VA7")) <> "" Then strList = strList & "	<li class=""info""><lable>" & tArrHead(6) & "：</lable><em>" & rsList("VA7") & "</em></li>" & vbCrlf
				End If

				strList = strList & "	<li class=""info status""><lable>状态：</lable><em>"
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
			strList = "<div class=""nolist""><em><i class=""hr-icon"">&#xef61;</i></em><span>没有退回课程</span></div>"
		End If
	Set rsList = Nothing

	strHtml = "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		body {background-color:#f2f2f2;} .hr-header {background-color:#061;}" & vbCrlf
	strHtml = strHtml & "		.listPanel {border-color:#e5e5e5;padding:10px;border-radius:0;margin:10px 0;}" & vbCrlf
	strHtml = strHtml & "		.listPanel .time {border-bottom:1px solid #f2f2f2;margin-bottom:5px;color:#39b2e4}" & vbCrlf
	strHtml = strHtml & "		.listPanel .info {display:flex;justify-content:space-between;align-items:stretch;}" & vbCrlf
	strHtml = strHtml & "		.listPanel .info lable {width:32%;text-align:right;color:#999} .listPanel .info em {width:70%;flex-grow:1;}" & vbCrlf
	strHtml = strHtml & "		.listPanel .morehref {border-top:1px solid #eee;margin-top:5px;text-align:right;} .morehref a {color:#39b2e4}" & vbCrlf
	strHtml = strHtml & "		.txt_war {color:#f30;} .txt_ok {color:#3a3;}" & vbCrlf
	strHtml = strHtml & "		.nolist {text-align:center;font-size:1.2rem;} .nolist em {color:#f30;font-size:4rem;}" & vbCrlf
	strHtml = strHtml & "		.hr-panel-title {color:#3a3;text-align:center;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf

	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)
	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ FastClick.attach(document.body); });" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<div class=""hr-panel-title"">" & tItemName & "</div>" & vbCrlf
	Response.Write "<div class=""weui-form-preview"">" & vbCrlf
	Response.Write "	<div class=""weui-form-preview__hd"">" & vbCrlf
	Response.Write "		<div class=""weui-form-preview__item""><label class=""weui-form-preview__label"">已退回数</label><em class=""weui-form-preview__value"">" & HR_Clng(TotalNum) & "</em></div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	'Response.Write "<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf

	Response.Write "<div class=""hr-sides-x10 hr-fix"">" & strList & "</div>" & vbCrlf

	Response.Write "<div class=""hr-shrink-x10""></div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "ManageGoback/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub
%>