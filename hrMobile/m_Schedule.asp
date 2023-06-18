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
SiteTitle = "我的日程"

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "View" Call View()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	SiteTitle = "授课日程"
	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<link type=""text/css"" href=""" & InstallDir & "Static/swiper/css/swiper.min.css?v=4.2.6"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "	.hr-header-slide {position: relative;width:100%;height:50px;line-height:50px;background-color:#eee;overflow: hidden;}" & vbCrlf
	strHtml = strHtml & "	.hr-header-slide .swiper-slide {width:auto; padding:0 10px;}" & vbCrlf
	strHtml = strHtml & "	.list-Course .weui-cell__hd em {width:40px;height:40px;line-height:40px;text-align:center;border:1px solid #eee;margin-right:10px;border-radius:25px;background-color:#eee;color:#777;font-size:1.2rem}" & vbCrlf
	strHtml = strHtml & "	.list-Course .weui-cell__hd em.icon1 {background-color:#f80;color:#fff}" & vbCrlf
	strHtml = strHtml & "	.list-Course a h3, .list-Course a h4, .list-Course a h5 {font-weight:normal;font-size:14px;color:#777;}" & vbCrlf
	strHtml = strHtml & "	.list-Course a h3 {font-size:15px;color:#000;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/swiper/js/swiper.min.js?v=4.2.6"">" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Dim strQuery, tItemName, tTmpTable, tTime, tCampus, tPeriod, icon1, tItemIcon, tCourse
	sqlTmp = GetRemainCourse()		'未上课
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open sqlTmp, Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			i = 0
			Do While Not rsTmp.EOF
				tItemName = GetTypeName("HR_Class", "ClassName", "ClassID", rsTmp("ItemID"))
				tTmpTable = GetTypeName("HR_Class", "Template", "ClassID", rsTmp("ItemID"))
				tItemIcon = GetTypeName("HR_Class", "ItemIcon", "ClassID", rsTmp("ItemID"))
				tCampus = "" : tPeriod = "" : tTime = "　" : tItemIcon = "xef2d" : icon1 = 0
				If Trim(tTmpTable) = "TempTableA" Then
					tPeriod = GetTypeName("HR_Sheet_" & rsTmp("ItemID"), "VA7", "ID", rsTmp("ID"))
					tCampus = GetTypeName("HR_Sheet_" & rsTmp("ItemID"), "VA11", "ID", rsTmp("ID"))
					tCourse = GetTypeName("HR_Sheet_" & rsTmp("ItemID"), "VA8", "ID", rsTmp("ID"))
				Else
					tCourse = GetTypeName("HR_Sheet_" & rsTmp("ItemID"), "VA5", "ID", rsTmp("ID"))
				End If
				If HR_IsNull(tPeriod) = False Then
					tTime = GetPeriodTime(tCampus, tPeriod, 0)
				End If
				If i < 4 Then icon1 = " icon1"
				strQuery = strQuery & "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "Course/View.html?ItemID=" & rsTmp("ItemID") & "&ID=" & rsTmp("ID") & """>" & vbCrlf
				strQuery = strQuery & "		<div class=""weui-cell__hd""><em class=""list-icon" & icon1 & """><i class=""hr-icon"">&#" & tItemIcon & ";</i></em></div>" & vbCrlf
				strQuery = strQuery & "		<div class=""weui-cell__bd""><h3>课程：" & tCourse & "</h3><h4 data-ygdm=""" & rsTmp("VA2") & """>项目：" & tItemName & "</h4><h5>节次：<b>" & tPeriod & "</b></h5></div>" & vbCrlf
				strQuery = strQuery & "		<div class=""weui-cell__ft list-date""><h4>" & FormatDate(ConvertNumDate(rsTmp("VA4")), 2) & "</h4><h5>" & tTime & "</h5></div>" & vbCrlf
				strQuery = strQuery & "	</a>" & vbCrlf
				rsTmp.MoveNext
				i = i + 1
			Loop
		Else
			Response.Write "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef5f;</i></h2><h3>" & DefYear & "学年您暂时没有课程！</h3></div>" & vbCrlf
		End If
	Set rsTmp = Nothing

	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells list-Course"">" & vbCrlf
	Response.Write "	" & strQuery & vbCrlf
	Response.Write "</div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	'strHtml = strHtml & "	var navSwiper = new Swiper("".hr-header-slide"", {" & vbCrlf
	'strHtml = strHtml & "	    freeMode: true, slidesPerView:""auto"", freeModeSticky: true," & vbCrlf
	'strHtml = strHtml & "	 });" & vbCrlf

	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub View()
	SiteTitle = "授课日程"
	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<link type=""text/css"" href=""" & InstallDir & "Static/swiper/css/swiper.min.css?v=4.2.6"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "	.hr-header-slide {position: relative;width:100%;height:50px;line-height:50px;background-color:#eee;overflow: hidden;}" & vbCrlf
	strHtml = strHtml & "	.hr-header-slide .swiper-slide {width:auto; padding:0 10px;}" & vbCrlf
	strHtml = strHtml & "	.list-Course .weui-cell__hd em {width:50px;height:50px;line-height:50px;text-align:center;border:1px solid #ccc;margin-right:10px;border-radius:25px;}" & vbCrlf
	strHtml = strHtml & "	.list-Course a h3, .list-Course a h4, .list-Course a h5 {font-weight:normal;font-size:14px;color:#777;}" & vbCrlf
	strHtml = strHtml & "	.list-Course a h3 {font-size:15px;color:#000;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/swiper/js/swiper.min.js?v=4.2.6"">" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Schedule/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".addNew a"").css(""display"",""none"");" & vbCrlf
	strHtml = strHtml & "	$("".viewMSG"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		$.get(""" & ParmPath & "Notice/View.html"",{id:$(this).data(""id"")}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "			$.alert(rsStr, ""查看通知"");" & vbCrlf
	strHtml = strHtml & "			$("".ShowCourse"").css(""display"",""none"");" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Function GetRemainCourse()				'取未上课程
	Dim iFun, funItem, arrItem, strFun : strFun = ""
	funItem = GetItemClassID(" And Template in('TempTableA','TempTableC') ")		'取考核项ID
	If HR_IsNull(funItem) = False Then
		arrItem = Split(FilterArrNull(funItem, ","), ",")
		For iFun = 0 To Ubound(arrItem)
			If iFun > 0 Then strFun = strFun & " union all "
			strFun = strFun & "(Select ID,ItemID,VA1,VA2,VA4 From HR_Sheet_" & arrItem(iFun) & " Where VA4>CAST(GetDate() As SMALLDATETIME) And VA1=" & HR_Clng(UserYGDM) & " And scYear=" & DefYear
			strFun = strFun & ")"
		Next
	End If
	GetRemainCourse = strFun
End Function
%>