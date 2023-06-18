<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
SiteTitle = "选择听课课程"
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "getItemCourse" Call getItemCourse()	'返回筛选结果
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tItemName, tItemID, tStartDate, tEndDate
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "	<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .weui-cell__ft {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.hr-pop-bg {display:none;position:fixed;background-color:rgba(0,0,0,0.3);width:100%;height:100%;z-index:1000;top:0;}" & vbCrlf
	strHtml = strHtml & "		.hr-pop-box {position:fixed;background-color:#fff;width:0;height:100%;z-index:1001;top:0;overflow:hidden;box-sizing:border-box;}" & vbCrlf
	strHtml = strHtml & "		.hr-pop-tit {border-bottom:1px solid #eee;height:46px;line-height:46px;box-sizing:border-box;padding:0 20px;}" & vbCrlf
	strHtml = strHtml & "		.hr-pop-tit .close {border:2px solid #59bfe4;height:35px;line-height:35px;width:35px;box-sizing:border-box;text-align:center;border-radius:50%;background-color:#2fabd8;color:#fff;}" & vbCrlf
	strHtml = strHtml & "		.popbox {box-sizing:border-box;overflow-y:auto;height:100%;}" & vbCrlf
	strHtml = strHtml & "		.item-list {box-sizing:border-box;padding:15px;}" & vbCrlf
	strHtml = strHtml & "		.item-list em {border-bottom:1px solid #eee;line-height:40px;padding:0 10px}" & vbCrlf
	strHtml = strHtml & "		.popbtn {padding:10px;}" & vbCrlf
	strHtml = strHtml & "		.filter-box ul li {padding:10px;border-bottom:3px solid #eee}" & vbCrlf
	strHtml = strHtml & "		.filter-box ul li em {font-size:0.8rem;color:#999;}" & vbCrlf
	strHtml = strHtml & "		.filter-box ul li h5 {font-size:0.8rem;}" & vbCrlf
	strHtml = strHtml & "		.filter-box ul li h3 b {font-size:0.8rem;color:#999;padding-left:10px;font-weight:normal}" & vbCrlf
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
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择项目：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Item"" class=""weui-input"" id=""Item"" type=""text"" value=""" & tItemName & """ readonly>" & vbCrlf
	Response.Write "			<input name=""ItemID"" class=""weui-input"" id=""ItemID"" type=""hidden"" value=""" & tItemID & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft itempop"" data-name=""Item""><i class=""hr-icon"">&#xef63;</i>选择</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">开始日期：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input type=""text"" class=""weui-input verify"" name=""StartDate"" id=""StartDate"" value=""" & tStartDate & """ readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">结束日期：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input type=""text"" class=""weui-input verify"" name=""EndDate"" id=""EndDate"" value=""" & tEndDate & """ readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""popbtn""><em class=""weui-btn weui-btn_primary"" id=""filterPost"">筛选</em></div>" & vbCrlf
	Response.Write "	<div class=""popbtn""><em class=""weui-btn weui-btn_primary"" id=""goh"">不筛选，直接评价</em></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf


	Response.Write "<div class=""filter-box"">" & vbCrlf
	Response.Write "	<div></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".hr-navmenu-main"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var tnext = $(this).next("".nav-child"");tnext.toggle();" & vbCrlf
	strHtml = strHtml & "		var dis = tnext.css(""display"");" & vbCrlf
	strHtml = strHtml & "		if(dis == ""block""){ $(this).find("".more i"").html(""&#xea45;"");}else{$(this).find("".more i"").html(""&#xea44;"");}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").append(""我的""); $("".navMenu em i"").html(""&#xf321;"");" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ location.href=""" & ParmPath & "Evaluate/TeachQuality.html""; });" & vbCrlf

	strHtml = strHtml & "	var str1 = ""<div class='hr-pop-bg'></div><div class='hr-pop-box'><div class='hr-rows hr-pop-tit'><em class='tit'>请选择</em><em class='close'><i class='hr-icon'>&#xe960;</i></em></div><div class='popbox'><p class='hr-pop-load'><i class='hr-icon'>&#xefe3;</i></p></div></div>"";" & vbCrlf	'提前增加选择层
	strHtml = strHtml & "	$(""body"").append(str1);" & vbCrlf
	strHtml = strHtml & "	$("".verify"").calendar({dateFormat: 'yyyy-mm-dd'});" & vbCrlf
	strHtml = strHtml & "	var arrItem =[" & GetSelectOptionItem() & "];" & vbCrlf		'业绩项目数据
	strHtml = strHtml & "	$("".itempop"").on(""click"",function(){" & vbCrlf			'选择项目
	strHtml = strHtml & "		var el1=$(this).data(""name""), popw = $("".hr-pop-box"").width(), str1="""";" & vbCrlf
	strHtml = strHtml & "		$("".hr-pop-bg"").fadeIn();$("".hr-pop-box"").animate({width:'60%'});" & vbCrlf
	strHtml = strHtml & "		var tid = $(""#Item"").data(""values""), teacher=$(""#TeacherID"").val();" & vbCrlf
	strHtml = strHtml & "		str1+=""<div class='item-list'>"";" & vbCrlf
	strHtml = strHtml & "		for(var i=0;i<arrItem.length;i++){" & vbCrlf
	strHtml = strHtml & "			str1+=""<em data-itemid='"" + arrItem[i].value + ""'>"" + arrItem[i].title + ""</em>"";" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "		str1+=""</div>"";" & vbCrlf
	strHtml = strHtml & "		$("".popbox"").html(str1);" & vbCrlf
	strHtml = strHtml & "		$("".hr-pop-tit .close"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "			$("".hr-pop-bg"").fadeOut();$("".hr-pop-box"").animate({width:'0'});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$("".item-list em"").on(""click"", function(){" & vbCrlf
	strHtml = strHtml & "			var itemid = $(this).data(""itemid""), item = $(this).text();" & vbCrlf
	strHtml = strHtml & "			console.log(item);" & vbCrlf
	strHtml = strHtml & "			$(""#ItemID"").val(itemid); $(""#Item"").val(item);" & vbCrlf
	strHtml = strHtml & "			$("".hr-pop-bg"").fadeOut();$("".hr-pop-box"").animate({width:'0'});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#filterPost"").on(""click"",function(){" & vbCrlf				'筛选课程
	strHtml = strHtml & "		var itemid=$(""#ItemID"").val(), sdate=$(""#StartDate"").val(), edate=$(""#EndDate"").val();" & vbCrlf
	strHtml = strHtml & "		$.get(""" & ParmPath & "CourseSelect/getItemCourse.html"",{ItemID:itemid, StartDate:sdate, EndDate:edate}, function(strForm){" & vbCrlf
	strHtml = strHtml & "			$("".filter-box"").html(strForm);" & vbCrlf
	strHtml = strHtml & "			$("".filter-box ul li"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "				location.href=""" & ParmPath & "Evaluate/EditQuality.html?AddNew=True&ItemID="" + itemid + ""&CourseID="" + $(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	$(""#goh"").on(""click"",function(){" & vbCrlf				'不筛选课程
	strHtml = strHtml & "		location.href=""" & ParmPath & "Evaluate/EditQuality.html?AddNew=True"";" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub getItemCourse()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tStartDate : tStartDate = Trim(Request("StartDate"))
	Dim tEndDate : tEndDate = Trim(Request("EndDate"))
	Dim tSheetName : tSheetName = "HR_Sheet_" & tItemID
	Dim sTime, eTime : ErrMsg = ""
	If Not(ChkTable(tSheetName)) Then ErrMsg = "未找该项目的数据！"
	If Not(IsDate(tStartDate)) Then ErrMsg = "开始日期不正确！"
	If Not(IsDate(tEndDate)) Then ErrMsg = "结束日期不正确！"
	If DateDiff("s", tStartDate, tEndDate) < 0 Then ErrMsg = "开始日期不能大于结束日期！"

	If ErrMsg<>"" Then Response.Write "<div class=""errtips"">" & ErrMsg & "</div>" : Exit Sub

	sTime = ConvertDateToNum(tStartDate) + 2
	eTime = ConvertDateToNum(tEndDate) + 2
	sql = "Select top 100 * From " & tSheetName & " Where scYear=" & DefYear
	sql = sql & " And VA4>=" & sTime & " And VA4<=" & eTime
	sql = sql & " Order By VA4 ASC"
	Set rs = Conn.Execute(sql)
		If Not(rs.BOF And rs.EOF) Then
			Response.Write "<ul class=""itemlist"">"
			Do While Not rs.EOF
				Response.Write "<li data-id=""" & rs("ID") & """><h3>" & rs("VA8") & "<b>" & ConvertNumDate(rs("VA4")) & "</b></h3><h4>" & rs("VA9") & "</h4><h5>" & rs("VA10") & "</h5><h5>" & rs("VA11") & " " & rs("VA12") & "</h5><em>" & rs("VA2") & " " & rs("VA1") & " 第" & rs("VA7")
				Response.Write "节[" & GetPeriodTime(Trim(rs("VA11")), rs("VA7"), 0) & "]</em></li>"
				rs.MoveNext
			Loop
			Response.Write "</ul>"
		End If
	Set rs = Nothing
End Sub

Function GetSelectOptionItem()				'取考核项目下拉
	Dim iFun, funItem, rsFun, sqlFun
	sqlFun = "Select ClassID, ClassName From HR_Class Where ModuleID=1001 And Child=0 And Template in('TempTableA','TempTableC')"
	sqlFun = sqlFun & " Order By RootID, OrderID"
	Set rsFun = Conn.Execute(sqlFun)
		If Not(rsFun.BOF And rsFun.EOF) Then
			iFun = 0
			Do While Not rsFun.EOF
				If iFun > 0 Then funItem = funItem & ","
				funItem = funItem & "{title:""" & rsFun("ClassName") & """,value:""" & rsFun("ClassID") & """}"
				rsFun.MoveNext
				iFun = iFun + 1
			Loop
		End If
	Set rsFun = Nothing
	GetSelectOptionItem = funItem
End Function
%>