<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "科室管理"

Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "View" Call View()
	Case "SelectDepart" Call SelectDepart()
	Case "GetDepart" Call GetDepart()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim Count1		'汇总科室
	Set rsTmp = Conn.Execute("Select count(0) From HR_Department Where DepartmentID>0")
		Count1 = HR_Clng(rsTmp(0))
	Set rsTmp = Nothing

	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells {margin:0;}" & vbCrlf

	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-cell"">" & vbCrlf
	Response.Write "	<div class=""weui-cell__bd"">科室数：</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell__ft"">" & Count1 & "</div>"  & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""weui-cells"">" & vbCrlf

	sqlTmp = "Select a.*,(Select Count(0) From HR_Teacher Where KSDM=a.KSDM) As TearcherNum From HR_Department a Where a.DepartmentID>0"
	sqlTmp = sqlTmp & " Order By a.RootID ASC, a.OrderID ASC"
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Dim CountTearcher
			Do While Not rsTmp.EOF
				Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageDepart/View.html?ID=" & rsTmp("DepartmentID") & """>" & vbCrlf
				Response.Write "		<div class=""weui-cell__bd viewMSG"" data-id=""" & rsTmp("DepartmentID") & """><p><i class=""hr-icon"">&#xecad;</i>" & rsTmp("KSMC") & "</p></div>" & vbCrlf
				Response.Write "		<div class=""weui-cell__ft"">教师" & HR_Clng(rsTmp("TearcherNum")) & "人</div>" & vbCrlf
				Response.Write "	</a>" & vbCrlf
				rsTmp.MoveNext
			Loop
		Else
			Response.Write "	<a class=""weui-cell weui-cell_access"" href=""javascript:;"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>暂时没有科室</p></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
			Response.Write "	</a>" & vbCrlf
		End If
	Set rsTmp = Nothing
	Response.Write "</div>" & vbCrlf

	Response.Write "<div id=""half"" class=""weui-popup__container popup-bottom"">" & vbCrlf
	Response.Write "	<div class=""weui-popup__overlay""></div>" & vbCrlf
	Response.Write "	<div class=""weui-popup__modal"">" & vbCrlf
	Response.Write "		<div class=""toolbar""><div class=""toolbar-inner""><span class=""picker-button close-popup"">关闭</span><h1 class=""title"">发布科室</h1></div></div>" & vbCrlf
	Response.Write "		<div class=""modal-content"">" & vbCrlf
	Response.Write "			<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "				<div class=""weui-cell"">" & vbCrlf
	Response.Write "					<div class=""weui-cell__bd""><textarea class=""weui-textarea title"" name=""title"" id=""title"" placeholder=""请输入通知标题"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "				</div>" & vbCrlf
	Response.Write "				<div class=""weui-cell"">" & vbCrlf
	Response.Write "					<div class=""weui-cell__bd""><textarea class=""weui-textarea content"" name=""content"" placeholder=""请输入通知内容"" rows=""5""></textarea></div>" & vbCrlf
	Response.Write "				</div>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<div class=""weui-btn-area""><em class=""weui-btn weui-btn_primary"" id=""subPost"">提　交</em></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-float-btn"">" & vbCrlf
	Response.Write "	<span class=""navExtend open-popup"" data-target=""#half""><i class=""hr-icon"">&#xf3c0;</i></span>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Manage/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf

	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

Sub View()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim tKSDM, tKSMC, tKSDD, tTearcherNum
	Set rs = Conn.Execute("Select a.*,(Select Count(0) From HR_Teacher Where KSDM=a.KSDM) As TearcherNum From HR_Department a Where a.DepartmentID=" & tmpID)
		If Not(rs.BOF And rs.EOF) Then
			tKSDM = Trim(rs("KSDM"))
			tKSMC = Trim(rs("KSMC"))
			tKSDD = Trim(rs("KSDD"))
			tTearcherNum = HR_Clng(rs("TearcherNum"))
		End If
	Set rs = Nothing
	SiteTitle = "科室详情"

	strHtml = "<link type=""text/css"" href=""" & InstallDir & "Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	strHtml = strHtml & "		.backbtn {padding:10px 5px;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)
	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<div class=""weui-cells"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">科室名称：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tKSMC & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">地　址：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tKSDD & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">教师数：</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & tTearcherNum & "</div>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-fix backbtn""><button type=""button"" name=""back"" class=""weui-btn weui-btn_plain-default"" id=""back"">返回</button></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "ManageDepart/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".backbtn"").on(""click"",function(){ location.href=""" & ParmPath & "ManageDepart/Index.html""; });" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write ReplaceCommonLabel(strFootHtml)
End Sub

Sub SelectDepart()
	Dim reObjTxt : reObjTxt = Trim(Request("reObjTxt"))
	Dim reObjValue : reObjValue = Trim(Request("reObjValue"))

	SiteTitle = "选择科室"
	tmpHtml = "<link type=""text/css"" href=""[@Web_Dir]Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	tmpHtml = tmpHtml & "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-box {border:0px solid #F60;background-color: rgba(255,255,255,1);box-sizing: border-box;overflow-y:auto;position:absolute;top:50px;left:0;bottom:0;right:0;z-index:10}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-letter {position:fixed;height:100%;width:25px;right:0;top:50px;background-color:#e3e3e3;z-index:11}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-letter ul {display:flex;text-align:center;flex-wrap:wrap;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-letter li {display:block;text-align:center;width:100%;font-size:0.8rem}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-list {overflow-y:auto;}" & vbCrlf
	tmpHtml = tmpHtml & "		.addInfo {line-height:25px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.addInfo .num_name b {padding-left:initial;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-sobar {background-color: #efeff4;} .hr-sobar .btnClose i {font-size:1.5rem;color:#f52}" & vbCrlf
	tmpHtml = tmpHtml & "		#searchCancel i {font-size:1.2rem;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-rows hr-sobar"">" & vbCrlf
	Response.Write "	<div class=""hr-row-fill"">" & vbCrlf
	Response.Write "	<div class=""weui-search-bar"" id=""searchBar"">" & vbCrlf
	Response.Write "		<form class=""weui-search-bar__form"">" & vbCrlf
	Response.Write "			<div class=""weui-search-bar__box"">" & vbCrlf
	Response.Write "				<i class=""weui-icon-search""></i><input type=""search"" class=""weui-search-bar__input"" id=""searchInput"" placeholder=""搜索"" required="""">" & vbCrlf
	Response.Write "				<a href=""javascript:"" class=""weui-icon-clear"" id=""searchClear""></a>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<label class=""weui-search-bar__label"" id=""searchText""><i class=""weui-icon-search""></i><span>搜索科室名称/科室代码</span></label>" & vbCrlf
	Response.Write "		</form>" & vbCrlf
	Response.Write "		<a href=""#"" class=""weui-search-bar__cancel-btn"" id=""searchCancel""><i class=""hr-icon"">&#xee30;</i></a>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""btnClose"">" & vbCrlf
	Response.Write "		<em><i class=""hr-icon"">&#xec58;</i></em>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-addbook-box"">" & vbCrlf
	Response.Write "	<div class=""hr-addbook-letter""><ul>" & vbCrlf
	Response.Write "	<li data-id=""*""><i class=""hr-icon"">&#xef3e;</i></li>" & vbCrlf
	For i = 65 To 90
		Response.Write "	<li data-id=""" & Chr(i) & """>" & Chr(i) & "</li>" & vbCrlf
	Next
	Response.Write "	<li data-id=""#"">#</li>" & vbCrlf
	Response.Write "	</ul></div>" & vbCrlf

	Response.Write "	<div class=""hr-addbook-list"" id=""listGroup"">" & vbCrlf
	Response.Write "		<div class=""weui-loadmore""><i class=""weui-loading""></i><span class=""weui-loadmore__tips"">正在加载</span></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf

	tmpHtml = tmpHtml & "	var boxHeight = $("".hr-addbook-box"").height()-10;console.log(""BOX_H:"" + boxHeight);" & vbCrlf
	tmpHtml = tmpHtml & "	var liH = boxHeight/28;" & vbCrlf
	tmpHtml = tmpHtml & "	$('.hr-addbook-letter li').height(liH);" & vbCrlf
	tmpHtml = tmpHtml & "	getTeacherList('');" & vbCrlf
	tmpHtml = tmpHtml & "	$("".hr-addbook-letter li"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var lmd =$(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		getTeacherList(lmd);" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	$("".btnClose"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var elParent = $(window.parent.document);" & vbCrlf
	tmpHtml = tmpHtml & "		elParent.find(""#full"").hide();" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	$(""#searchInput"").bind(""input propertychange"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		getTeacherList($(this).val());" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	function getTeacherList(fLetter){" & vbCrlf		'取教师数据
	tmpHtml = tmpHtml & "		$.get(""" & ParmPath & "ManageDepart/GetDepart.html"",{Type:0,Letter:fLetter}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#listGroup"").html(rsStr);" & vbCrlf
	tmpHtml = tmpHtml & "			$("".num_name"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "				var elParent = $(window.parent.document);" & vbCrlf
	tmpHtml = tmpHtml & "				elParent.find(""#full"").hide();" & vbCrlf
	tmpHtml = tmpHtml & "				var ygdm = $(this).data(""ygdm"");" & vbCrlf
	tmpHtml = tmpHtml & "				elParent.find(""#" & reObjTxt & """).val($(this).data(""ygxm""));" & vbCrlf
	tmpHtml = tmpHtml & "				elParent.find(""#" & reObjValue & """).val(ygdm);" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	}" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub
Sub GetDepart()
	Dim sqlList, rsList, tFirstWord
	Dim listType : listType = HR_Clng(Request("Type"))
	Dim tLetter : tLetter = Trim(ReplaceBadChar(Request("Letter")))
	sqlList = "Select *,Left(PYDM, 1) As Letter From HR_Department Where Cast(KSDM As int)>0 And PYDM<>''"
	If HR_IsNull(tLetter) = False Then
		tFirstWord = UCase(Left(tLetter, 1))
		If Asc(tFirstWord) > 64 And Asc(tFirstWord) < 91 Then
			sqlList = sqlList & " And PYDM like('" & tLetter & "%')"
		ElseIf HR_Clng(tFirstWord) > 0 Then
			sqlList = sqlList & " And KSDM like('" & tLetter & "%')"
		Else
			sqlList = sqlList & " And KSMC like('%" & tLetter & "%')"
		End If
	End If
	sqlList = sqlList & " Order By Letter ASC, KSMC ASC"
	Set rsList = Server.CreateObject("ADODB.RecordSet")
		rsList.Open(sqlList), Conn, 1, 1
		If Not(rsList.BOF And rsList.EOF) Then
			m = 0
			Redim ZM(26,1)
			Dim py
			For i = 65 To 90
				ZM(i-64,0) = Chr(i)
				ZM(i-64,1) = "0"
			Next
			Do While Not rsList.EOF
				py = UCase(Left(rsList("PYDM"), 1))
				If ZM(ASC(py)-64,1) = "0" Then
					Response.Write "<div class=""sort_letter"" id=""l_" & py & """>" & py & "</div>"
					ZM(ASC(py)-64,1) = "1"
				End If
				Response.Write "	<div class=""sort_list addInfo"">" & vbCrlf
				Response.Write "		<div class=""num_name"" data-ygdm=""" & rsList("KSDM") & """ data-ygxm=""" & rsList("KSMC") & """>" & rsList("KSMC") & "<b> [" & rsList("KSDM") & "]</b><br><b>地址：" & rsList("KSDD") & "</b></div>" & vbCrlf
				Response.Write "	</div>" & vbCrlf
				rsList.MoveNext
				m = m + 1
			Loop
		Else
			Response.Write "	<div class=""sort_list"">没有科室</div>" & vbCrlf
		End If
	Set rsList = Nothing
End Sub
%>