<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->

<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "通讯录"
If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "GetTeacher" Call GetTeacher()
	Case "Message" Call Message()
	Case "View" Call View()
	Case "SelectTeacher" Call SelectTeacher()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-box {border:0px solid #F60;background-color: rgba(255,255,255,1);box-sizing: border-box;overflow-y:auto;position:absolute;top:100px;left:0;bottom:0;right:0;z-index:10}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-letter {position:fixed;height:100%;width:25px;right:0;top:100px;background-color:#e3e3e3;z-index:11}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-letter ul {display:flex;text-align:center;flex-wrap:wrap;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-letter li {display:block;text-align:center;width:100%;font-size:0.8rem}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-addbook-list {overflow-y:auto;}" & vbCrlf
	tmpHtml = tmpHtml & "		.addInfo {line-height:25px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.addInfo .num_name b {padding-left:initial;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml
	Response.Write getHeadNav(0)

	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-search-bar"" id=""searchBar"">" & vbCrlf
	Response.Write "	<form class=""weui-search-bar__form"">" & vbCrlf
	Response.Write "		<div class=""weui-search-bar__box"">" & vbCrlf
	Response.Write "			<i class=""weui-icon-search""></i><input type=""search"" class=""weui-search-bar__input"" id=""searchInput"" placeholder=""搜索"" required="""">" & vbCrlf
	Response.Write "			<a href=""javascript:"" class=""weui-icon-clear"" id=""searchClear""></a>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<label class=""weui-search-bar__label"" id=""searchText""><i class=""weui-icon-search""></i><span>搜索姓名/拼音/工号</span></label>" & vbCrlf
	Response.Write "	</form>" & vbCrlf
	Response.Write "	<a href=""#"" class=""weui-search-bar__cancel-btn"" id=""searchCancel"">取消</a>" & vbCrlf
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
	'Response.Write "	<div class=""initials""><ul><li><i class=""hr-icon"">&#xef3e;</i></li></ul></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	'Response.Write "<span id=""get1"">取值</span>" & vbCrlf
	'response.write Asc("通") & "  dsdf"
	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf

	tmpHtml = tmpHtml & "	var boxHeight = $("".hr-addbook-box"").height()-10;console.log(""BOX_H:"" + boxHeight);" & vbCrlf
	tmpHtml = tmpHtml & "	var liH = boxHeight/28;" & vbCrlf
	tmpHtml = tmpHtml & "	$('.hr-addbook-letter li').height(liH);" & vbCrlf
	tmpHtml = tmpHtml & "	getTeacherList("""");" & vbCrlf
	tmpHtml = tmpHtml & "	$("".hr-addbook-letter li"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var lmd =$(this).data(""id"");console.log(lmd);" & vbCrlf
	tmpHtml = tmpHtml & "		getTeacherList(lmd);" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	$(""#searchInput"").bind(""input propertychange"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		getTeacherList($(this).val());" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	function getTeacherList(fLetter){" & vbCrlf		'取教师数据
	tmpHtml = tmpHtml & "		$.get(""" & ParmPath & "Directories/GetTeacher.html"",{Type:3,Letter:fLetter}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#listGroup"").html(rsStr);" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	}" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1)
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub SelectTeacher()
	Dim reObjTxt : reObjTxt = Trim(Request("reObjTxt"))
	Dim reObjValue : reObjValue = Trim(Request("reObjValue"))

	SiteTitle = "选择教师"
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
	Response.Write "				<i class=""weui-icon-search""></i><input type=""search"" class=""weui-search-bar__input"" id=""searchInput"" autocomplete=""off"" placeholder=""搜索"" required="""">" & vbCrlf
	Response.Write "				<a href=""javascript:"" class=""weui-icon-clear"" id=""searchClear""></a>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<label class=""weui-search-bar__label"" id=""searchText""><i class=""weui-icon-search""></i><span>搜索姓名/拼音/工号</span></label>" & vbCrlf
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
	'Response.Write "	<div class=""initials""><ul><li><i class=""hr-icon"">&#xef3e;</i></li></ul></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	'Response.Write "<span id=""get1"">取值</span>" & vbCrlf
	'response.write Asc("通") & "  dsdf"
	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf

	tmpHtml = tmpHtml & "	var boxHeight = $("".hr-addbook-box"").height()-10;" & vbCrlf
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
	tmpHtml = tmpHtml & "		$.get(""" & ParmPath & "Directories/GetTeacher.html"",{Type:0,Letter:fLetter}, function(rsStr){" & vbCrlf
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
Sub GetTeacher()
	Dim sqlList, rsList, tFirstWord, tTel, tTel1, tPRZC
	Dim listType : listType = HR_Clng(Request("Type"))
	Dim tLetter : tLetter = Trim(ReplaceBadChar(Request("Letter")))
	sqlList = "Select Top 2000 *,Left(XMJP, 1) AS Letter From HR_Teacher Where Cast(YGDM As int)>0 And XMJP<>''"
	If listType > 0 Then sqlList = sqlList & " And ApiType=" & listType
	If HR_IsNull(tLetter) = False Then
		tFirstWord = UCase(Left(tLetter, 1))
		If Asc(tFirstWord) > 64 And Asc(tFirstWord) < 91 Then
			sqlList = sqlList & " And XMJP like('" & tLetter & "%')"
		ElseIf HR_Clng(tFirstWord) > 0 Then
			sqlList = sqlList & " And YGDM like('" & tLetter & "%')"
		Else
			sqlList = sqlList & " And YGXM like('%" & tLetter & "%')"
		End If
	End If
	sqlList = sqlList & " Order By Letter ASC, YGXM ASC"
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
				'If m > 0 Then strList = strList & ","
				tTel = ""
				tPRZC = rsList("PRZC") : tPRZC = Replace(tPRZC, "无职称", "")
				py = UCase(Left(rsList("XMJP"), 1))
				If HR_IsNull(rsList("SJHM")) = False Then tTel = "<br>手机：" & Trim(rsList("SJHM")) & " " & Trim(rsList("DH"))
				If ZM(ASC(py)-64,1) = "0" Then
					Response.Write "<div class=""sort_letter"" id=""l_" & py & """>" & py & "</div>"
					ZM(ASC(py)-64,1) = "1"
				End If
				Response.Write "	<div class=""sort_list addInfo"">" & vbCrlf
				Response.Write "		<div class=""num_name"" data-ygdm=""" & rsList("YGDM") & """ data-ygxm=""" & rsList("YGXM") & """>" & rsList("YGXM") & "<b> [工号：" & rsList("YGDM") & "]</b><br><b>" & rsList("KSMC") & " " & tPRZC & "</b></div>" & vbCrlf
				Response.Write "	</div>" & vbCrlf
				rsList.MoveNext
				m = m + 1
			Loop
		Else
			Response.Write "	<div class=""sort_list"">没有教师</div>" & vbCrlf
		End If
	Set rsList = Nothing
End Sub
%>