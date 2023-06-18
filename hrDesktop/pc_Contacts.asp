<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incCNtoPY.asp"-->
<!--#include file="./incCommon.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "通讯录"

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "Float" Call FloatWin()
	Case "GetListData" Call GetListData()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	
End Sub

Sub FloatWin()
	SiteTitle = "通讯录查询"
	Dim tType : tType = HR_Clng(Request("Type"))
	Dim tKey : tKey = Trim(ReplaceBadChar(Request("Key")))
	Dim tValue : tValue = Trim(ReplaceBadChar(Request("Value")))

	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-search-bar .hr-search-box, .hr-search-bar .hr-search-label {padding-left:10px;border-radius:5px;height:28px;line-height:28px;display:block;background-color:#fff;margin:6px 10px;color:#777;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-search-bar .hr-search-label {display:none;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-search-input {border:0;line-height:25px;width:93%;background-color:transparent;}" & vbCrlf
	'tmpHtml = tmpHtml & "		.addInfo .num_name b {padding-left:initial;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-sobar {background-color: #efeff4;} .hr-sobar .btnClose {padding-right:10px;cursor: pointer;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-sobar .btnClose i {font-size:1.5rem;color:#f52}" & vbCrlf
	tmpHtml = tmpHtml & "		#searchCancel i {font-size:1.2rem;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf

	strHtml = getPageHead("Index", 1)		'Header
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	Call ReplaceCommonLabel(strHtml)
	Response.Write strHtml

	Response.Write "<div class=""hr-rows hr-sobar"">" & vbCrlf
	Response.Write "	<div class=""hr-row-fill"">" & vbCrlf
	Response.Write "		<div class=""hr-search-bar"" id=""searchBar"">" & vbCrlf
	Response.Write "			<div class=""hr-search-box"">" & vbCrlf
	Response.Write "				<i class=""hr-icon"">&#xef27;</i><input type=""search"" class=""hr-search-input"" id=""searchInput"" placeholder=""搜索"" required="""">" & vbCrlf
	Response.Write "				<span class=""weui-icon-clear"" id=""searchClear""></span>" & vbCrlf
	Response.Write "			</div>" & vbCrlf
	Response.Write "			<label class=""hr-search-label"" id=""searchText""><i class=""hr-icon"">&#xef27;</i><span>搜索姓名/拼音/工号" & tKey & "</span></label>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
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
	Response.Write "	<div class=""hr-addbook-list"" id=""listGroup"" value="""" code="""" name="""">" & vbCrlf
	Response.Write "		<div class=""layui-layer layui-layer-loading""><div class=""layui-layer-content layui-layer-loading1""></div></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	var boxHeight = $("".hr-addbook-box"").height()-10, liH = boxHeight/28;" & vbCrlf
	tmpHtml = tmpHtml & "	$('.hr-addbook-letter li').height(liH);" & vbCrlf
	tmpHtml = tmpHtml & "	getTeacherList(3, """");" & vbCrlf
	tmpHtml = tmpHtml & "	$("".hr-addbook-letter li"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var lmd =$(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		getTeacherList(0,lmd);" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	
	tmpHtml = tmpHtml & "	$(""#searchInput"").bind(""input propertychange"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		getTeacherList(0,$(this).val());" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "	var index = parent.layer.getFrameIndex(window.name);" & vbCrlf
	tmpHtml = tmpHtml & "	$("".btnClose"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		parent.layer.close(index);" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "	function getTeacherList(fType, fLetter){" & vbCrlf		'取教师数据
	tmpHtml = tmpHtml & "		$.get(""" & ParmPath & "/Contacts/GetListData.html"",{Type:fType, Letter:fLetter}, function(rsStr){" & vbCrlf
	tmpHtml = tmpHtml & "			$(""#listGroup"").html(rsStr);" & vbCrlf
	tmpHtml = tmpHtml & "			$("".num_name"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "				var elcode=$(""#listGroup"").attr(""code""), elname=$(""#listGroup"").attr(""name"");" & vbCrlf
	tmpHtml = tmpHtml & "				var iframe1=$(""#listGroup"").attr(""value""), elParent = $(""#"" + iframe1,parent.document.body).contents();" & vbCrlf
	tmpHtml = tmpHtml & "				var ygdm = $(this).data(""ygdm""), ygxm = $(this).data(""ygxm"");" & vbCrlf
	tmpHtml = tmpHtml & "				elParent.find(""#"" + elname).val(ygxm); elParent.find(""#"" + elcode).val(ygdm);" & vbCrlf
	tmpHtml = tmpHtml & "				parent.layer.close(index);" & vbCrlf
	tmpHtml = tmpHtml & "			});" & vbCrlf
	tmpHtml = tmpHtml & "		});" & vbCrlf
	tmpHtml = tmpHtml & "	}" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = getPageFoot("Index", 0) & "" & vbCrlf
	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub GetListData()
	Dim sqlList, rsList, tFirstWord, tTel, tTel1, tPRZC
	Dim listType : listType = HR_Clng(Request("Type"))
	Dim tLetter : tLetter = Trim(ReplaceBadChar(Request("Letter")))
	sqlList = "Select Top 3000 * From HR_Teacher Where Cast(YGDM As int)>0 And ltrim(rtrim(XMJP))!=''"
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
	sqlList = sqlList & " Order By YGXM collate Chinese_PRC_CS_AS_KS_WS"
	Set rsList = Server.CreateObject("ADODB.RecordSet")
		rsList.Open(sqlList), Conn, 1, 1
		If Not(rsList.BOF And rsList.EOF) Then
			m = 0
			Redim ZM(27,1)
			Dim py
			For i = 65 To 90
				ZM(i-65,0) = Chr(i)
				ZM(i-65,1) = "0"
			Next
			ZM(27,0) = "*" : ZM(27,1) = "0"
			Do While Not rsList.EOF
				'If m > 0 Then strList = strList & ","
				tTel = ""
				tPRZC = rsList("PRZC") : tPRZC = Replace(tPRZC, "无职称", "")
				If HR_Clng(rsList("XMJP")) > 0 Or HR_IsNull(rsList("XMJP")) Then
					Conn.Execute("Update HR_Teacher Set XMJP=""1"" Where TeacherID=" & rsList("TeacherID"))
				End If
				py = Left(UCase(rsList("XMJP")), 1)
				If HR_IsNull(rsList("SJHM")) = False Then tTel = "<br>手机：" & Trim(rsList("SJHM")) & " " & Trim(rsList("DH"))
				If ZM(ASC(py)-65,1) = "0" Then
					Response.Write "<div class=""sort_letter"" id=""l_" & py & """>" & py & "</div>"
					ZM(ASC(py)-65,1) = "1"
				End If
				Response.Write "	<div class=""sort_list"">" & vbCrlf
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