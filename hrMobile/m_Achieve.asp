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
SiteTitle = "我的业绩"

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "List" Call List()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tEduYear : tEduYear = HR_CLng(Request("EduYear"))
	If tEduYear<2017 Or tEduYear>Year(Date()) Then tEduYear = DefYear			'若未指定学年，则为当前学年
	strHtml = "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-inline {border-bottom:1px solid #e3e3e3;padding:8px;background:#eee;}" & vbCrlf
	strHtml = strHtml & "		.hr-inline em:first-child {width:5em;text-align:center;color:#777;}" & vbCrlf
	strHtml = strHtml & "		.hr-inline .yearbar {width:140px; border-radius:30px;background:#fff; position:relative; padding-right:10px}" & vbCrlf
	strHtml = strHtml & "		.hr-inline .yearbar .yearinput {border:0; width:100%; text-align:center; font-size:18px;}" & vbCrlf
	strHtml = strHtml & "		.hr-inline .yearbar span {color:#f30; position:absolute; top:-2px; right:7px;font-size:18px;}" & vbCrlf
	strHtml = strHtml & "		.viewPanel li b {width:auto;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-fix"">" & vbCrlf
	Response.Write "	<div class=""hr-rows hr-inline"">" & vbCrlf
	Response.Write "		<em class=""hr-item hr-fixed"">总学时数</em><em class=""hr-item hr-grow"">"
	sqlTmp = "Select top 1 SumScore From HR_KPI_SUM Where YGDM>0 And YGDM=" & HR_Clng(UserYGDM) & " And scYear=" & tEduYear
	Set rsTmp = Conn.Execute(sqlTmp)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Response.Write "" & HR_CDbl(rsTmp(0)) & ""
		Else
			Response.Write 0
		End If
	Set rsTmp = Nothing
	Response.Write "</em>" & vbCrlf
	Response.Write "		<tt class=""hr-item hr-fixed"">学年</tt>" & vbCrlf
	Response.Write "		<tt class=""hr-item hr-fixed yearbar""><input name=""EduYear"" id=""EduYear"" class=""yearinput"" value=""" & tEduYear-1 & "-" & tEduYear & """ readonly /><span><i class=""hr-icon"">&#xf0d7;</i></span></tt>"  & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-fix"">" & vbCrlf	
	Dim tItemName, tTemplate, tStuType, tSheetName, sql1, rs1, noPassNum, PassNum
	Dim noData : noData = False
	Set rsTmp = Conn.Execute("Select * From HR_Class Order By ClassType ASC,RootID ASC,OrderID ASC")
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Response.Write "	<ul class=""hr-panel-item"">" & vbCrlf
			Do While Not rsTmp.EOF
				If rsTmp("Child") = 0 Then
					tItemName = Trim(rsTmp("ClassName"))
					tTemplate = Trim(rsTmp("Template"))
					tStuType = Trim(rsTmp("StudentType"))
					tSheetName = "HR_Sheet_" & rsTmp("ClassID")

					If ChkTable(tSheetName) Then
						noPassNum = 0 : PassNum = 0
						Set rs1 = Conn.Execute("Select Sum(VA3) From " & tSheetName & " Where Passed=" & HR_False & " And VA1=" & HR_Clng(UserYGDM) & " And scYear=" & tEduYear)
							noPassNum = HR_CDbl(rs1(0))
						Set rs1 = Nothing

						sql1 = "Select Sum(VA3) From " & tSheetName & " Where Passed=" & HR_True & " And VA1=" & HR_Clng(UserYGDM) & " And scYear=" & tEduYear
						Set rs1 = Conn.Execute(sql1)
							PassNum = HR_CDbl(rs1(0))
						Set rs1 = Nothing
						If noPassNum > 0 Or PassNum > 0 Then
							Response.Write "<div class=""hr-flex_item"" data-id=""" & rsTmp("ClassID") & """>"
							Response.Write "<a class=""hr-navmenu"" href=""" & ParmPath & "Course/List.html?ItemID=" & rsTmp("ClassID") & "&EduYear=" & tEduYear & """>"
							Response.Write "<em class=""title""><i class=""hr-icon"">&#xf34f;</i>" & tItemName & "</em><em class=""tips"">[已审]" & PassNum
							Response.Write "/[未审]" & noPassNum
							Response.Write "</em><em class=""more""><i class=""hr-icon"">&#xef91;</i></em></a>"
							Response.Write "</div>"
							noData = True
						End If
					End If
				End If
				rsTmp.MoveNext
			Loop
			Response.Write "	</ul>" & vbCrlf
		End If
	Set rsTmp = Nothing
	Response.Write "</div>" & vbCrlf
	If noData = False Then Response.Write "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef61;</i></h2><h3>您在" & tEduYear-1 & "-" & tEduYear & "学年暂无业绩！</h3></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".addNew a"").css(""display"",""none"");" & vbCrlf
	Dim tmpYearJson, cYear : cYear = Year(Date()) + 1
	tmpYearJson = ""
	For i = cYear To cYear-5 Step -1
		If i<cYear Then tmpYearJson = tmpYearJson & ","
		tmpYearJson = tmpYearJson & "{title:""" & i-1 & "-" & i & """, value:""" & i & """}"
	Next
	strHtml = strHtml & "	$(""#EduYear"").select({title:""选择学年""," & vbCrlf
	strHtml = strHtml & "		items:[" & tmpYearJson & "]," & vbCrlf
	strHtml = strHtml & "		onClose:function(res){" & vbCrlf
	strHtml = strHtml & "			location.href=""" & ParmPath & "Achieve/Index.html?EduYear="" + res.data.values;" & vbCrlf
	'strHtml = strHtml & "			console.log(res.data.values);" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub
%>