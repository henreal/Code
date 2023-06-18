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
SiteTitle = "业绩报表"

Dim strHeadHtml : strHeadHtml =	getPageHead(1)
Dim strFootHtml : strFootHtml =	getPageFoot(1)

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "ShowList" Call ShowList()
	Case "View" Call View()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()

	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dl {align-items:stretch;border-bottom:1px solid #eee;padding:5px 0;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dt {width:60%;text-align:right;color:#999}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-swap-items dd {flex-grow:2;width:40%;box-sizing: border-box;padding-right:3px}" & vbCrlf
	tmpHtml = tmpHtml & "		.viewbox, .popbtn {padding:10px 0;box-sizing: border-box;background-color:#fff;} .popbtn {padding:10px 5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.closebtn {width:2rem;text-align:center;height:2rem;line-height:2rem;position:fixed;bottom:3rem;right:1rem;background-color:rgba(3,169,244,0.5);color:#fff;z-index:100;display:none;border-radius: 5px;font-size:1.3rem}" & vbCrlf
	tmpHtml = tmpHtml & "		.txtNum0{color:#999} .txtNum1{color:#f30}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", tmpHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){ FastClick.attach(document.body); });" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write ReplaceCommonLabel(strHeadHtml)
	Response.Write ReplaceCommonLabel(getHeadNav(0))
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"" id=""ListBox""></div>" & vbCrlf
	Response.Write "<div id=""fullView"" class=""weui-popup__container"">" & vbCrlf
	Response.Write "	<div class=""weui-popup__overlay""></div>" & vbCrlf
	Response.Write "	<div class=""weui-popup__modal""><div class=""viewbox""></div><div class=""popbtn""><button class=""weui-btn weui-btn_primary close-popup"">关闭</button></div></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf	
	Response.Write "<div class=""hr-shrink-x20""></div>" & vbCrlf
	Response.Write "<div class=""closebtn close-popup"" title=""关闭弹窗""><i class=""hr-icon"">&#xee30;</i></div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Manage/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	ReportItem("""");" & vbCrlf

	strHtml = strHtml & "	function ReportItem(fygdm){" & vbCrlf		'异步加载列表
	strHtml = strHtml & "		$.get(""" & ParmPath & "ManageReport/ShowList.html"",{ygdm:fygdm}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "			$(""#ListBox"").html(rsStr);" & vbCrlf
	strHtml = strHtml & "			$("".view"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "				$.get(""" & ParmPath & "ManageReport/View.html"",{ygdm:$(this).data(""ygdm"")}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "					$("".viewbox"").html(rsStr);$(""#fullView"").popup();$("".closebtn"").show();" & vbCrlf
	strHtml = strHtml & "					$("".close-popup"").on(""click"",function(){$("".closebtn"").hide();});" & vbCrlf
	strHtml = strHtml & "				});" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "	}" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub ShowList()
	sqlTmp = "Select Top 300 * From HR_KPI_SUM Where ID>0 And scYear=" & DefYear
	sqlTmp = sqlTmp & " Order By SumScore DESC"
	Set rsTmp = Server.CreateObject("ADODB.RecordSet")
		rsTmp.Open(sqlTmp), Conn, 1, 1
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			Do While Not rsTmp.EOF
				Response.Write "	<a class=""weui-cell weui-cell_access view"" data-ygdm=""" & rsTmp("YGDM") & """ href=""#"">" & vbCrlf
				Response.Write "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xef2f;</i></div><div class=""weui-cell__bd weui-cell_primary""><p>" & rsTmp("YGXM") & " [" & rsTmp("YGDM") & "]</p></div>" & vbCrlf
				Response.Write "		<div class=""weui-cell__ft"">学时：" & rsTmp("SumScore") & "<br>业绩分：" & rsTmp("TotalScore") & "</div>" & vbCrlf
				Response.Write "	</a>" & vbCrlf
				rsTmp.MoveNext
			Loop
		Else
			Response.Write "	<a class=""weui-cell weui-cell_access"" href=""javascript:;"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd""><p>暂时没有记录</p></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
			Response.Write "	</a>" & vbCrlf
		End If
	Set rsTmp = Nothing
End Sub

Sub View()
	Dim tmpID : tmpID = HR_Clng(Request("ygdm"))
	Dim arrField : arrField = Split(Trim(GetStatisTableField()), "||")
	Dim arrStuType : arrStuType = Split(XmlText("Common", "StudentType", ""), "|")
	Dim tmpMSG, ValueNum, ValueNum1, tFieldKey, tFieldName, tItemName
	If tmpID > 0 Then
		sqlTmp = "Select * From HR_KPI_SUM Where YGDM=" & tmpID & " And scYear=" & DefYear & ""
		Set rsTmp = Conn.Execute(sqlTmp)
			If Not(rsTmp.BOF And rsTmp.EOF) Then
				tmpMSG = tmpMSG & "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
				tmpMSG = tmpMSG & "<div class=""hr-swap-box"">" & vbCrlf
				tmpMSG = tmpMSG & "	<div class=""hr-swap-items"">" & vbCrlf
				tmpMSG = tmpMSG & "		<dl class=""hr-rows""><dt>教　师：</dt><dd>" & Trim(rsTmp("YGXM")) & " [" & rsTmp("YGDM") & "]</dd></dl>" & vbCrlf
				tmpMSG = tmpMSG & "		<dl class=""hr-rows""><dt>科　室：</dt><dd>" & Trim(rsTmp("KSMC")) & "</dd></dl>" & vbCrlf
				tmpMSG = tmpMSG & "		<dl class=""hr-rows""><dt>职　称：</dt><dd>" & Trim(rsTmp("PRZC")) & "</dd></dl>" & vbCrlf
				tmpMSG = tmpMSG & "		<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf
				tmpMSG = tmpMSG & "		<dl class=""hr-rows""><dt>学　年：</dt><dd>" & Trim(rsTmp("scYear")) & "</dd></dl>" & vbCrlf
				tmpMSG = tmpMSG & "		<dl class=""hr-rows""><dt>学时数：</dt><dd>" & Trim(rsTmp("SumScore")) & "学时</dd></dl>" & vbCrlf
				tmpMSG = tmpMSG & "		<dl class=""hr-rows""><dt>业绩分：</dt><dd>" & Trim(rsTmp("TotalScore")) & "分</dd></dl>" & vbCrlf
				tmpMSG = tmpMSG & "		<dl class=""hr-rows""><dt>等　级：</dt><dd>" & Trim(rsTmp("Grade")) & "</dd></dl>" & vbCrlf
				tmpMSG = tmpMSG & "		<div class=""hr-gap-20 hr-fix""></div>" & vbCrlf

				Set rs = Conn.Execute(Replace(sqlTmp, "HR_KPI_SUM", "HR_KPI"))
				For i = 0 To Ubound(arrField)
					ValueNum = rsTmp(arrField(i))
					tFieldKey = Split(Replace(arrField(i), "F", ""), "_")
					tItemName = GetTypeName("HR_Class","ClassName","ClassID", HR_Clng(tFieldKey(0))) & "："

					If Ubound(tFieldKey) = 1 Then
						tItemName = tItemName & "<br>" & arrStuType(HR_Clng(tFieldKey(1))-1) & "　"
					End If
					If HR_CDbl(ValueNum) > 0 Then ValueNum = "<b class=""txtNum1"">" & FormatNumber(ValueNum, 2, -1) & "</b>" Else ValueNum="<b class=""txtNum0"">0</b>"
					If HR_CDbl(rs(arrField(i))) > 0 Then ValueNum1 = "<b class=""txtNum1"">" & FormatNumber(rs(arrField(i)), 2, -1) & "</b>" Else ValueNum1="<b class=""txtNum0"">0</b>"
					tmpMSG = tmpMSG & "		<dl class=""hr-rows""><dt>" & tItemName & "</dt><dd>学时 " & ValueNum & "<br>分值 " & ValueNum1 & "</dd></dl>" & vbCrlf
				Next
				Set rs = Nothing
				tmpMSG = tmpMSG & "	</div>" & vbCrlf
				tmpMSG = tmpMSG & "</div>" & vbCrlf
				

			End If
		Set rsTmp = Nothing
	End If
	Response.Write tmpMSG
End Sub

%>