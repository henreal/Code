<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./m_ManageCommon.asp"-->
<!--#include file="./m_ManageEvaluateCEX.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
SiteTitle = "评价"

Dim strHeadHtml : strHeadHtml =	ReplaceCommonLabel(getPageHead(1))
Dim strFootHtml : strFootHtml =	ReplaceCommonLabel(getPageFoot(1))

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))
Select Case Action
	Case "Index" Call MainBody()
	Case "TeachQuality" Call TeachQuality()
	Case "EditQuality" Call EditQuality()
	Case "SaveQuality" Call SaveQuality()
	Case "getItemCourse" Call getItemCourse()
	Case "CEX" Call CEX()
	Case "ApplyModify" Call ApplyModify()
	Case "AgreeApply" Call AgreeApply()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		body {background-color:#fff;}" & vbCrlf
	strHtml = strHtml & "		.navExtend {height: initial;flex-grow:2;text-align:right;}" & vbCrlf
	strHtml = strHtml & "		.navExtend span {font-size:1.2rem;display:line-block;background-color:#f7ce93;padding:2px 3px;color:#035;border-radius: 2px}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf
	
	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageEvaluate/TeachQuality.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xe9dc;</i></div><div class=""weui-cell__bd weui-cell_primary viewMSG"" data-id=""1""><p>课堂教学质量评价</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Response.Write "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageEvaluate/CEX.html"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xe9dc;</i></div><div class=""weui-cell__bd weui-cell_primary viewMSG"" data-id=""2""><p>mini-CEX<sup>plus</sup>记录</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft""></div>" & vbCrlf
	Response.Write "	</a>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "/Index.html""; });" & vbCrlf
	strHtml = strHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var content = $("".weui-textarea"").val();" & vbCrlf
	strHtml = strHtml & "		if(content==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""内容太少"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else{" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Propose/SavePost.html"", {Content:content}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(rsStr.reMessge, function(){ $.closePopup();location.reload(); });" & vbCrlf
	strHtml = strHtml & "			},""json"");" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "		" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub TeachQuality()
	SiteTitle = "课堂教学质量评价"
	Dim rsList, strList, dataTable, tParm, tmpID
	If Ubound(arrParm) > 1 Then
		tParm = Trim(arrParm(2))
		tmpID = HR_Clng(Request("ID"))
		Select Case tParm
			Case "ViewQuality" Call ViewQuality()
			Case "SaveQuality" Call SaveQuality()
		End Select
		Exit Sub
	End If
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	strHtml = strHtml & "		.popbtn {position:fixed;bottom:0px;left:0;right:0;padding:10px;z-index:10;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Set rsList = Conn.Execute("Select * From HR_Evaluate Where ParticipantID>0")
		If Not(rsList.BOF And rsList.EOF) Then
			Do While Not rsList.EOF
				dataTable = "HR_Sheet_" & rsList("ItemID")
				strList = strList & "	<a class=""weui-cell weui-cell_access"" href=""" & ParmPath & "ManageEvaluate/TeachQuality/ViewQuality.html?ID=" & rsList("ID") & """>" & vbCrlf
				strList = strList & "		<div class=""weui-cell__hd iconTit""><i class=""hr-icon"">&#xead1;</i></div><div class=""weui-cell__bd weui-cell_primary viewMSG"" data-id=""1""><p>" & rsList("Teacher") & " " & rsList("Course") & "<br>评价人：" & Trim(rsList("Participant")) & "　" & FormatDate(rsList("CreateTime"),10) & "</p></div>" & vbCrlf
				strList = strList & "		<div class=""weui-cell__ft""></div>" & vbCrlf
				strList = strList & "	</a>" & vbCrlf
				rsList.MoveNext
			Loop
		Else
			strList = "<div class=""hr-noinfo""><h2><i class=""hr-icon"">&#xef7b;</i></h2><h3>暂时还没有教师发表评价！</h3></div>" & vbCrlf
		End If
	Set rsList = Nothing

	Response.Write "<div class=""weui-cells__title""></div>" & vbCrlf
	Response.Write "<div class=""weui-cells"">" & vbCrlf
	Response.Write " " & strList
	Response.Write "</div>" & vbCrlf

	Response.Write "<div class=""hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""popbtn"">" & vbCrlf
	'Response.Write "	<em class=""pass""><a href=""" & ParmPath & "ManageEvaluate/ApplyModify.html"" class=""weui-btn weui-btn_primary"" id=""subApply"">查看修改申请</a></em>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	strHtml = strHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var content = $("".weui-textarea"").val();" & vbCrlf
	strHtml = strHtml & "		if(content==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""内容太少"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else{" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Propose/SavePost.html"", {Content:content}, function(rsStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(rsStr.reMessge, function(){ $.closePopup();location.reload(); });" & vbCrlf
	strHtml = strHtml & "			},""json"");" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "		" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub EditQuality()
	SiteTitle = "课堂教学质量评价"
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-swap-items dl {align-items:stretch;border-bottom:1px solid #eee;padding:5px 0;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .weui-cell__ft {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.weui-cells_form .popWin i {font-size:1.2rem;position:relative;top:2px;color:#4ce}" & vbCrlf
	strHtml = strHtml & "		.weui-toast {margin-left: auto;} .weui-textarea{font-size:1rem}" & vbCrlf
	strHtml = strHtml & "		.hr-tips {border-top:1px solid #eee;padding:5px} .hr-tips .tipsIcon {color:#f30;padding-right:5px;font-size:1.5rem;}" & vbCrlf
	strHtml = strHtml & "		.title {border-top:1px solid #eee;padding:10px} .title h3 {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.remarkbar {padding:5px 10px;border-top:1px solid #eee;color:#444;background-color:#f2f2f2}" & vbCrlf
	strHtml = strHtml & "		.weui-count .weui-count__number {font-size:1.1rem;}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课教师：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Teacher"" class=""weui-input"" id=""Teacher"" type=""text"" value="""" data-key=""Teacher"" data-value=""TeacherID"" placeholder="""">" & vbCrlf
	Response.Write "			<input name=""TeacherID"" class=""weui-input"" id=""TeacherID"" type=""hidden"" value="""">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft popWin"" data-id=""Teacher""><i class=""hr-icon"">&#xeeed;</i>选择</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择项目：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Item"" class=""weui-input opt1"" id=""Item"" type=""text"" value="""" readonly>" & vbCrlf
	Response.Write "			<input name=""ItemID"" class=""weui-input"" id=""ItemID"" type=""hidden"" value=""0"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">选择课程：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Course"" class=""weui-input opt2"" id=""Course"" type=""text"" value="""" readonly>" & vbCrlf
	Response.Write "			<input name=""CourseID"" class=""weui-input"" id=""CourseID"" type=""hidden"" value=""0"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">开课学院：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Campus"" class=""weui-input opt1"" id=""Campus"" type=""text"" value="""" readonly>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""hr-rows hr-tips"">" & vbCrlf
	Response.Write "		<em class=""tipsIcon""><i class=""hr-icon"">&#xf06a;</i></em>" & vbCrlf
	Response.Write "		<em class=""hr-row-fill tipstxt"">评价标准：>9优/9-6良/<6欠佳</em>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""title""><h3>教学态度与基本技能</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、要求脱稿讲授，语言准确流畅，逻辑性强，富感染力，语速、语调适宜、抑扬顿挫。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score1"" id=""Score1"" class=""weui-count__number Score"" type=""number"" value=""1"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark1"" id=""Remark1"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、精神饱满，教态大方，仪表端正。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score2"" id=""Score2"" class=""weui-count__number Score"" type=""number"" value=""1"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark2"" id=""Remark2"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、PPT设计科学，板书工整，教案讲稿规范。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score3"" id=""Score3"" class=""weui-count__number Score"" type=""number"" value=""1"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark3"" id=""Remark3"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>教学设计与方法</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、运用先进教学理念、方法进行教学，三维目标明确，学情清楚，因材施教，循循善诱。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score4"" id=""Score4"" class=""weui-count__number Score"" type=""number"" value=""1"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark4"" id=""Remark4"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、教学设计科学，新课导入、知识教授、总结巩固、课外自主学习等教学环节设计合理。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score5"" id=""Score5"" class=""weui-count__number Score"" type=""number"" value=""1"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark5"" id=""Remark5"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、广泛使用多媒体、互联网等现代化教学手段进行辅助教学。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score6"" id=""Score6"" class=""weui-count__number Score"" type=""number"" value=""1"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark6"" id=""Remark6"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>教学内容</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">符合教学大纲（或课程标准）要求，授课内容正确，重点难点突出，深度与广度适宜，联系实际，例证恰当，适当关注学科进展。（10分） </em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score7"" id=""Score7"" class=""weui-count__number Score"" type=""number"" value=""1"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark7"" id=""Remark7"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>教学效果</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、课堂驾驭能力强，师生互动性、课堂纪律、学习气氛好。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score8"" id=""Score8"" class=""weui-count__number Score"" type=""number"" value=""1"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark8"" id=""Remark8"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、完成教学任务，实现教学目的，学生反馈教学效果好。（10分）</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score9"" id=""Score9"" class=""weui-count__number Score"" type=""number"" value=""1"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark9"" id=""Remark9"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>整体评价</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">整体评价（10分） </em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>评分：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><a class=""weui-count__btn weui-count__decrease""></a><input name=""Score10"" id=""Score10"" class=""weui-count__number Score"" type=""number"" value=""1"" /><a class=""weui-count__btn weui-count__increase""></a></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Remark10"" id=""Remark10"" placeholder=""请输入备注"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><p>总评得分（100分）：</p></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__ft"">" & vbCrlf
	Response.Write "			<div class=""weui-count""><input name=""TotalScore"" id=""TotalScore"" class=""weui-count__number"" type=""number"" value=""0"" readonly /></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>优点</h3></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Merit"" id=""Merit"" placeholder=""请输入优点"" rows=""2""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""title""><h3>问题与建议</h3></div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、教学</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Suggest1"" id=""Suggest1"" placeholder=""请输入内容"" rows=""3""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、学风</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Suggest2"" id=""Suggest2"" placeholder=""请输入内容"" rows=""3""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、硬件</em></div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Suggest3"" id=""Suggest3"" placeholder=""请输入内容"" rows=""3""></textarea></div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""weui-btn-area""><em class=""weui-btn weui-btn_primary"" id=""subPost"">提交评价</em></div>" & vbCrlf
	Response.Write "</form>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "<div class=""hr-header-hide""></div>" & vbCrlf
	Response.Write "<div id=""full"" class=""hr-popup"">" & vbCrlf
	Response.Write "	<iframe src=""about:bank"" name=""listFrame"" id=""listFrame"" title=""ListFrame"" width=""100%"" height=""100%"" frameborder=""0""></iframe>" & vbCrlf
	Response.Write "</div>" & vbCrlf

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf

	strHtml = strHtml & "	$("".popWin"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		$(""#full"").show();var obj=$(this).data(""id"");" & vbCrlf
	strHtml = strHtml & "		$(""#listFrame"").attr(""src"",""" & ParmPath & "Directories/SelectTeacher.html?Type=3&reObjTxt="" + $(""#""+obj).data(""key"") + ""&reObjValue="" +  $(""#""+obj).data(""value""));" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	$(""#Course"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择"",items:[]," & vbCrlf
	strHtml = strHtml & "		onClose:function(){" & vbCrlf
	strHtml = strHtml & "			var tid = $(""#Course"").data(""values"");" & vbCrlf
	strHtml = strHtml & "			$(""#CourseID"").val(tid);" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf

	strHtml = strHtml & "	var arrItem =[" & GetSelectOptionItem() & "];" & vbCrlf
	strHtml = strHtml & "	$(""#Item"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择"",items:arrItem," & vbCrlf
	strHtml = strHtml & "		onClose:function(){" & vbCrlf
	strHtml = strHtml & "			var tid = $(""#Item"").data(""values""), teacher=$(""#TeacherID"").val();console.log(tid);" & vbCrlf
	strHtml = strHtml & "			$(""#ItemID"").val(tid);" & vbCrlf
	strHtml = strHtml & "			$.get(""" & ParmPath & "Evaluate/getItemCourse.html"",{ItemID:tid,TeacherID:teacher}, function(strForm){" & vbCrlf
	strHtml = strHtml & "				var reData = eval(""("" + strForm + "")"");" & vbCrlf
	strHtml = strHtml & "				$(""#Course"").select(""update"", reData);" & vbCrlf
	strHtml = strHtml & "				$(""#Course"").val("""");" & vbCrlf
	strHtml = strHtml & "			});" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	
	strHtml = strHtml & "	var arrCampus =[" & GetCampusArrData("", 0) & "];" & vbCrlf
	strHtml = strHtml & "	$(""#Campus"").select({" & vbCrlf
	strHtml = strHtml & "		title: ""请选择院区"",items:arrCampus," & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf


	strHtml = strHtml & "	var maxNum = 10, minNum = 1;" & vbCrlf
	strHtml = strHtml & "	$('.weui-count__decrease').click(function (e) {" & vbCrlf
	strHtml = strHtml & "		var $input = $(e.currentTarget).parent().find('.weui-count__number');" & vbCrlf
	strHtml = strHtml & "		var number = parseInt($input.val() || ""0"") - 1" & vbCrlf
	strHtml = strHtml & "		if (number < minNum) number = minNum;" & vbCrlf
	strHtml = strHtml & "		$input.val(number);" & vbCrlf
	strHtml = strHtml & "		CountTotalScore(number);" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf
	strHtml = strHtml & "	$('.weui-count__increase').click(function (e) {" & vbCrlf
	strHtml = strHtml & "		var $input = $(e.currentTarget).parent().find('.weui-count__number');" & vbCrlf
	strHtml = strHtml & "		var number = parseInt($input.val() || ""0"") + 1" & vbCrlf
	strHtml = strHtml & "		if (number > maxNum) number = maxNum;" & vbCrlf
	strHtml = strHtml & "		$input.val(number);" & vbCrlf
	strHtml = strHtml & "		CountTotalScore(number);" & vbCrlf
	strHtml = strHtml & "	})" & vbCrlf

	strHtml = strHtml & "	$(""#subPost"").on(""click"",function(){" & vbCrlf
	strHtml = strHtml & "		var TeacherID = $(""#TeacherID"").val(), ItemID = $(""#ItemID"").val(), CourseID = parseInt($(""#CourseID"").val()), Campus = $(""#Campus"").val();" & vbCrlf
	strHtml = strHtml & "		if(TeacherID==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""请选择授课教师"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else if(ItemID==""""){" & vbCrlf
	strHtml = strHtml & "			$.toast(""请选择项目"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else if(CourseID==0){" & vbCrlf
	strHtml = strHtml & "			$.toast(""请选择课程"", ""cancel"", function(){ return false; });" & vbCrlf
	strHtml = strHtml & "		}else{" & vbCrlf
	strHtml = strHtml & "			$.post(""" & ParmPath & "Evaluate/SaveQuality.html"", $(""#EditForm"").serialize(), function(rsStr){" & vbCrlf
	strHtml = strHtml & "				$.toast(rsStr.reMessge, function(){ $.closePopup();location.href=""" & ParmPath & "/Evaluate/TeachQuality.html""; });" & vbCrlf
	strHtml = strHtml & "			},""json"");" & vbCrlf
	strHtml = strHtml & "		}" & vbCrlf
	strHtml = strHtml & "		" & vbCrlf
	strHtml = strHtml & "	});" & vbCrlf
	strHtml = strHtml & "	function CountTotalScore(score){" & vbCrlf
	strHtml = strHtml & "		var total=0;" & vbCrlf
	strHtml = strHtml & "		$("".Score"").each(function(){" & vbCrlf
	strHtml = strHtml & "			total = total + parseInt($(this).val());" & vbCrlf
	strHtml = strHtml & "		});" & vbCrlf
	strHtml = strHtml & "		$(""#TotalScore"").val(total);" & vbCrlf
	strHtml = strHtml & "	}" & vbCrlf
	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub SaveQuality()
	ErrMsg = ""
	Dim rsSave
	If HR_Clng(Request("TeacherID")) = 0 Then ErrMsg = "您没有选择授课教师"
	If HR_Clng(Request("ItemID")) = 0 Then ErrMsg = "您没有选择项目"
	If HR_Clng(Request("CourseID")) = 0 Then ErrMsg = "您没有选择课程"
	If HR_IsNull(ErrMsg) = False Then Response.Write "{""Return"":false,""Err"":500,""reMessge"":""" & ErrMsg & """}" : Exit Sub

	Set rsSave = Server.CreateObject("ADODB.RecordSet")
		rsSave.Open("Select * From HR_Evaluate"), Conn, 1, 3
		rsSave.AddNew
		rsSave("ID") = GetNewID("HR_Evaluate", "ID")
		rsSave("Participant") = UserYGXM
		rsSave("ParticipantID") = UserYGDM
		rsSave("ItemID") = HR_Clng(Request("ItemID"))
		rsSave("CourseID") = HR_Clng(Request("CourseID"))
		rsSave("Course") = Trim(Request("Course"))
		rsSave("Title") = "课堂教学质量评价"
		rsSave("TeacherID") = HR_Clng(Request("TeacherID"))
		rsSave("Teacher") = Trim(Request("Teacher"))
		
		rsSave("Campus") = Trim(Request("Campus"))
		rsSave("Score1") = HR_Clng(Request("Score1"))
		rsSave("Remark1") = Trim(Request("Remark1"))
		rsSave("Score2") = HR_Clng(Request("Score2"))
		rsSave("Remark2") = Trim(Request("Remark2"))
		rsSave("Score3") = HR_Clng(Request("Score3"))
		rsSave("Remark3") = Trim(Request("Remark3"))
		rsSave("Score4") = HR_Clng(Request("Score4"))
		rsSave("Remark4") = Trim(Request("Remark4"))
		rsSave("Score5") = HR_Clng(Request("Score5"))
		rsSave("Remark5") = Trim(Request("Remark5"))
		rsSave("Score6") = HR_Clng(Request("Score6"))
		rsSave("Remark6") = Trim(Request("Remark6"))
		rsSave("Score7") = HR_Clng(Request("Score7"))
		rsSave("Remark7") = Trim(Request("Remark7"))
		rsSave("Score8") = HR_Clng(Request("Score8"))
		rsSave("Remark8") = Trim(Request("Remark8"))
		rsSave("Score9") = HR_Clng(Request("Score9"))
		rsSave("Remark9") = Trim(Request("Remark9"))
		rsSave("Score10") = HR_Clng(Request("Score10"))
		rsSave("Remark10") = Trim(Request("Remark10"))
		rsSave("TotalScore") = HR_Clng(Request("TotalScore"))
		rsSave("Merit") = Trim(Request("Merit"))
		rsSave("Suggest1") = Trim(Request("Suggest1"))
		rsSave("Suggest2") = Trim(Request("Suggest2"))
		rsSave("Suggest3") = Trim(Request("Suggest3"))
		rsSave("CreateTime") = Now()
		rsSave.Update
	Set rsSave = Nothing
	ErrMsg = "评价已提交成功！"
	Response.Write "{""Return"":true,""Err"":0,""reMessge"":""" & ErrMsg & """}"
End Sub

Sub ViewQuality()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))

	SiteTitle = "课堂教学质量评价"
	strHtml = "<link type=""text/css"" href=""/Static/weui/css/jquery-weui.min.css?v=1.2.1"" rel=""stylesheet"" media=""all"">" & vbCrlf
	strHtml = strHtml & "<style type=""text/css"">" & vbCrlf
	strHtml = strHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	strHtml = strHtml & "		.weui-toast {margin-left: auto;} .weui-textarea{font-size:1rem}" & vbCrlf
	strHtml = strHtml & "		.hr-tips {border-top:1px solid #eee;padding:5px} .hr-tips .tipsIcon {color:#f30;padding-right:5px;font-size:1.5rem;}" & vbCrlf
	strHtml = strHtml & "		.title {border-top:1px solid #eee;padding:10px} .title h3 {font-size:1.2rem;}" & vbCrlf
	strHtml = strHtml & "		.remarkbar {padding:5px 10px;border-top:1px solid #eee;color:#444;background-color:#f2f2f2}" & vbCrlf
	strHtml = strHtml & "	</style>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@Page_Title]", SiteTitle)
	strHeadHtml = Replace(strHeadHtml, "[@HeadStyle]", strHtml)

	strHtml = "<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	strHtml = strHtml & "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "		$(document).ready(function(){});" & vbCrlf
	strHtml = strHtml & "	</script>" & vbCrlf
	strHeadHtml = Replace(strHeadHtml, "[@HeadScript]", strHtml)
	Response.Write strHeadHtml
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf

	Dim dataTable, tItemName, tCourse, tStuClass, tContents, tAddress, tClassTime, tPeriod
	Set rsTmp = Conn.Execute("Select * From HR_Evaluate Where ID=" & tmpID)
		If Not(rsTmp.BOF And rsTmp.EOF) Then
			dataTable = "HR_Sheet_" & rsTmp("ItemID")
			Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">总评得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("TotalScore") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">评价人：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Participant") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">评价时间：" & FormatDate(rsTmp("CreateTime"), 10) & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课教师：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Teacher") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">项目：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & tItemName & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf

			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">课程：</label></div><div class=""weui-cell__bd"">" & rsTmp("Course") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">课程内容：</label></div><div class=""weui-cell__bd"">" & rsTmp("Contents") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课对象：</label></div><div class=""weui-cell__bd"">" & rsTmp("StuClass") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">授课时间：</label></div><div class=""weui-cell__bd"">" & rsTmp("ClassTime") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">开课学院：</label></div><div class=""weui-cell__bd"">" & rsTmp("Campus") & "</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			Response.Write "	<div class=""hr-rows hr-tips"">" & vbCrlf
			Response.Write "		<em class=""tipsIcon""><i class=""hr-icon"">&#xf06a;</i></em>" & vbCrlf
			Response.Write "		<em class=""hr-row-fill tipstxt"">评价标准：>9优/9-6良/<6欠佳</em>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""title""><h3>教学态度与基本技能</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、要求脱稿讲授，语言准确流畅，逻辑性强，富感染力，语速、语调适宜、抑扬顿挫。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score1") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark1") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、精神饱满，教态大方，仪表端正。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score2") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark2") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、PPT设计科学，板书工整，教案讲稿规范。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score3") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark3") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>教学设计与方法</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、运用先进教学理念、方法进行教学，三维目标明确，学情清楚，因材施教，循循善诱。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score4") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark4") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、教学设计科学，新课导入、知识教授、总结巩固、课外自主学习等教学环节设计合理。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score5") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark5") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、广泛使用多媒体、互联网等现代化教学手段进行辅助教学。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score6") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark6") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>教学内容</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">符合教学大纲（或课程标准）要求，授课内容正确，重点难点突出，深度与广度适宜，联系实际，例证恰当，适当关注学科进展。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score7") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark7") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>教学效果</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、课堂驾驭能力强，师生互动性、课堂纪律、学习气氛好。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score8") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark8") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、完成教学任务，实现教学目的，学生反馈教学效果好。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score9") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark9") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>整体评价</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">整体评价。（10分）</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell"">" & vbCrlf
			Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">得分：</label></div>" & vbCrlf
			Response.Write "		<div class=""weui-cell__bd"">" & rsTmp("Score10") & " 分</div>" & vbCrlf
			Response.Write "	</div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Remark10") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
			

			Response.Write "	<div class=""title""><h3>优点</h3></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Merit") & "</div></div>" & vbCrlf

			Response.Write "	<div class=""title""><h3>问题与建议</h3></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">1、教学</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Suggest1") & "</div></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">2、学风</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Suggest2") & "</div></div>" & vbCrlf
			Response.Write "	<div class=""remarkbar""><em class=""remarktxt"">3、硬件</em></div>" & vbCrlf
			Response.Write "	<div class=""weui-cell""><div class=""weui-cell__bd"">" & rsTmp("Suggest3") & "</div></div>" & vbCrlf

			
			Response.Write "</div>" & vbCrlf
		End If

	strHtml = "<script type=""text/javascript"">" & vbCrlf
	strHtml = strHtml & "	$("".navMenu em"").on(""click"",function(){ $("".layerNav"").toggle(); });" & vbCrlf
	strHtml = strHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf

	strHtml = strHtml & "</script>" & vbCrlf
	strFootHtml = Replace(strFootHtml, "[@FootScript]", strHtml)
	Response.Write strFootHtml
End Sub

Sub ApplyModify()
	Dim tPasser
	SiteTitle = "修改申请"
	tmpHtml = vbCrlf & "	<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-header {background-color:#061;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-item-panel {border-bottom:15px solid #eee;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-item-href {padding:10px;border-bottom:1px solid #ddd;padding-bottom:4px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-item-href em {color:#09c;font-size:1.3rem;} .hr-item-href tt {color:#999;font-size:1.3rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-item-panel dl {display: flex;padding:5px 0}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-item-panel dt {width:6rem;text-align:right;color:#999;flex-shrink:0}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-item-foot {border-top:1px solid #eee;padding:5px;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-item-foot .data {color:#999} .hr-item-foot .data i {color:#f90}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-item-foot .agree {color:#fff;background-color:#b446e4;padding:3px 10px;border-radius: 5px;}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml)
	strHtml = Replace(strHtml, "[@HeadScript]", "")
	strHtml = strHtml & "<header class=""hr-rows hr-header"">" & vbCrlf
	strHtml = strHtml & "	<nav class=""navBack""><em><i class=""hr-icon"">&#xf320;</i></em></nav>" & vbCrlf
	strHtml = strHtml & "	<nav class=""navTitle""><span>" & SiteTitle & "</span></nav>" & vbCrlf
	strHtml = strHtml & "	<nav class=""navMenu""><em><i class=""hr-icon"">&#xf32a;</i></em></nav>" & vbCrlf
	strHtml = strHtml & "</header>" & vbCrlf
	strHtml = strHtml & "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	strHtml = strHtml & "<div class=""hr-items"">" & vbCrlf
	sql = "Select a.*,b.Teacher,b.Course,b.ClassTime,b.Campus,(Select YGXM From HR_Teacher Where CAST(YGDM AS Int)=a.YGDM) As Applyer From HR_Apply a Left Join HR_Evaluate b On a.RelateID=b.ID Where a.Module='Evaluate' And a.RelateID>0"
	sql = sql & " Order By a.CreateTime DESC"
	Set rs = Conn.Execute(sql)
		If Not(rs.BOF And rs.EOF) Then
			Do While Not rs.EOF
				tPasser = Trim(strGetTypeName("HR_Teacher", "YGXM", "YGDM", rs("Passer")))
				strHtml = strHtml & "	<div class=""hr-item-panel"">" & vbCrlf
				strHtml = strHtml & "		<a href=""" & ParmPath & "ManageEvaluate/TeachQuality/ViewQuality.html?ID=" & rs("RelateID") & """ class=""hr-rows hr-item-href"">" & vbCrlf
				strHtml = strHtml & "			<em>" & rs("Applyer") & "老师申请修改课堂教学评价</em><tt><i class=""hr-icon hr-icon-top"">&#xf31a;</i></tt>" & vbCrlf
				strHtml = strHtml & "		</a>" & vbCrlf
				strHtml = strHtml & "		<dl><dt>修改理由：</dt><dd>" & rs("Reason") & "</dd></dl>" & vbCrlf
				strHtml = strHtml & "		<dl><dt>授课教师：</dt><dd>" & rs("Teacher") & "</dd></dl>" & vbCrlf
				strHtml = strHtml & "		<dl><dt>课程名称：</dt><dd>" & rs("Course") & "</dd></dl>" & vbCrlf
				strHtml = strHtml & "		<dl><dt>开课学院：</dt><dd>" & Trim(rs("Campus")) & "</dd></dl>" & vbCrlf
				strHtml = strHtml & "		<h4 class=""hr-rows hr-item-foot""><em class=""data""><i class=""hr-icon hr-icon-top"">&#xeedb;</i>" & FormatDate(rs("CreateTime"), 10) & "</em>" & vbCrlf
				If HR_CBool(rs("Passed")) Then
					strHtml = strHtml & "			<em class=""passed"" data-id=""" & rs("ID") & """>已通过【" & tPasser & "】</em>" & vbCrlf
				Else
					strHtml = strHtml & "			<em class=""agree"" data-id=""" & rs("ID") & """>同意</em>" & vbCrlf
				End If
				strHtml = strHtml & "		</h4>" & vbCrlf
				strHtml = strHtml & "	</div>" & vbCrlf
				rs.MoveNext
			Loop
		Else
			strHtml = strHtml & "	<div class=""hr-item-nodata"">还没有教师申请修改评价</div>" & vbCrlf
		End If
	Set rs = Nothing
	strHtml = strHtml & "</div>" & vbCrlf
	strHtml = strHtml & getPageFoot(1)
	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navMenu em"").on(""click"",function(){ location.href=""" & ParmPath & "ManageEvaluate/TeachQuality.html"";  });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ history.back(); });" & vbCrlf
	tmpHtml = tmpHtml & "	$("".agree"").on(""click"",function(){" & vbCrlf
	tmpHtml = tmpHtml & "		var tid = $(this).data(""id"");" & vbCrlf
	tmpHtml = tmpHtml & "		$.post(""" & ParmPath & "ManageEvaluate/AgreeApply.html"",{ID:tid}, function(strForm){" & vbCrlf
	tmpHtml = tmpHtml & "			if(strForm.err){" & vbCrlf
	tmpHtml = tmpHtml & "				$.toast(strForm.errmsg, ""cancel"");" & vbCrlf
	tmpHtml = tmpHtml & "			}else{" & vbCrlf
	tmpHtml = tmpHtml & "				$.toast(strForm.errmsg, function(){ location.href=""" & ParmPath & "ManageEvaluate/TeachQuality.html""; });" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		},""json"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf
	tmpHtml = tmpHtml & "</script>" & vbCrlf

	strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write ReplaceCommonLabel(strHtml)
End Sub

Sub getItemCourse()
	Dim tItemID : tItemID = HR_Clng(Request("ItemID"))
	Dim tTeacherID : tTeacherID = HR_Clng(Request("TeacherID"))
	Dim tTableName : tTableName = "HR_Sheet_" & tItemID
	Dim strTmp
	strTmp = strTmp & "{""items"":["
	If ChkTable(tTableName) Then
		sql = "Select a.*,b.ClassName,b.Template From " & tTableName & " a Inner Join HR_Class b on a.ItemID=b.ClassID Where VA1=" & tTeacherID
		Set rs = Conn.Execute(sql)
			If Not(rs.BOF And rs.EOF) Then
				i = 0
				Do While Not rs.EOF
					If i>0 Then strTmp = strTmp & ","
					strTmp = strTmp & "{""title"":""" & FormatDate(ConvertNumDate(rs("VA4")), 4)
					If rs("Template") = "TempTableA" Then
						strTmp = strTmp & " " & rs("VA8") & "_" & rs("VA7") & "节"
					Else
						strTmp = strTmp & " " & rs("VA6")
					End If
					strTmp = strTmp & """,""value"":""" & rs("ID") & """}"
					rs.MoveNext
					i = i + 1
				Loop
			Else
				strTmp = strTmp & "{""title"":""该教师在本项目中没有课程"",""value"":""0""}"
			End If
		Set rs = Nothing
	End If
	strTmp = strTmp & "]}"
	Response.Write strTmp
End Sub
Sub AgreeApply()
	Dim tmpID : tmpID = HR_Clng(Request("ID"))
	Dim RelateID
	sql = "Select a.*,b.Teacher,b.Course,b.ClassTime,b.Campus From HR_Apply a Left Join HR_Evaluate b On a.RelateID=b.ID Where a.Module='Evaluate' And a.RelateID>0 And a.ID=" & tmpID
	sql = sql & " Order By a.CreateTime DESC"
	Set rs = Conn.Execute(sql)
		If Not(rs.BOF And rs.EOF) Then
			RelateID = HR_CLng(rs("RelateID"))
			If RelateID > 0 Then
				Conn.Execute("Update HR_Evaluate Set Passed=" & HR_False & " Where ID=" & RelateID)
				Conn.Execute("Update HR_Apply Set Passed=" & HR_True & ",Passer=" & UserYGDM & ",PassTime=GETDATE() Where ID=" & tmpID)
				Response.Write "{""err"":false,""errcode"":0,""errmsg"":""您已同意该老师的申请！"",""icon"":1}"
			End If
		Else
			Response.Write "{""err"":true,""errcode"":400,""errmsg"":""修改申请不存在或已删除！"",""icon"":2}"
		End If
	Set rs = Nothing
End Sub

Function GetCampusArrData(fCampus, fType)		'取校院区数据
	Dim strFun, iFun, fArrCampus : fArrCampus = Split(XmlText("Common", "Campus", ""), "|")
	For iFun = 0 To Ubound(fArrCampus)
		If iFun > 0 Then strFun = strFun & ","
		strFun = strFun & """" & fArrCampus(iFun) & """"
	Next
	GetCampusArrData = strFun
End Function

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