<!--#include file="../Core/Lead.asp"-->
<!--#include file="../Core/incKernel.asp"-->
<!--#include file="../Core/incPublic.asp"-->
<!--#include file="../Core/incFront.asp"-->
<!--#include file="../Core/incVerify.asp"-->
<!--#include file="../Core/incWechat.asp"-->
<!--#include file="./incCommon.asp"-->
<!--#include file="./incPurview.asp"-->
<!--#include file="./m_ManageCommon.asp"-->
<%
Dim strParm : strParm = Trim(Request("Parm"))
Dim arrParm : arrParm = Split(strParm, "/")
Dim scriptCtrl : SiteTitle = "填写代课申请"
If ChkWechatTokenQY() = False Then Call GetWechatTokenQY()		'提前检查企业微信Token是否过期
If ChkTokenBobao() = False Then Call GetTokenBobao()			'检查信息播报Token

If IsNull(strParm) Or strParm = "" Then Call MainBody() : Response.End
If Ubound(arrParm) > 0 Then Action = Trim(ReplaceBadChar(arrParm(1)))

Select Case Action
	Case "Index" Call MainBody()
	Case "getItemJson" Call getItemJson()
	Case "Step2" Call Step2()
	Case "Step2Save" Call Step2Save()
	Case Else Response.Write GetErrBody(0)
End Select

Sub MainBody()
	Dim tid : tid = HR_Clng(Request("tid"))
	Dim cid : cid = HR_Clng(Request("cid"))
	Dim arrItem : arrItem = GetTableDataQuery("HR_Class", "", 1, "ClassID=" & tid & "")		'// 取考核项目信息
	If Hr_Clng(arrItem(0, 1)) = 0 Then
		ErrMsg = "考核项目不存在"
		Response.Write GetErrBody(0) : Response.End
	End If
	Dim tTable : tTable = "HR_Sheet_" & tid
	Dim arrCourse : arrCourse = GetTableDataQuery(tTable, "", 1, "ID=" & cid & "")		'// 取课程信息
	If Hr_Clng(arrCourse(0, 1)) = 0 Then
		ErrMsg = "课程不存在"
		Response.Write GetErrBody(0) : Response.End
	End If
	If UserYGDM = 0 Then
		ErrMsg = "您的登陆已经失效"
		Response.Write GetErrBody(0) : Response.End
	End If
	Dim oldTeachTime : oldTeachTime = FormatDate(ConvertNumDate(arrCourse(7, 1)),2)		'//取原上课时间
	Dim tmpID : tmpID = 0
	SiteTitle = "第三步：选择更改时间"
	tmpHtml = "<style type=""text/css"">" & vbCrlf
	tmpHtml = tmpHtml & "		.text-box {border-bottom:1px solid #ddd;}" & vbCrlf
	tmpHtml = tmpHtml & "		.text-box .title {border-bottom:1px solid #083F56;padding:10px;} .text-box .title h3 {font-size:1.2rem;}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-pop-btn {padding:10px; position:fixed; bottom:0px; width:100%; box-sizing:border-box; background:#eee; border-top:1px solid #083F56; z-index:10}" & vbCrlf
	tmpHtml = tmpHtml & "		.hr-pop-btn em {width:40%;} .weui-btn + .weui-btn {margin-top:0}" & vbCrlf
	tmpHtml = tmpHtml & "	</style>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"" src=""" & InstallDir & "Static/weui/lib/fastclick.js""></script>" & vbCrlf
	tmpHtml = tmpHtml & "	<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "		$(document).ready(function(){});" & vbCrlf
	tmpHtml = tmpHtml & "	</script>" & vbCrlf

	strHtml = getPageHead(1)
	strHtml = Replace(strHtml, "[@HeadStyle]", tmpHtml) : strHtml = Replace(strHtml, "[@HeadScript]", "")
	Response.Write ReplaceCommonLabel(strHtml)
	Response.Write getHeadNav(0)
	Response.Write "<div class=""hr-fix hr-header-hide""></div>" & vbCrlf
	Response.Write "<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "<form id=""EditForm"" name=""EditForm"" action="""">" & vbCrlf
	Response.Write "<div class=""weui-cells weui-cells_form"">" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">原上课时间：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""Teacher"" class=""weui-input"" id=""Teacher"" type=""text"" value=""" & oldTeachTime & """ data-key=""Teacher"" data-value=""TeacherID"" placeholder="""">" & vbCrlf
	Response.Write "			<input name=""TeacherID"" class=""weui-input"" id=""TeacherID"" type=""hidden"" value=""" & UserYGDM & """>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	'Response.Write "		<div class=""weui-cell__ft popWin"" data-id=""Teacher""><i class=""hr-icon"">&#xeeed;</i>选择</div>" & vbCrlf	'教师不用选
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""weui-cell"">" & vbCrlf
	Response.Write "		<div class=""weui-cell__hd""><label class=""weui-label"">新上课时间：</label></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell__bd"">" & vbCrlf
	Response.Write "			<input name=""NewTeachTime"" class=""weui-input"" id=""NewTeachTime"" type=""text"" value="""" placeholder=""点此选择新的时间"">" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf
	Response.Write "	<div class=""hr-gap-20""></div>" & vbCrlf
	Response.Write "	<div class=""text-box"">" & vbCrlf
	Response.Write "		<div class=""title""><h3>备注说明：</h3></div>" & vbCrlf
	Response.Write "		<div class=""weui-cell"">" & vbCrlf
	Response.Write "			<div class=""weui-cell__bd""><textarea class=""weui-textarea"" name=""Intro"" id=""Intro"" placeholder=""请输入内容"" rows=""5""></textarea></div>" & vbCrlf
	Response.Write "		</div>" & vbCrlf
	Response.Write "	</div>" & vbCrlf

	Response.Write "	<input name=""ItemID"" id=""ItemID"" type=""hidden"" value=""" & tid & """>" & vbCrlf
	Response.Write "	<input name=""CourseID"" id=""CourseID"" type=""hidden"" value=""" & cid & """>" & vbCrlf
	Response.Write "	<div class=""hr-rows hr-pop-btn""><em class=""weui-btn weui-btn_primary"" id=""SendForm"" data-id=""" & tmpID & """>提交</em><em class=""weui-btn weui-btn_warn navBack"">返回</em></div>" & vbCrlf
	Response.Write "</div>" & vbCrlf
	Response.Write "</form>" & vbCrlf
	tmpHtml = "<script type=""text/javascript"">" & vbCrlf
	tmpHtml = tmpHtml & "	$("".navBack"").on(""click"",function(){ location.href=""" & ParmPath & "Daike/Step2.html?tid=" & tid & "&cid=" & cid & """; });" & vbCrlf
	tmpHtml = tmpHtml & "	$(""#NewTeachTime"").calendar({dateFormat: 'yyyy-mm-dd'});" & vbCrlf

	tmpHtml = tmpHtml & "	$(""#SendForm"").on(""click"",function(){" & vbCrlf		'//提交表单
	tmpHtml = tmpHtml & "		console.log($(""#EditForm"").serialize())" & vbCrlf
	tmpHtml = tmpHtml & "		$.post(""" & ParmPath & "DaikeChangeTime/Step2Save.html"",$(""#EditForm"").serialize(), function(res){" & vbCrlf
	tmpHtml = tmpHtml & "			if(res.err){" & vbCrlf		'处理返回值的消息提示
	tmpHtml = tmpHtml & "				$.toast(res.errmsg, 'cancel');" & vbCrlf	
	tmpHtml = tmpHtml & "			}else{" & vbCrlf
	tmpHtml = tmpHtml & "				$.toast(res.errmsg);" & vbCrlf
	tmpHtml = tmpHtml & "			}" & vbCrlf
	tmpHtml = tmpHtml & "		},""json"");" & vbCrlf
	tmpHtml = tmpHtml & "	});" & vbCrlf

	tmpHtml = tmpHtml & "</script>" & vbCrlf
	strHtml = getPageFoot(1) : strHtml = Replace(strHtml, "[@FootScript]", tmpHtml)
	Response.Write strHtml
End Sub

Sub Step2Save()
	Dim tNow : tNow = Now() : ErrMsg = ""
	Dim cid : cid = Trim(Request.Form("CourseID"))	'获取课程ID
	Dim tid : tid = Trim(Request.Form("ItemID"))	'获取考核项目ID
	Dim notes : notes = Trim(Request.Form("Intro"))	'获取备注说明
	Dim new_teach_time : new_teach_time = Trim(Request.Form("NewTeachTime"))	'获取新的时间

	If HR_isNull(new_teach_time) Then ErrMsg = "请选择<br>新上课时间"
	If HR_isNull(notes) Then ErrMsg = "请输入备注说明"
	If Not(HR_isNull(ErrMsg)) Then
		strTmp = "{""err"":true, ""errcode"": 500, ""errmsg"":""" & ErrMsg & """, ""icon"":2}"
		Response.Write strTmp : Response.End
	End If

	strTmp = "{""err"":false, ""errcode"": 0, ""errmsg"":""更新时间成功"", ""icon"":1}"
	Response.Write strTmp

	'====== 发送提醒消息
	Dim msg_title : msg_title = UserYGXM & " 申请调换课程，需要您审核"
	Dim msg_content : msg_content = UserYGXM & "老师申请调换课程。<br>理由：" & tReason & "<br>申请时间：" & FormatDate(Now(), 1)
	Dim send_url : send_url = ""
	Call SentWechatMSG_QYCard(tPasserID, msg_title, send_url, msg_content)		'发送卡片消息给教研主任


End Sub
%>